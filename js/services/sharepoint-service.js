import { SHAREPOINT_CONFIG } from '../config/config.js';

class SharePointService {
    constructor() {
        this.siteUrl = SHAREPOINT_CONFIG.siteUrl;
        this.listName = SHAREPOINT_CONFIG.listName;
        this.apiUrl = SHAREPOINT_CONFIG.apiUrl;
        this.contextApiUrl = SHAREPOINT_CONFIG.contextApiUrl;
        this.currentUser = null;
        
        // Debug logging
        console.log('SharePoint Service Configuration:');
        console.log('- Site URL:', this.siteUrl);
        console.log('- List Name:', this.listName);
        console.log('- API URL:', this.apiUrl);
        console.log('- Context API URL:', this.contextApiUrl);
    }

    async getRequestDigest() {
        try {
            const contextUrl = `${this.contextApiUrl}contextinfo`;
            console.log('Attempting to fetch contextinfo from:', contextUrl);
            
            const response = await fetch(contextUrl, {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose'
                },
                credentials: 'include'
            });
            
            if (!response.ok) {
                console.error('Context API response status:', response.status, 'URL:', contextUrl);
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const data = await response.json();
            return data.d.GetContextWebInformation.FormDigestValue;
        } catch (error) {
            console.error('Error getting request digest:', error);
            throw error;
        }
    }

    async getCurrentUser() {
        try {
            const response = await fetch(`${this.apiUrl}currentuser`, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                },
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const data = await response.json();
            this.currentUser = data.d;
            return data.d;
        } catch (error) {
            console.error('Error getting current user:', error);
            throw error;
        }
    }

    // Test SharePoint connection
    async testConnection() {
        try {
            // First test basic web access
            const webResponse = await fetch(`${this.apiUrl}`, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                },
                credentials: 'include'
            });
            
            if (!webResponse.ok) {
                throw new Error(`Cannot access SharePoint web: ${webResponse.status}`);
            }

            // Then test list access
            const listResponse = await fetch(`${this.apiUrl}lists/getbytitle('${this.listName}')`, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                },
                credentials: 'include'
            });
            
            if (!listResponse.ok) {
                throw new Error(`Cannot access list '${this.listName}': ${listResponse.status}`);
            }

            return { success: true, message: 'SharePoint connection successful' };
        } catch (error) {
            console.error('SharePoint connection test failed:', error);
            throw error;
        }
    }

    async createItem(itemData) {
        try {
            const digest = await this.getRequestDigest();
            const currentUser = await this.getCurrentUser();
            
            // Add required SharePoint metadata type
            const itemWithMetadata = {
                __metadata: {
                    type: `SP.Data.${this.listName}ListItem`
                },
                ...itemData,
                Username: currentUser.Title || currentUser.LoginName
            };

            console.log('Creating item with data:', itemWithMetadata);

            const response = await fetch(`${this.apiUrl}lists/getbytitle('${this.listName}')/items`, {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': digest
                },
                credentials: 'include',
                body: JSON.stringify(itemWithMetadata)
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                console.error('Create item error response:', errorText);
                throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
            }
            
            const data = await response.json();
            return data.d;
        } catch (error) {
            console.error('Error creating item:', error);
            throw error;
        }
    }

    async updateItem(itemId, itemData) {
        try {
            const digest = await this.getRequestDigest();
            
            // Add required SharePoint metadata type for updates
            const itemWithMetadata = {
                __metadata: {
                    type: `SP.Data.${this.listName}ListItem`
                },
                ...itemData
            };

            console.log('Updating item with data:', itemWithMetadata);
            
            const response = await fetch(`${this.apiUrl}lists/getbytitle('${this.listName}')/items(${itemId})`, {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': digest,
                    'X-HTTP-Method': 'MERGE',
                    'IF-MATCH': '*'
                },
                credentials: 'include',
                body: JSON.stringify(itemWithMetadata)
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                console.error('Update item error response:', errorText);
                throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
            }
            
            return { success: true, itemId };
        } catch (error) {
            console.error('Error updating item:', error);
            throw error;
        }
    }

    async getAvailableDatesWithIncompleteCases() {
        try {
            // Query SharePoint for items where Status is not 'Afgehandeld'
            const filter = "Status ne 'Afgehandeld'";
            const select = "HearingDate,Status,Id";
            const url = `${this.apiUrl}lists/getbytitle('${this.listName}')/items?$filter=${encodeURIComponent(filter)}&$select=${select}&$orderby=HearingDate desc`;
            
            console.log('Fetching available dates from:', url);
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                },
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const data = await response.json();
            const items = data.d.results;
            
            // Group by hearing date and count items
            const dateGroups = {};
            items.forEach(item => {
                if (item.HearingDate) {
                    // Parse SharePoint date and convert to ISO date string (YYYY-MM-DD)
                    const date = new Date(item.HearingDate).toISOString().split('T')[0];
                    if (!dateGroups[date]) {
                        dateGroups[date] = [];
                    }
                    dateGroups[date].push(item);
                }
            });
            
            // Convert to array with date info
            const availableDates = Object.keys(dateGroups)
                .map(date => ({
                    date: date, // ISO format YYYY-MM-DD
                    displayDate: new Date(date + 'T12:00:00.000Z').toLocaleDateString('nl-NL'), // Add time to avoid timezone issues
                    count: dateGroups[date].length,
                    items: dateGroups[date]
                }))
                .sort((a, b) => new Date(b.date + 'T12:00:00.000Z') - new Date(a.date + 'T12:00:00.000Z')); // Most recent first
            
            return availableDates;
        } catch (error) {
            console.error('Error fetching available dates:', error);
            throw error;
        }
    }

    async getCasesByDate(targetDate) {
        try {
            // Ensure targetDate is in ISO format and create proper date range
            const startDate = new Date(targetDate + 'T00:00:00.000Z');
            const endDate = new Date(targetDate + 'T23:59:59.999Z');
            
            // Use ISO string format for SharePoint datetime filtering
            const filter = `HearingDate ge datetime'${startDate.toISOString()}' and HearingDate le datetime'${endDate.toISOString()}' and Status ne 'Afgehandeld'`;
            const url = `${this.apiUrl}lists/getbytitle('${this.listName}')/items?$filter=${encodeURIComponent(filter)}&$orderby=StartTime asc`;
            
            console.log('Fetching cases for date:', targetDate, 'ISO range:', startDate.toISOString(), 'to', endDate.toISOString());
            console.log('Query URL:', url);
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                },
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const data = await response.json();
            return data.d.results;
        } catch (error) {
            console.error('Error fetching cases by date:', error);
            throw error;
        }
    }

    async getCaseByZaaknummer(zaaknummer) {
        try {
            if (!zaaknummer || zaaknummer.trim() === '') {
                return null;
            }
            
            // Search for existing case by Title (Zaaknummer)
            const filter = `Title eq '${zaaknummer.replace(/'/g, "''")}'`; // Escape single quotes
            const url = `${this.apiUrl}lists/getbytitle('${this.listName}')/items?$filter=${encodeURIComponent(filter)}&$top=1`;
            
            console.log('Checking for existing case with Zaaknummer:', zaaknummer);
            console.log('Query URL:', url);
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                },
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const data = await response.json();
            const results = data.d.results;
            
            if (results && results.length > 0) {
                console.log('Found existing case:', results[0]);
                return results[0];
            }
            
            return null;
        } catch (error) {
            console.error('Error checking for existing case:', error);
            throw error;
        }
    }

    transformCaseToSharePoint(caseData) {
        // Helper function to format date to ISO string for SharePoint
        const formatDateForSharePoint = (dateStr) => {
            if (!dateStr) return null;
            try {
                const date = new Date(dateStr);
                if (isNaN(date.getTime())) return null;
                return date.toISOString();
            } catch (error) {
                console.warn('Invalid date format:', dateStr);
                return null;
            }
        };

        return {
            Title: caseData.zaaknummer || '',
            Feitcode: caseData.feitcode || '',
            CJIBNummer: caseData.cjibNummer || '', // CJIB Number field - must exist in SharePoint
            // CJIBLast4: caseData.cjibLast4 || '',   // Excluded: Display-only field (last 4 digits of CJIBNummer)
            Betrokkene: caseData.betrokkene || '',
            Eigenaar: caseData.eigenaar || '',
            Soort: caseData.soort || '',
            AantekeningHoorverzoek: caseData.aantekeninghoorverzoek || '',
            Feitomschrijving: caseData.feitomschrijving || '',
            Vooronderzoek: caseData.vooronderzoek || '',
            ReactiePMBU: caseData.reactie || '',
            HearingDate: formatDateForSharePoint(caseData.hearingDate),
            StartTime: caseData.startTime || '',
            EndTime: caseData.endTime || '',
            Verslaglegger: caseData.verslaglegger || '',
            GesprokenMet: caseData.gesprokenMet || '',
            Bedrijfsnaam: caseData.bedrijfsnaam || '',
            Status: caseData.status || 'Bezig met uitwerken'
        };
    }

    // Method to lookup Feitomschrijving based on Feitcode
    async getFeitomschrijvingByFeitcode(feitcode) {
        if (!feitcode || feitcode.trim() === '') {
            return '';
        }

        try {
            const lookupConfig = SHAREPOINT_CONFIG.feitcodeLookup;
            const filter = `Feitcode eq '${feitcode.trim()}'`;
            const apiUrl = `${lookupConfig.apiUrl}lists/getbytitle('${lookupConfig.listName}')/items?$filter=${encodeURIComponent(filter)}&$select=Feitcode,Feitomschrijving`;
            
            console.log('Fetching Feitomschrijving from:', apiUrl);
            
            const response = await fetch(apiUrl, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                },
                credentials: 'include'
            });
            
            if (!response.ok) {
                console.warn(`HTTP error when fetching Feitomschrijving: ${response.status}`);
                return '';
            }
            
            const data = await response.json();
            
            if (data.d && data.d.results && data.d.results.length > 0) {
                const result = data.d.results[0];
                console.log('Found Feitomschrijving:', result.Feitomschrijving);
                return result.Feitomschrijving || '';
            } else {
                console.log(`No Feitomschrijving found for Feitcode: ${feitcode}`);
                return '';
            }
            
        } catch (error) {
            console.error('Error fetching Feitomschrijving:', error);
            return '';
        }
    }
}

export { SharePointService };