// SharePoint CRUD Service
import { SHAREPOINT_CONFIG } from './config/config.js';

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

    // Get request digest for POST/UPDATE/DELETE operations
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

    // Get current user information
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

    // Create a new list item
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

    // Update an existing list item
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

    // Delete a list item
    async deleteItem(itemId) {
        try {
            const digest = await this.getRequestDigest();
            
            const response = await fetch(`${this.apiUrl}lists/getbytitle('${this.listName}')/items(${itemId})`, {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-RequestDigest': digest,
                    'X-HTTP-Method': 'DELETE',
                    'IF-MATCH': '*'
                },
                credentials: 'include'
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
            }
            
            return { success: true, itemId };
        } catch (error) {
            console.error('Error deleting item:', error);
            throw error;
        }
    }

    // Get all items from the list
    async getItems(filter = '', orderBy = '') {
        try {
            let url = `${this.apiUrl}lists/getbytitle('${this.listName}')/items`;
            
            const params = [];
            if (filter) params.push(`$filter=${encodeURIComponent(filter)}`);
            if (orderBy) params.push(`$orderby=${encodeURIComponent(orderBy)}`);
            
            if (params.length > 0) {
                url += '?' + params.join('&');
            }
            
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
            console.error('Error getting items:', error);
            throw error;
        }
    }

    // Get a single item by ID
    async getItemById(itemId) {
        try {
            const response = await fetch(`${this.apiUrl}lists/getbytitle('${this.listName}')/items(${itemId})`, {
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
            return data.d;
        } catch (error) {
            console.error('Error getting item:', error);
            throw error;
        }
    }

    // Transform case data to SharePoint format
    transformCaseToSharePoint(caseData) {
        return {
            Title: caseData.zaaknummer || '',
            Feitcode: caseData.feitcode || '',
            CJIBNummer: caseData.cjibNummer || '', // CJIB Number field - must exist in SharePoint
            // CJIBLast4: caseData.cjibLast4 || '',   // Excluded: Display-only field (last 4 digits of CJIBNummer)
            Feitomschrijving: caseData.feitomschrijving || '',
            Vooronderzoek: caseData.vooronderzoek || '',
            ReactiePMBU: caseData.reactie || '',
            HearingDate: caseData.hearingDate || null,
            StartTime: caseData.startTime || '',
            EndTime: caseData.endTime || '',
            Status: caseData.status || 'Bezig met uitwerken'
        };
    }

    // Transform SharePoint data to case format
    transformSharePointToCase(spData) {
        return {
            id: `case-${spData.Id}`,
            sharePointId: spData.Id,
            zaaknummer: spData.Title || '',
            feitcode: spData.Feitcode || '',
            feitomschrijving: spData.Feitomschrijving || '',
            vooronderzoek: spData.Vooronderzoek || '',
            reactie: spData.ReactiePMBU || '',
            hearingDate: spData.HearingDate || '',
            startTime: spData.StartTime || '',
            endTime: spData.EndTime || '',
            status: spData.Status || 'Bezig met uitwerken',
            username: spData.Username || '',
            isModified: false
        };
    }
}

// Export singleton instance
export const sharePointService = new SharePointService();
