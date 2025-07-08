<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hoorzitting Notulen Logger</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- Preact and HTM for JSX-less React -->
    <script type="module">
        // Corrected imports from a more reliable ES Module CDN (esm.sh)
        import { h, render, Component } from 'https://esm.sh/preact';
        import { useState, useCallback, useEffect } from 'https://esm.sh/preact/hooks';
        import htm from 'https://esm.sh/htm';

        // Initialize htm with Preact
        const html = htm.bind(h);

        // Import SharePoint service and config
        // Note: In production, these would be proper ES6 imports
        // For now, we'll define them inline since ES modules can't import relative paths in this context
        
        // SharePoint Configuration
        const SHAREPOINT_CONFIG = {
            siteUrl: 'https://som.org.om.local/sites/MulderT/T/',
            listName: 'PMREG',
            apiUrl: 'https://som.org.om.local/sites/MulderT/T/_api/web/',
            contextApiUrl: 'https://som.org.om.local/sites/MulderT/T/_api/', // Separate URL for contextinfo
            listUrl: 'https://som.org.om.local/sites/MulderT/T/PMREG/',
        };

        // Status choices (updated workflow)
        const STATUS_CHOICES = [
            'Nieuw',
            'Voorbereiding',
            'In behandeling',
            'Aangehouden',
            'Klaarzetten voor DocGen',
            'Afgehandeld'
        ];

        // SharePoint Service Class
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
        }

        const sharePointService = new SharePointService();

        // --- Helper function for consistent ISO date handling ---
        const ensureISODate = (dateInput) => {
            if (!dateInput) return '';
            try {
                const date = new Date(dateInput);
                if (isNaN(date.getTime())) return '';
                return date.toISOString().split('T')[0]; // Return YYYY-MM-DD format
            } catch (error) {
                console.warn('Invalid date input:', dateInput);
                return '';
            }
        };

        // --- Helper function to calculate end time (start time + 4 minutes) ---
        const calculateEndTime = (startTime) => {
            if (!startTime || !startTime.match(/^\d{1,2}:\d{2}$/)) {
                return '';
            }
            
            try {
                const [hours, minutes] = startTime.split(':').map(Number);
                
                // Validate input time
                if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
                    return '';
                }
                
                // Calculate end time (+4 minutes)
                const startMinutes = hours * 60 + minutes;
                const endMinutes = startMinutes + 4;
                
                // Handle day rollover (if time goes past 23:59)
                const finalMinutes = endMinutes % (24 * 60);
                const endHours = Math.floor(finalMinutes / 60);
                const endMins = finalMinutes % 60;
                
                return `${endHours.toString().padStart(2, '0')}:${endMins.toString().padStart(2, '0')}`;
            } catch (error) {
                console.warn('Error calculating end time:', error);
                return '';
            }
        };

        // --- Helper function to generate initial empty cases ---
        const createInitialCases = (count) => {
            return Array.from({ length: count }, (_, i) => ({
                id: `case-${i}`,
                sharePointId: null,
                zaaknummer: '',
                feitcode: '',
                cjibNummer: '',
                cjibLast4: '',
                betrokkene: '',
                eigenaar: '',
                soort: '',
                aantekeninghoorverzoek: '',
                feitomschrijving: '',
                vooronderzoek: '',
                reactie: '',
                hearingDate: ensureISODate(new Date()), // Today's date in ISO format
                startTime: '',
                endTime: '',
                verslaglegger: '',
                gesprokenMet: '',
                bedrijfsnaam: '',
                status: 'Nieuw', // Default status for new cases
                isModified: false,
            }));
        };

        // --- CaseCard Component ---
        // Represents a single case with its input fields.
        const CaseCard = ({ caseData, index, onUpdate, onFocus, isActive, onSaveIndividual, onTempSave, connectionStatus, useGlobalGesprokenMet, handleIndividualTempSave, handleIndividualPrepareForDocGen, handleIndividualFinalize }) => {
            const { id, zaaknummer, feitcode, cjibNummer, cjibLast4, betrokkene, eigenaar, soort, aantekeninghoorverzoek, feitomschrijving, vooronderzoek, reactie, hearingDate, startTime, endTime, verslaglegger, gesprokenMet, status, isModified, sharePointId } = caseData;
            
            // Add debounce timer for duplicate checking
            const [duplicateCheckTimer, setDuplicateCheckTimer] = useState(null);
            
            // Cleanup timer on component unmount
            useEffect(() => {
                return () => {
                    if (duplicateCheckTimer) {
                        clearTimeout(duplicateCheckTimer);
                    }
                };
            }, [duplicateCheckTimer]);

            const handleInputChange = (e) => {
                const { name, value } = e.target;
                let updatedData = { ...caseData, [name]: value, isModified: true };
                
                // Auto-generate CJIB Last 4 when CJIB number changes
                if (name === 'cjibNummer') {
                    const last4 = value.slice(-4);
                    updatedData.cjibLast4 = last4;
                }
                
                // Auto-calculate end time when start time changes
                if (name === 'startTime' && value) {
                    const endTime = calculateEndTime(value);
                    if (endTime) {
                        updatedData.endTime = endTime;
                    }
                }
                
                // Check for duplicates when zaaknummer changes (with debounce)
                if (name === 'zaaknummer' && value.trim() !== '' && !caseData.sharePointId) {
                    // Clear existing timer
                    if (duplicateCheckTimer) {
                        clearTimeout(duplicateCheckTimer);
                    }
                    
                    // Set new timer for duplicate checking (1 second delay)
                    const newTimer = setTimeout(() => {
                        checkForDuplicate(value.trim(), index);
                    }, 1000);
                    
                    setDuplicateCheckTimer(newTimer);
                }
                
                onUpdate(index, updatedData);
            };

            // Function to check for duplicate zaaknummer
            const checkForDuplicate = async (zaaknummer, caseIndex) => {
                try {
                    const existingCase = await sharePointService.getCaseByZaaknummer(zaaknummer);
                    if (existingCase) {
                        // Show confirmation dialog to user
                        const confirmLoad = confirm(
                            `Zaak "${zaaknummer}" bestaat al in SharePoint.\n\n` +
                            `Wilt u de bestaande gegevens laden?\n\n` +
                            `Ja = Bestaande gegevens laden\n` +
                            `Nee = Doorgaan met nieuwe zaak (kan duplicaat maken)`
                        );
                        
                        if (confirmLoad) {
                            // Load existing case data
                            const loadedCaseData = {
                                id: caseData.id,
                                sharePointId: existingCase.Id,
                                zaaknummer: existingCase.Title || '',
                                feitcode: existingCase.Feitcode || '',
                                cjibNummer: existingCase.CJIBNummer || '',
                                cjibLast4: (existingCase.CJIBNummer || '').slice(-4),
                                betrokkene: existingCase.Betrokkene || '',
                                eigenaar: existingCase.Eigenaar || '',
                                soort: existingCase.Soort || '',
                                aantekeninghoorverzoek: existingCase.AantekeningHoorverzoek || '',
                                feitomschrijving: existingCase.Feitomschrijving || '',
                                vooronderzoek: existingCase.Vooronderzoek || '',
                                reactie: existingCase.ReactiePMBU || '',
                                hearingDate: existingCase.HearingDate ? new Date(existingCase.HearingDate).toISOString().split('T')[0] : '',
                                startTime: existingCase.StartTime || '',
                                endTime: existingCase.EndTime || '',
                                verslaglegger: existingCase.Verslaglegger || '',
                                gesprokenMet: existingCase.GesprokenMet || '',
                                bedrijfsnaam: existingCase.Bedrijfsnaam || '',
                                status: existingCase.Status || 'Bezig met uitwerken',
                                isModified: false, // Not modified since we just loaded it
                            };
                            
                            onUpdate(caseIndex, loadedCaseData);
                        }
                    }
                } catch (error) {
                    console.error('Error checking for duplicate:', error);
                    // Don't show error to user for duplicate checking
                }
            };

            const handleFocus = () => {
                onFocus(index);
            };

            const handleSaveCase = () => {
                onSaveIndividual(index);
            };

            const handleTempSave = () => {
                onTempSave(index);
            };

            const cardBorderColor = isModified ? 'border-blue-500' : 'border-gray-200';
            const activeShadow = isActive ? 'shadow-xl' : 'shadow-md';
            const hasSharePointId = sharePointId !== null && sharePointId !== undefined;

            return html`
                <div 
                    id=${`case-card-${index}`}
                    class="bg-white p-6 rounded-lg border-l-4 ${cardBorderColor} ${activeShadow} transition-all duration-300 mb-4"
                >
                    <div class="flex justify-between items-center mb-4">
                        <div class="flex items-center space-x-3">
                            <h3 class="text-xl font-bold text-gray-700">Zaak #${index + 1}</h3>
                            ${hasSharePointId && html`
                                <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
                                    <svg class="w-3 h-3 mr-1" fill="currentColor" viewBox="0 0 20 20">
                                        <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd"></path>
                                    </svg>
                                    SharePoint
                                </span>
                            `}
                        </div>
                    </div>
                    <div class="grid grid-cols-1 gap-6">
                        <!-- Group 1: Basic Case Information -->
                        <div class="bg-gray-50 p-4 rounded-lg">
                            <h4 class="text-sm font-semibold text-gray-700 mb-3 uppercase tracking-wide">Zaak Informatie</h4>
                            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                                <!-- Zaaknummer -->
                                <div class="flex flex-col">
                                    <label for=${`zaaknummer-${id}`} class="mb-1 font-semibold text-gray-600">Zaaknummer</label>
                                    <input
                                        type="text"
                                        id=${`zaaknummer-${id}`}
                                        name="zaaknummer"
                                        value=${zaaknummer}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                        placeholder="bv. 123456789"
                                    />
                                </div>

                                <!-- CJIB Nummer -->
                                <div class="flex flex-col">
                                    <label for=${`cjibNummer-${id}`} class="mb-1 font-semibold text-gray-600">CJIB Nummer</label>
                                    <input
                                        type="text"
                                        id=${`cjibNummer-${id}`}
                                        name="cjibNummer"
                                        value=${cjibNummer}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                        placeholder="Volledig CJIB nummer"
                                    />
                                </div>

                                <!-- CJIB Laatste 4 (Read-only) -->
                                <div class="flex flex-col">
                                    <label for=${`cjibLast4-${id}`} class="mb-1 font-semibold text-gray-600">CJIB Laatste 4</label>
                                    <input
                                        type="text"
                                        id=${`cjibLast4-${id}`}
                                        name="cjibLast4"
                                        value=${cjibLast4}
                                        readonly
                                        class="p-3 border border-gray-300 rounded-md bg-gray-50 text-gray-600 outline-none"
                                        placeholder="Auto"
                                    />
                                </div>
                            </div>
                        </div>

                        <!-- Group 2: Party Information -->
                        <div class="bg-gray-50 p-4 rounded-lg">
                            <h4 class="text-sm font-semibold text-gray-700 mb-3 uppercase tracking-wide">Betrokkenen</h4>
                            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                                <!-- Betrokkene -->
                                <div class="flex flex-col">
                                    <label for=${`betrokkene-${id}`} class="mb-1 font-semibold text-gray-600">Betrokkene</label>
                                    <input
                                        type="text"
                                        id=${`betrokkene-${id}`}
                                        name="betrokkene"
                                        value=${betrokkene}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                        placeholder="Naam van de betrokkene"
                                    />
                                </div>

                                <!-- Eigenaar -->
                                <div class="flex flex-col">
                                    <label for=${`eigenaar-${id}`} class="mb-1 font-semibold text-gray-600">Eigenaar</label>
                                    <input
                                        type="text"
                                        id=${`eigenaar-${id}`}
                                        name="eigenaar"
                                        value=${eigenaar}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                        placeholder="Naam van de eigenaar"
                                    />
                                </div>

                                <!-- Soort -->
                                <div class="flex flex-col">
                                    <label for=${`soort-${id}`} class="mb-1 font-semibold text-gray-600">Soort</label>
                                    <input
                                        type="text"
                                        id=${`soort-${id}`}
                                        name="soort"
                                        value=${soort}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                        placeholder="Soort zaak/overtreding"
                                    />
                                </div>
                            </div>
                        </div>

                        <!-- Group 3: Timing Information -->
                        <div class="bg-gray-50 p-4 rounded-lg">
                            <h4 class="text-sm font-semibold text-gray-700 mb-3 uppercase tracking-wide">Tijd en Datum</h4>
                            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                                <!-- Datum Hoorzitting -->
                                <div class="flex flex-col">
                                    <label for=${`hearingDate-${id}`} class="mb-1 font-semibold text-gray-600">Datum Hoorzitting</label>
                                    <input
                                        type="date"
                                        id=${`hearingDate-${id}`}
                                        name="hearingDate"
                                        value=${hearingDate}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                    />
                                </div>

                                <!-- Starttijd -->
                                <div class="flex flex-col">
                                    <label for=${`startTime-${id}`} class="mb-1 font-semibold text-gray-600">Starttijd</label>
                                    <input
                                        type="time"
                                        id=${`startTime-${id}`}
                                        name="startTime"
                                        value=${startTime}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                    />
                                </div>

                                <!-- Eindtijd -->
                                <div class="flex flex-col">
                                    <label for=${`endTime-${id}`} class="mb-1 font-semibold text-gray-600">Eindtijd</label>
                                    <input
                                        type="time"
                                        id=${`endTime-${id}`}
                                        name="endTime"
                                        value=${endTime}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                    />
                                </div>
                            </div>
                        </div>

                        <!-- Group 4: Violation Information -->
                        <div class="bg-gray-50 p-4 rounded-lg">
                            <h4 class="text-sm font-semibold text-gray-700 mb-3 uppercase tracking-wide">Overtreding Details</h4>
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                                <!-- Feitcode -->
                                <div class="flex flex-col">
                                    <label for=${`feitcode-${id}`} class="mb-1 font-semibold text-gray-600">Feitcode</label>
                                    <input
                                        type="text"
                                        id=${`feitcode-${id}`}
                                        name="feitcode"
                                        value=${feitcode}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                        placeholder="bv. R584"
                                    />
                                </div>

                                <!-- Status -->
                                <div class="flex flex-col">
                                    <label for=${`status-${id}`} class="mb-1 font-semibold text-gray-600">Status</label>
                                    <select
                                        id=${`status-${id}`}
                                        name="status"
                                        value=${status}
                                        onChange=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                    >
                                        ${STATUS_CHOICES.map(choice => html`
                                            <option value=${choice} selected=${status === choice}>${choice}</option>
                                        `)}
                                    </select>
                                </div>
                            </div>
                            
                            <!-- Feitomschrijving (full width) -->
                            <div class="mt-4 flex flex-col">
                                <label for=${`feitomschrijving-${id}`} class="mb-1 font-semibold text-gray-600">Feitomschrijving</label>
                                <input
                                    type="text"
                                    id=${`feitomschrijving-${id}`}
                                    name="feitomschrijving"
                                    value=${feitomschrijving}
                                    onInput=${handleInputChange}
                                    onFocus=${handleFocus}
                                    class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                    placeholder="Omschrijving van de overtreding"
                                />
                            </div>
                        </div>

                        <!-- Group 5: Investigation & Communication -->
                        <div class="bg-gray-50 p-4 rounded-lg">
                            <h4 class="text-sm font-semibold text-gray-700 mb-3 uppercase tracking-wide">Onderzoek en Communicatie</h4>
                            
                            <!-- Vooronderzoek -->
                            <div class="mb-4 flex flex-col">
                                <label for=${`vooronderzoek-${id}`} class="mb-1 font-semibold text-gray-600">Vooronderzoek</label>
                                <textarea
                                    id=${`vooronderzoek-${id}`}
                                    name="vooronderzoek"
                                    value=${vooronderzoek}
                                    onInput=${handleInputChange}
                                    onFocus=${handleFocus}
                                    class="p-3 border border-gray-300 rounded-md h-24 resize-y focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                    placeholder="Resultaten van het vooronderzoek..."
                                ></textarea>
                            </div>

                            <!-- Gesproken Met (conditional) -->
                            ${!useGlobalGesprokenMet && html`
                                <div class="mb-4 flex flex-col">
                                    <label for=${`gesprokenMet-${id}`} class="mb-1 font-semibold text-gray-600">Gesproken Met</label>
                                    <input
                                        type="text"
                                        id=${`gesprokenMet-${id}`}
                                        name="gesprokenMet"
                                        value=${gesprokenMet}
                                        onInput=${handleInputChange}
                                        onFocus=${handleFocus}
                                        class="p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                        placeholder="Met wie is er gesproken?"
                                    />
                                </div>
                            `}

                            <!-- Reactie burger/gemachtigde -->
                            <div class="mb-4 flex flex-col">
                                <label for=${`reactie-${id}`} class="mb-1 font-semibold text-gray-600">Reactie burger/gemachtigde</label>
                                <textarea
                                    id=${`reactie-${id}`}
                                    name="reactie"
                                    value=${reactie}
                                    onInput=${handleInputChange}
                                    onFocus=${handleFocus}
                                    class="p-3 border border-gray-300 rounded-md h-32 resize-y focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                    placeholder="Noteer hier het gesprek..."
                                ></textarea>
                            </div>

                            <!-- Aantekening Hoorverzoek -->
                            <div class="flex flex-col">
                                <label for=${`aantekeninghoorverzoek-${id}`} class="mb-1 font-semibold text-gray-600">Aantekening Hoorverzoek</label>
                                <textarea
                                    id=${`aantekeninghoorverzoek-${id}`}
                                    name="aantekeninghoorverzoek"
                                    value=${aantekeninghoorverzoek}
                                    onInput=${handleInputChange}
                                    onFocus=${handleFocus}
                                    class="p-3 border border-gray-300 rounded-md h-24 resize-y focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition"
                                    placeholder="Aantekeningen betreffende het hoorverzoek..."
                                ></textarea>
                            </div>
                        </div>
                        
                        <!-- Individual Action Buttons -->
                        <div class="bg-gray-100 p-4 rounded-lg border-t border-gray-200">
                            <div class="flex justify-between items-center flex-wrap gap-2">
                                <div class="flex space-x-2">
                                    <button
                                        onClick=${() => handleIndividualTempSave(index)}
                                        disabled=${connectionStatus !== 'success'}
                                        class="bg-orange-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-orange-700 focus:outline-none focus:ring-2 focus:ring-orange-300 transition-all duration-300 text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                                        title="Tijdelijk opslaan - status wordt 'In behandeling'"
                                    >
                                        Tijdelijk Opslaan
                                    </button>
                                    <button
                                        onClick=${() => handleIndividualPrepareForDocGen(index)}
                                        disabled=${connectionStatus !== 'success'}
                                        class="bg-yellow-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-yellow-700 focus:outline-none focus:ring-2 focus:ring-yellow-300 transition-all duration-300 text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                                        title="Klaarzetten voor DocGen"
                                    >
                                        Klaarzetten DocGen
                                    </button>
                                    <button
                                        onClick=${() => handleIndividualFinalize(index)}
                                        disabled=${connectionStatus !== 'success'}
                                        class="bg-green-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-green-300 transition-all duration-300 text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                                        title="Definitief afhandelen - status wordt 'Afgehandeld'"
                                    >
                                        Definitief
                                    </button>
                                </div>
                                
                                <!-- Legacy Individual Save Button -->
                                <div class="flex space-x-2">
                                    ${hasSharePointId && html`
                                        <button
                                            onClick=${handleTempSave}
                                            disabled=${connectionStatus !== 'success'}
                                            class="bg-orange-500 text-white font-bold py-2 px-4 rounded-lg hover:bg-orange-600 focus:outline-none focus:ring-2 focus:ring-orange-300 transition-all duration-300 text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                                            title="Oude tijdelijke opslag (behoudt huidige status)"
                                        >
                                            Temp. Opslaan (Legacy)
                                        </button>
                                    `}
                                    <button
                                        onClick=${handleSaveCase}
                                        disabled=${connectionStatus !== 'success'}
                                        class="bg-blue-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-300 transition-all duration-300 text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                                        title=${hasSharePointId ? "Opslaan zonder status wijziging" : "Nieuwe zaak opslaan"}
                                    >
                                        ${hasSharePointId ? "Opslaan" : "Nieuwe Zaak"}
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        };

        // --- Main App Component ---
        // Manages the state for all cases and the overall application logic.
        const App = () => {
            const [cases, setCases] = useState(() => createInitialCases(20));
            const [activeCaseIndex, setActiveCaseIndex] = useState(0);
            const [showInfoModal, setShowInfoModal] = useState(false);
            const [showConfirmModal, setShowConfirmModal] = useState(false);
            const [modalContent, setModalContent] = useState({ title: '', message: '' });
            const [isLoading, setIsLoading] = useState(false);
            const [connectionStatus, setConnectionStatus] = useState('checking'); // checking, success, failed
            const [showDateMenu, setShowDateMenu] = useState(false);
            const [availableDates, setAvailableDates] = useState([]);
            const [loadingDates, setLoadingDates] = useState(false);
            const [globalVerslaglegger, setGlobalVerslaglegger] = useState('');
            const [globalGesprokenMet, setGlobalGesprokenMet] = useState('');
            const [useGlobalGesprokenMet, setUseGlobalGesprokenMet] = useState(true);
            const [isGemachtigde, setIsGemachtigde] = useState(true);
            const [globalBedrijfsnaam, setGlobalBedrijfsnaam] = useState('');

            // Test SharePoint connection on load
            useEffect(() => {
                const testConnection = async () => {
                    try {
                        await sharePointService.testConnection();
                        setConnectionStatus('success');
                        console.log('SharePoint connection test successful');
                    } catch (error) {
                        setConnectionStatus('failed');
                        console.error('SharePoint connection test failed:', error);
                        setModalContent({
                            title: 'SharePoint Verbindingsfout',
                            message: `Kan geen verbinding maken met SharePoint: ${error.message}`
                        });
                        setShowInfoModal(true);
                    }
                };
                
                testConnection();
            }, []);

            // Close date menu when clicking outside
            useEffect(() => {
                const handleClickOutside = (event) => {
                    if (showDateMenu && !event.target.closest('.date-menu-container')) {
                        setShowDateMenu(false);
                    }
                };

                if (showDateMenu) {
                    document.addEventListener('mousedown', handleClickOutside);
                    return () => document.removeEventListener('mousedown', handleClickOutside);
                }
            }, [showDateMenu]);

            // Effect to scroll to the active card
            useEffect(() => {
                const activeCard = document.getElementById(`case-card-${activeCaseIndex}`);
                if (activeCard) {
                    activeCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
                }
            }, [activeCaseIndex]);

            // Update a specific case in the state
            const handleUpdateCase = useCallback((index, updatedCase) => {
                const newCases = [...cases];
                newCases[index] = updatedCase;
                setCases(newCases);
            }, [cases]);
            
            // Set the currently focused case
            const handleFocusCase = useCallback((index) => {
                setActiveCaseIndex(index);
            }, []);

            // Excel Import Function
            const handleExcelImport = (event) => {
                const file = event.target.files[0];
                if (!file) return;

                const reader = new FileReader();
                reader.onload = async (e) => {
                    setIsLoading(true); // Show loading during duplicate checking
                    
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const sheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet);

                        // Helper function to find column value case-insensitively
                        const findColumnValue = (row, possibleNames) => {
                            for (const name of possibleNames) {
                                const key = Object.keys(row).find(k => k.toLowerCase() === name.toLowerCase());
                                if (key && row[key] !== undefined && row[key] !== null && row[key] !== '') {
                                    return row[key];
                                }
                            }
                            return '';
                        };

                        // Helper function to parse date and time
                        const parseDateTimeField = (dateTimeStr) => {
                            if (!dateTimeStr) return { date: '', startTime: '', endTime: '' };
                            
                            console.log('Parsing date/time field:', dateTimeStr);
                            
                            try {
                                // Try multiple formats that might appear in Excel
                                const formats = [
                                    // Primary format: dd-mm-yyyy hh:mm
                                    {
                                        regex: /^(\d{1,2})-(\d{1,2})-(\d{4})\s+(\d{1,2}):(\d{2})$/,
                                        name: 'dd-mm-yyyy hh:mm'
                                    },
                                    // Alternative format: dd/mm/yyyy hh:mm
                                    {
                                        regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})$/,
                                        name: 'dd/mm/yyyy hh:mm'
                                    },
                                    // Alternative format: dd.mm.yyyy hh:mm
                                    {
                                        regex: /^(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})$/,
                                        name: 'dd.mm.yyyy hh:mm'
                                    },
                                    // Format with more flexible spacing
                                    {
                                        regex: /^(\d{1,2})[-\/\.](\d{1,2})[-\/\.](\d{4})\s*(\d{1,2}):(\d{2})$/,
                                        name: 'flexible dd-mm-yyyy hh:mm'
                                    }
                                ];
                                
                                for (const format of formats) {
                                    const match = dateTimeStr.toString().match(format.regex);
                                    if (match) {
                                        const [, day, month, year, hours, minutes] = match;
                                        console.log(`Matched format: ${format.name}`, { day, month, year, hours, minutes });
                                        
                                        // Create proper date object and format to ISO date (YYYY-MM-DD)
                                        const dateObj = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                                        if (isNaN(dateObj.getTime())) {
                                            console.warn('Invalid date components:', day, month, year);
                                            continue; // Try next format
                                        }
                                        
                                        const formattedDate = dateObj.toISOString().split('T')[0]; // YYYY-MM-DD format
                                        
                                        // Format start time as HH:MM
                                        const startTime = `${hours.padStart(2, '0')}:${minutes}`;
                                        
                                        // Calculate end time using helper function
                                        const endTime = calculateEndTime(startTime);
                                        
                                        console.log('Successfully parsed:', { date: formattedDate, startTime, endTime });
                                        return { date: formattedDate, startTime, endTime };
                                    }
                                }
                                
                                console.warn('Date/time format does not match any expected pattern:', dateTimeStr);
                                console.log('Supported formats:');
                                console.log('- dd-mm-yyyy hh:mm (e.g., 15-03-2024 14:30)');
                                console.log('- dd/mm/yyyy hh:mm (e.g., 15/03/2024 14:30)');
                                console.log('- dd.mm.yyyy hh:mm (e.g., 15.03.2024 14:30)');
                                
                            } catch (error) {
                                console.warn('Error parsing date/time:', dateTimeStr, error);
                            }
                            
                            return { date: '', startTime: '', endTime: '' };
                        };

                        // Map Excel data to our case format with duplicate checking
                        const importedCases = [];
                        const duplicateInfo = [];
                        
                        for (let index = 0; index < Math.min(jsonData.length, 20); index++) {
                            const row = jsonData[index];
                            
                            // Find CJIB number (handle both with and without dash)
                            const cjibNumber = findColumnValue(row, ['CJIB-Nummer', 'CJIBNummer', 'cjibNummer', 'CJIB Nummer']);
                            
                            // Parse date and time from combined field
                            const dateTimeField = findColumnValue(row, ['Datum en Tijd hoorzitting', 'Datum en tijd hoorzitting', 'Datum_en_Tijd_hoorzitting']);
                            const { date, startTime, endTime } = parseDateTimeField(dateTimeField);
                            
                            const zaaknummer = findColumnValue(row, ['Registratienummer', 'zaaknummer', 'Zaaknummer']);
                            
                            let caseData = {
                                id: `case-${index}`,
                                sharePointId: null,
                                zaaknummer: zaaknummer,
                                feitcode: findColumnValue(row, ['Feitcode', 'feitcode']),
                                cjibNummer: cjibNumber,
                                cjibLast4: cjibNumber ? cjibNumber.toString().slice(-4) : '',
                                betrokkene: findColumnValue(row, ['Betrokkene', 'betrokkene']),
                                eigenaar: findColumnValue(row, ['Eigenaar', 'eigenaar']),
                                soort: findColumnValue(row, ['Soort', 'soort']),
                                aantekeninghoorverzoek: findColumnValue(row, ['Aantekening hoorverzoek', 'AantekeningHoorverzoek', 'aantekeninghoorverzoek']),
                                feitomschrijving: '', // Set to blank as requested
                                vooronderzoek: findColumnValue(row, ['Vooronderzoek', 'vooronderzoek']),
                                reactie: '',
                                hearingDate: date || ensureISODate(new Date()),
                                startTime: startTime,
                                endTime: endTime,
                                verslaglegger: findColumnValue(row, ['Verslaglegger', 'verslaglegger']),
                                gesprokenMet: '',
                                bedrijfsnaam: findColumnValue(row, ['Bedrijfsnaam', 'bedrijfsnaam', 'Bedrijf']),
                                status: 'Nieuw',
                                isModified: true,
                            };
                            
                            // Check for existing case if zaaknummer is provided
                            if (zaaknummer && zaaknummer.trim() !== '') {
                                try {
                                    const existingCase = await sharePointService.getCaseByZaaknummer(zaaknummer);
                                    if (existingCase) {
                                        // Merge existing data with Excel data, prioritizing Excel data for most fields
                                        caseData = {
                                            ...caseData,
                                            sharePointId: existingCase.Id,
                                            // Keep existing data for fields that are typically not in Excel
                                            feitomschrijving: existingCase.Feitomschrijving || '',
                                            reactie: existingCase.ReactiePMBU || '',
                                            gesprokenMet: existingCase.GesprokenMet || '',
                                            status: existingCase.Status || 'Bezig met uitwerken',
                                            // Excel data takes precedence for most other fields
                                            isModified: true // Mark as modified so it gets updated
                                        };
                                        
                                        duplicateInfo.push(`Zaak ${zaaknummer} - bestaande data samengevoegd`);
                                    }
                                } catch (error) {
                                    console.warn(`Error checking for existing case ${zaaknummer}:`, error);
                                    // Continue with import even if duplicate check fails
                                }
                            }
                            
                            importedCases.push(caseData);
                        }

                        // Fill remaining slots with empty cases
                        while (importedCases.length < 20) {
                            const index = importedCases.length;
                            importedCases.push({
                                id: `case-${index}`,
                                sharePointId: null,
                                zaaknummer: '',
                                feitcode: '',
                                cjibNummer: '',
                                cjibLast4: '',
                                betrokkene: '',
                                eigenaar: '',
                                soort: '',
                                aantekeninghoorverzoek: '',
                                feitomschrijving: '',
                                vooronderzoek: '',
                                reactie: '',
                                hearingDate: ensureISODate(new Date()),
                                startTime: '',
                                endTime: '',
                                verslaglegger: '',
                                gesprokenMet: '',
                                bedrijfsnaam: '',
                                status: 'Nieuw',
                                isModified: false,
                            });
                        }

                        setCases(importedCases);
                        
                        // Create message with duplicate information
                        let message = `${Math.min(jsonData.length, 20)} zaken zijn geïmporteerd uit het Excel bestand.`;
                        if (duplicateInfo.length > 0) {
                            message += `\n\nDuplicaten gevonden en samengevoegd:\n${duplicateInfo.join('\n')}`;
                        }
                        
                        setModalContent({
                            title: 'Excel Import Voltooid',
                            message: message
                        });
                        setShowInfoModal(true);

                    } catch (error) {
                        console.error('Error importing Excel:', error);
                        setModalContent({
                            title: 'Import Fout',
                            message: `Er is een fout opgetreden bij het importeren: ${error.message}`
                        });
                        setShowInfoModal(true);
                    } finally {
                        setIsLoading(false); // Hide loading state
                    }
                };
                reader.readAsArrayBuffer(file);
                
                // Reset file input
                event.target.value = '';
            };

            // Save individual case to SharePoint
            const handleSaveIndividual = async (index) => {
                const caseData = cases[index];
                setIsLoading(true);
                
                try {
                    // Apply global fields based on settings
                    const finalCaseData = {
                        ...caseData,
                        verslaglegger: globalVerslaglegger, // Always global
                        gesprokenMet: useGlobalGesprokenMet ? globalGesprokenMet : caseData.gesprokenMet,
                        bedrijfsnaam: isGemachtigde ? globalBedrijfsnaam : ''
                    };
                    
                    const sharePointData = sharePointService.transformCaseToSharePoint(finalCaseData);
                    
                    let result;
                    if (caseData.sharePointId) {
                        // Update existing item
                        result = await sharePointService.updateItem(caseData.sharePointId, sharePointData);
                    } else {
                        // Create new item
                        result = await sharePointService.createItem(sharePointData);
                        
                        // Update the case with the SharePoint ID
                        const updatedCase = { ...finalCaseData, sharePointId: result.Id, isModified: false };
                        handleUpdateCase(index, updatedCase);
                    }
                    
                    setModalContent({
                        title: 'Zaak Opgeslagen',
                        message: `Zaak #${index + 1} is succesvol opgeslagen in SharePoint.`
                    });
                    setShowInfoModal(true);
                    
                } catch (error) {
                    console.error('Error saving individual case:', error);
                    setModalContent({
                        title: 'Fout',
                        message: `Er is een fout opgetreden bij het opslaan van zaak #${index + 1}: ${error.message}`
                    });
                    setShowInfoModal(true);
                } finally {
                    setIsLoading(false);
                }
            };

            // Temporary save for existing cases (update only)
            const handleTempSave = async (index) => {
                const caseData = cases[index];
                
                // Only allow temp save for existing SharePoint items
                if (!caseData.sharePointId) {
                    setModalContent({
                        title: 'Geen bestaande zaak',
                        message: `Zaak #${index + 1} moet eerst definitief worden opgeslagen voordat deze tijdelijk kan worden bijgewerkt.`
                    });
                    setShowInfoModal(true);
                    return;
                }
                
                setIsLoading(true);
                
                try {
                    // Apply global fields based on settings
                    const finalCaseData = {
                        ...caseData,
                        verslaglegger: globalVerslaglegger, // Always global
                        gesprokenMet: useGlobalGesprokenMet ? globalGesprokenMet : caseData.gesprokenMet,
                        bedrijfsnaam: isGemachtigde ? globalBedrijfsnaam : ''
                    };
                    
                    const sharePointData = sharePointService.transformCaseToSharePoint(finalCaseData);
                    
                    // Add temporary status flag to indicate this is a work-in-progress update
                    const tempData = {
                        ...sharePointData,
                        Status: 'In behandeling' // Force status to indicate work in progress
                    };
                    
                    await sharePointService.updateItem(caseData.sharePointId, tempData);
                    
                    // Update local state to reflect saved changes
                    const updatedCase = { ...finalCaseData, isModified: false, status: 'In behandeling' };
                    handleUpdateCase(index, updatedCase);
                    
                    setModalContent({
                        title: 'Tijdelijk Opgeslagen',
                        message: `Zaak #${index + 1} is tijdelijk opgeslagen. Status is ingesteld op 'In behandeling' voor verdere bewerking.`
                    });
                    setShowInfoModal(true);
                    
                } catch (error) {
                    console.error('Error temporary saving case:', error);
                    setModalContent({
                        title: 'Fout',
                        message: `Er is een fout opgetreden bij het tijdelijk opslaan van zaak #${index + 1}: ${error.message}`
                    });
                    setShowInfoModal(true);
                } finally {
                    setIsLoading(false);
                }
            };

            // Handle saving all cases
            const handleSaveAll = async () => {
                setIsLoading(true);
                const errors = [];
                const successes = [];
                
                try {
                    for (let i = 0; i < cases.length; i++) {
                        const caseData = cases[i];
                        
                        // Only save cases that have some data or are modified
                        if (caseData.isModified || caseData.zaaknummer || caseData.feitcode || caseData.reactie) {
                            try {
                                // Apply global fields based on settings
                                const finalCaseData = {
                                    ...caseData,
                                    verslaglegger: globalVerslaglegger, // Always global
                                    gesprokenMet: useGlobalGesprokenMet ? globalGesprokenMet : caseData.gesprokenMet,
                                    bedrijfsnaam: isGemachtigde ? globalBedrijfsnaam : ''
                                };
                                
                                const sharePointData = sharePointService.transformCaseToSharePoint(finalCaseData);
                                
                                let result;
                                if (caseData.sharePointId) {
                                    result = await sharePointService.updateItem(caseData.sharePointId, sharePointData);
                                    successes.push(`Zaak #${i + 1} bijgewerkt`);
                                } else {
                                    result = await sharePointService.createItem(sharePointData);
                                    // Update the case with the SharePoint ID
                                    const updatedCase = { ...finalCaseData, sharePointId: result.Id, isModified: false };
                                    handleUpdateCase(i, updatedCase);
                                    successes.push(`Zaak #${i + 1} aangemaakt`);
                                }
                            } catch (error) {
                                console.error(`Error saving case ${i + 1}:`, error);
                                errors.push(`Zaak #${i + 1}: ${error.message}`);
                            }
                        }
                    }
                    
                    if (errors.length === 0) {
                        setModalContent({
                            title: 'Alle Zaken Opgeslagen',
                            message: `${successes.length} zaken zijn succesvol opgeslagen in SharePoint.`
                        });
                    } else {
                        setModalContent({
                            title: 'Gedeeltelijk Opgeslagen',
                            message: `${successes.length} zaken opgeslagen. ${errors.length} fouten:\n${errors.join('\n')}`
                        });
                    }
                    
                } catch (error) {
                    console.error('Error in bulk save:', error);
                    setModalContent({
                        title: 'Fout',
                        message: `Er is een algemene fout opgetreden: ${error.message}`
                    });
                } finally {
                    setIsLoading(false);
                    setShowInfoModal(true);
                }
            };

            // Bulk status update function
            const updateAllCasesStatus = async (newStatus, actionName) => {
                setIsLoading(true);
                const errors = [];
                const successes = [];
                
                try {
                    for (let i = 0; i < cases.length; i++) {
                        const caseData = cases[i];
                        
                        // Only update cases that have data (not empty cases)
                        if (caseData.zaaknummer || caseData.feitcode || caseData.sharePointId) {
                            try {
                                // Apply global fields based on settings
                                const finalCaseData = {
                                    ...caseData,
                                    status: newStatus,
                                    verslaglegger: globalVerslaglegger, // Always global
                                    gesprokenMet: useGlobalGesprokenMet ? globalGesprokenMet : caseData.gesprokenMet,
                                    bedrijfsnaam: isGemachtigde ? globalBedrijfsnaam : ''
                                };
                                
                                const sharePointData = sharePointService.transformCaseToSharePoint(finalCaseData);
                                
                                let result;
                                if (caseData.sharePointId) {
                                    // Update existing item
                                    result = await sharePointService.updateItem(caseData.sharePointId, sharePointData);
                                } else {
                                    // Create new item
                                    result = await sharePointService.createItem(sharePointData);
                                    
                                    // Update the case with the SharePoint ID
                                    const updatedCase = { ...finalCaseData, sharePointId: result.Id, isModified: false };
                                    handleUpdateCase(i, updatedCase);
                                }
                                
                                // Update local state
                                const updatedCase = { ...finalCaseData, isModified: false };
                                handleUpdateCase(i, updatedCase);
                                successes.push(`Zaak #${i + 1}`);
                                
                            } catch (error) {
                                console.error(`Error updating case ${i + 1}:`, error);
                                errors.push(`Zaak #${i + 1}: ${error.message}`);
                            }
                        }
                    }
                    
                    if (errors.length === 0) {
                        setModalContent({
                            title: `${actionName} Voltooid`,
                            message: `${successes.length} zaken zijn bijgewerkt naar status "${newStatus}".`
                        });
                    } else {
                        setModalContent({
                            title: `${actionName} Gedeeltelijk Voltooid`,
                            message: `${successes.length} zaken bijgewerkt. ${errors.length} fouten:\n${errors.join('\n')}`
                        });
                    }
                    
                } catch (error) {
                    console.error(`Error in ${actionName}:`, error);
                    setModalContent({
                        title: 'Fout',
                        message: `Er is een algemene fout opgetreden: ${error.message}`
                    });
                } finally {
                    setIsLoading(false);
                    setShowInfoModal(true);
                }
            };

            // Handle tijdelijk opslaan all (set all to "In behandeling")
            const handleTempSaveAll = async () => {
                await updateAllCasesStatus('In behandeling', 'Tijdelijk Opslaan');
            };

            // Handle klaarzetten voor DocGen all
            const handlePrepareForDocGen = async () => {
                await updateAllCasesStatus('Klaarzetten voor DocGen', 'Klaarzetten voor DocGen');
            };

            // Handle definitief all (set all to "Afgehandeld")
            const handleFinalizeAll = async () => {
                await updateAllCasesStatus('Afgehandeld', 'Definitief Afhandelen');
            };

            // Individual status update function
            const updateIndividualCaseStatus = async (index, newStatus, actionName) => {
                const caseData = cases[index];
                
                // Check if case has SharePoint ID for updates
                if (!caseData.sharePointId && !caseData.zaaknummer && !caseData.feitcode) {
                    setModalContent({
                        title: 'Geen data',
                        message: `Zaak #${index + 1} heeft geen data om bij te werken.`
                    });
                    setShowInfoModal(true);
                    return;
                }
                
                setIsLoading(true);
                
                try {
                    // Apply global fields based on settings
                    const finalCaseData = {
                        ...caseData,
                        status: newStatus,
                        verslaglegger: globalVerslaglegger, // Always global
                        gesprokenMet: useGlobalGesprokenMet ? globalGesprokenMet : caseData.gesprokenMet,
                        bedrijfsnaam: isGemachtigde ? globalBedrijfsnaam : ''
                    };
                    
                    const sharePointData = sharePointService.transformCaseToSharePoint(finalCaseData);
                    
                    let result;
                    if (caseData.sharePointId) {
                        // Update existing item
                        result = await sharePointService.updateItem(caseData.sharePointId, sharePointData);
                    } else {
                        // Create new item
                        result = await sharePointService.createItem(sharePointData);
                        
                        // Update the case with the SharePoint ID
                        const updatedCase = { ...finalCaseData, sharePointId: result.Id, isModified: false };
                        handleUpdateCase(index, updatedCase);
                    }
                    
                    // Update local state
                    const updatedCase = { ...finalCaseData, isModified: false };
                    handleUpdateCase(index, updatedCase);
                    
                    setModalContent({
                        title: `${actionName} Voltooid`,
                        message: `Zaak #${index + 1} is bijgewerkt naar status "${newStatus}".`
                    });
                    setShowInfoModal(true);
                    
                } catch (error) {
                    console.error(`Error in ${actionName} for case ${index + 1}:`, error);
                    setModalContent({
                        title: 'Fout',
                        message: `Er is een fout opgetreden bij ${actionName} van zaak #${index + 1}: ${error.message}`
                    });
                    setShowInfoModal(true);
                } finally {
                    setIsLoading(false);
                }
            };

            // Individual status update handlers
            const handleIndividualTempSave = async (index) => {
                await updateIndividualCaseStatus(index, 'In behandeling', 'Tijdelijk Opslaan');
            };

            const handleIndividualPrepareForDocGen = async (index) => {
                await updateIndividualCaseStatus(index, 'Klaarzetten voor DocGen', 'Klaarzetten voor DocGen');
            };

            const handleIndividualFinalize = async (index) => {
                await updateIndividualCaseStatus(index, 'Afgehandeld', 'Definitief Afhandelen');
            };
            
            // --- Reset Logic ---
            const handleResetAll = () => {
                setShowConfirmModal(true);
            };
            
            const confirmReset = () => {
                setCases(createInitialCases(20));
                setActiveCaseIndex(0);
                setShowConfirmModal(false);
                setModalContent({
                    title: 'Formulier Gereset',
                    message: 'Alle velden zijn leeggemaakt.'
                });
                setShowInfoModal(true);
            };

            const cancelReset = () => {
                setShowConfirmModal(false);
            };
            
            const closeInfoModal = () => {
                setShowInfoModal(false);
            };

            // Date menu functions
            const handleToggleDateMenu = async () => {
                if (!showDateMenu && availableDates.length === 0) {
                    // Load available dates when opening menu for the first time
                    setLoadingDates(true);
                    try {
                        const dates = await sharePointService.getAvailableDatesWithIncompleteCases();
                        setAvailableDates(dates);
                    } catch (error) {
                        console.error('Error loading available dates:', error);
                        setModalContent({
                            title: 'Fout bij Laden Datums',
                            message: `Kan beschikbare datums niet laden: ${error.message}`
                        });
                        setShowInfoModal(true);
                    } finally {
                        setLoadingDates(false);
                    }
                }
                setShowDateMenu(!showDateMenu);
            };

            const handleLoadCasesForDate = async (selectedDate) => {
                setIsLoading(true);
                setShowDateMenu(false);
                
                try {
                    const sharePointCases = await sharePointService.getCasesByDate(selectedDate);
                    
                    // Transform SharePoint data to our case format
                    const transformedCases = sharePointCases.map((spCase, index) => ({
                        id: `case-${index}`,
                        sharePointId: spCase.Id,
                        zaaknummer: spCase.Title || '',
                        feitcode: spCase.Feitcode || '',
                        cjibNummer: spCase.CJIBNummer || '',
                        cjibLast4: (spCase.CJIBNummer || '').slice(-4),
                        betrokkene: spCase.Betrokkene || '',
                        eigenaar: spCase.Eigenaar || '',
                        soort: spCase.Soort || '',
                        aantekeninghoorverzoek: spCase.AantekeningHoorverzoek || '',
                        feitomschrijving: spCase.Feitomschrijving || '',
                        vooronderzoek: spCase.Vooronderzoek || '',
                        reactie: spCase.ReactiePMBU || '',
                        hearingDate: spCase.HearingDate ? new Date(spCase.HearingDate).toISOString().split('T')[0] : '',
                        startTime: spCase.StartTime || '',
                        endTime: spCase.EndTime || '',
                        verslaglegger: spCase.Verslaglegger || '',
                        gesprokenMet: spCase.GesprokenMet || '',
                        bedrijfsnaam: spCase.Bedrijfsnaam || '',
                        status: spCase.Status || 'Bezig met uitwerken',
                        isModified: false,
                    }));

                    // Fill remaining slots with empty cases
                    while (transformedCases.length < 20) {
                        const index = transformedCases.length;
                        transformedCases.push({
                            id: `case-${index}`,
                            sharePointId: null,
                            zaaknummer: '',
                            feitcode: '',
                            cjibNummer: '',
                            cjibLast4: '',
                            betrokkene: '',
                            eigenaar: '',
                            soort: '',
                            aantekeninghoorverzoek: '',
                            feitomschrijving: '',
                            vooronderzoek: '',
                            reactie: '',
                            hearingDate: ensureISODate(selectedDate),
                            startTime: '',
                            endTime: '',
                            verslaglegger: '',
                            gesprokenMet: '',
                            bedrijfsnaam: '',
                            status: 'Nieuw',
                            isModified: false,
                        });
                    }

                    setCases(transformedCases);
                    setActiveCaseIndex(0);
                    
                    const displayDate = new Date(selectedDate).toLocaleDateString('nl-NL');
                    setModalContent({
                        title: 'Zaken Geladen',
                        message: `${sharePointCases.length} onafgeronde zaken voor ${displayDate} zijn geladen.`
                    });
                    setShowInfoModal(true);
                    
                } catch (error) {
                    console.error('Error loading cases for date:', error);
                    setModalContent({
                        title: 'Fout bij Laden Zaken',
                        message: `Kan zaken voor de geselecteerde datum niet laden: ${error.message}`
                    });
                    setShowInfoModal(true);
                } finally {
                    setIsLoading(false);
                }
            };

            return html`
                <div class="bg-gray-50 min-h-screen font-sans">
                    <!-- Loading Overlay -->
                    ${isLoading && html`
                        <div class="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
                            <div class="bg-white p-8 rounded-lg shadow-2xl max-w-sm w-full text-center mx-4">
                                <div class="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
                                <p class="text-gray-700">Bezig met opslaan...</p>
                            </div>
                        </div>
                    `}
                    
                    <!-- Header -->
                    <header class="bg-white shadow-sm sticky top-0 z-20">
                        <!-- Main Header Row -->
                        <div class="container mx-auto px-4 sm:px-6 lg:px-8 py-4">
                            <div class="flex flex-wrap justify-between items-center gap-4">
                                <div class="flex items-center space-x-4">
                                    <h1 class="text-3xl font-bold text-gray-800">Hoorzitting Notulen</h1>
                                    <div class="flex items-center space-x-2">
                                        ${connectionStatus === 'checking' && html`
                                            <div class="w-3 h-3 bg-yellow-500 rounded-full animate-pulse"></div>
                                            <span class="text-sm text-yellow-600">Verbinding testen...</span>
                                        `}
                                        ${connectionStatus === 'success' && html`
                                            <div class="w-3 h-3 bg-green-500 rounded-full"></div>
                                            <span class="text-sm text-green-600">SharePoint verbonden</span>
                                        `}
                                        ${connectionStatus === 'failed' && html`
                                            <div class="w-3 h-3 bg-red-500 rounded-full"></div>
                                            <span class="text-sm text-red-600">Verbindingsfout</span>
                                        `}
                                    </div>
                                </div>
                                <div class="flex items-center justify-between">
                                    <!-- Left side buttons -->
                                    <div class="flex items-center space-x-3">
                                    <!-- Date Menu Button -->
                                    <div class="relative date-menu-container">
                                        <button
                                            onClick=${handleToggleDateMenu}
                                            disabled=${isLoading || connectionStatus !== 'success'}
                                            class="bg-indigo-600 text-white font-bold py-2 px-6 rounded-lg hover:bg-indigo-700 focus:outline-none focus:ring-4 focus:ring-indigo-300 transition-all duration-300 shadow-md disabled:opacity-50 disabled:cursor-not-allowed flex items-center space-x-2"
                                        >
                                            <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"></path>
                                            </svg>
                                            <span>Laden per Datum</span>
                                        </button>
                                        
                                        <!-- Floating Date Menu -->
                                        ${showDateMenu && html`
                                            <div class="absolute top-full right-0 mt-2 w-80 bg-white rounded-lg shadow-xl border border-gray-200 z-30 max-h-96 overflow-y-auto">
                                                <div class="p-4 border-b border-gray-200">
                                                    <h3 class="text-lg font-semibold text-gray-800">Beschikbare Datums</h3>
                                                    <p class="text-sm text-gray-600">Selecteer een datum om onafgeronde zaken te laden</p>
                                                </div>
                                                
                                                ${loadingDates && html`
                                                    <div class="p-6 text-center">
                                                        <div class="animate-spin rounded-full h-8 w-8 border-b-2 border-indigo-600 mx-auto mb-2"></div>
                                                        <p class="text-gray-600">Datums laden...</p>
                                                    </div>
                                                `}
                                                
                                                ${!loadingDates && availableDates.length === 0 && html`
                                                    <div class="p-6 text-center">
                                                        <p class="text-gray-600">Geen onafgeronde zaken gevonden</p>
                                                    </div>
                                                `}
                                                
                                                ${!loadingDates && availableDates.length > 0 && html`
                                                    <div class="max-h-72 overflow-y-auto">
                                                        ${availableDates.map(dateInfo => html`
                                                            <button
                                                                key=${dateInfo.date}
                                                                onClick=${() => handleLoadCasesForDate(dateInfo.date)}
                                                                class="w-full px-4 py-3 text-left hover:bg-gray-50 border-b border-gray-100 transition-colors duration-200 flex justify-between items-center group"
                                                            >
                                                                <div>
                                                                    <div class="font-medium text-gray-800 group-hover:text-indigo-600">
                                                                        ${dateInfo.displayDate}
                                                                    </div>
                                                                    <div class="text-sm text-gray-500">
                                                                        ${dateInfo.count} ${dateInfo.count === 1 ? 'zaak' : 'zaken'} te voltooien
                                                                    </div>
                                                                </div>
                                                                <svg class="w-5 h-5 text-gray-400 group-hover:text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path>
                                                                </svg>
                                                            </button>
                                                        `)}
                                                    </div>
                                                `}
                                                
                                                <div class="p-3 border-t border-gray-200">
                                                    <button
                                                        onClick=${() => setShowDateMenu(false)}
                                                        class="w-full px-3 py-2 text-sm text-gray-600 hover:text-gray-800 hover:bg-gray-50 rounded transition-colors duration-200"
                                                    >
                                                        Sluiten
                                                    </button>
                                                </div>
                                            </div>
                                        `}
                                    </div>
                                    
                                    <input
                                        type="file"
                                        accept=".xlsx,.xls"
                                        onChange=${handleExcelImport}
                                        style="display: none;"
                                        id="excel-import"
                                    />
                                    <button
                                        onClick=${() => document.getElementById('excel-import').click()}
                                        disabled=${isLoading || connectionStatus !== 'success'}
                                        class="bg-purple-600 text-white font-bold py-2 px-6 rounded-lg hover:bg-purple-700 focus:outline-none focus:ring-4 focus:ring-purple-300 transition-all duration-300 shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                                    >
                                        Excel Import
                                    </button>
                                    <button
                                        onClick=${handleResetAll}
                                        disabled=${isLoading}
                                        class="bg-red-600 text-white font-bold py-2 px-6 rounded-lg hover:bg-red-700 focus:outline-none focus:ring-4 focus:ring-red-300 transition-all duration-300 shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                                    >
                                        Resetten
                                    </button>
                                    </div>
                                
                                <!-- Right side buttons -->
                                <div class="flex items-center space-x-3">
                                    <button
                                        onClick=${handleTempSaveAll}
                                        disabled=${isLoading || connectionStatus !== 'success'}
                                        class="bg-orange-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-orange-700 focus:outline-none focus:ring-4 focus:ring-orange-300 transition-all duration-300 shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                                    >
                                        Alles Tijdelijk Opslaan
                                    </button>
                                    <button
                                        onClick=${handlePrepareForDocGen}
                                        disabled=${isLoading || connectionStatus !== 'success'}
                                        class="bg-yellow-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-yellow-700 focus:outline-none focus:ring-4 focus:ring-yellow-300 transition-all duration-300 shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                                    >
                                        Alles Klaarzetten DocGen
                                    </button>
                                    <button
                                        onClick=${handleFinalizeAll}
                                        disabled=${isLoading || connectionStatus !== 'success'}
                                        class="bg-green-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-green-700 focus:outline-none focus:ring-4 focus:ring-green-300 transition-all duration-300 shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                                    >
                                        Alles Definitief
                                    </button>
                                </div>
                            </div>
                        </div>
                        
                        <!-- Global Controls -->
                        <div class="border-t border-gray-200 bg-gray-50">
                            <div class="container mx-auto px-4 sm:px-6 lg:px-8 py-3">
                                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                                    <!-- Row 1: Verslaglegger (always global) and GesprokenMet toggle -->
                                    <div class="flex flex-wrap items-center gap-4">
                                        <!-- Global Verslaglegger (always shown) -->
                                        <div class="flex items-center space-x-3">
                                            <label for="global-verslaglegger" class="text-sm font-medium text-gray-700">Verslaglegger:</label>
                                            <input
                                                type="text"
                                                id="global-verslaglegger"
                                                value=${globalVerslaglegger}
                                                onInput=${(e) => setGlobalVerslaglegger(e.target.value)}
                                                class="px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition w-48"
                                                placeholder="Naam van de verslaglegger"
                                            />
                                        </div>
                                        
                                        <!-- GesprokenMet Toggle -->
                                        <div class="flex items-center space-x-3">
                                            <label class="text-sm font-medium text-gray-700">Gesproken Met:</label>
                                            <div class="relative inline-block w-12 h-6">
                                                <input
                                                    type="checkbox"
                                                    id="gesproken-met-toggle"
                                                    checked=${useGlobalGesprokenMet}
                                                    onChange=${(e) => setUseGlobalGesprokenMet(e.target.checked)}
                                                    class="sr-only"
                                                />
                                                <label
                                                    for="gesproken-met-toggle"
                                                    class=${`block w-12 h-6 rounded-full cursor-pointer transition-colors duration-300 ${useGlobalGesprokenMet ? 'bg-blue-600' : 'bg-gray-300'}`}
                                                >
                                                    <span
                                                        class=${`block w-4 h-4 bg-white rounded-full shadow transform transition-transform duration-300 mt-1 ${useGlobalGesprokenMet ? 'translate-x-7' : 'translate-x-1'}`}
                                                    ></span>
                                                </label>
                                            </div>
                                            <span class="text-sm text-gray-600">
                                                ${useGlobalGesprokenMet ? 'Globaal' : 'Per zaak'}
                                            </span>
                                        </div>
                                    </div>
                                    
                                    <!-- Row 2: GesprokenMet input and Gemachtigde/Burger toggle -->
                                    <div class="flex flex-wrap items-center gap-4">
                                        <!-- Global GesprokenMet Input (only shown when toggle is ON) -->
                                        ${useGlobalGesprokenMet && html`
                                            <div class="flex items-center space-x-3">
                                                <label for="global-gesproken-met" class="text-sm font-medium text-gray-700">Gesproken Met:</label>
                                                <input
                                                    type="text"
                                                    id="global-gesproken-met"
                                                    value=${globalGesprokenMet}
                                                    onInput=${(e) => setGlobalGesprokenMet(e.target.value)}
                                                    class="px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition w-48"
                                                    placeholder="Met wie gesproken"
                                                />
                                            </div>
                                        `}
                                        
                                        <!-- Gemachtigde/Burger Toggle -->
                                        <div class="flex items-center space-x-3">
                                            <span class="text-sm font-medium text-gray-700">Type:</span>
                                            <div class="relative inline-block w-16 h-6">
                                                <input
                                                    type="checkbox"
                                                    id="type-toggle"
                                                    checked=${isGemachtigde}
                                                    onChange=${(e) => setIsGemachtigde(e.target.checked)}
                                                    class="sr-only"
                                                />
                                                <label
                                                    for="type-toggle"
                                                    class=${`block w-16 h-6 rounded-full cursor-pointer transition-colors duration-300 ${isGemachtigde ? 'bg-green-600' : 'bg-blue-600'}`}
                                                >
                                                    <span
                                                        class=${`block w-4 h-4 bg-white rounded-full shadow transform transition-transform duration-300 mt-1 ${isGemachtigde ? 'translate-x-1' : 'translate-x-11'}`}
                                                    ></span>
                                                </label>
                                            </div>
                                            <span class="text-sm text-gray-600 font-medium">
                                                ${isGemachtigde ? 'Gemachtigde' : 'Burger'}
                                            </span>
                                        </div>
                                        
                                        <!-- Bedrijfsnaam (only shown when Gemachtigde is selected) -->
                                        ${isGemachtigde && html`
                                            <div class="flex items-center space-x-3">
                                                <label for="global-bedrijfsnaam" class="text-sm font-medium text-gray-700">Bedrijfsnaam:</label>
                                                <input
                                                    type="text"
                                                    id="global-bedrijfsnaam"
                                                    value=${globalBedrijfsnaam}
                                                    onInput=${(e) => setGlobalBedrijfsnaam(e.target.value)}
                                                    class="px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-green-500 focus:border-transparent outline-none transition w-48"
                                                    placeholder="Naam van het bedrijf"
                                                />
                                            </div>
                                        `}
                                    </div>
                                </div>
                            </div>
                        </div>
                    </header>

                    <!-- Main Content -->
                    <main class="container mx-auto px-4 sm:px-6 lg:px-8 py-8">
                        <div class="max-w-6xl mx-auto">
                            ${cases.map((caseItem, index) => html`
                                <${CaseCard}
                                    key=${caseItem.id}
                                    caseData=${caseItem}
                                    index=${index}
                                    onUpdate=${handleUpdateCase}
                                    onFocus=${handleFocusCase}
                                    onSaveIndividual=${handleSaveIndividual}
                                    onTempSave=${handleTempSave}
                                    connectionStatus=${connectionStatus}
                                    useGlobalGesprokenMet=${useGlobalGesprokenMet}
                                    isActive=${index === activeCaseIndex}
                                    handleIndividualTempSave=${handleIndividualTempSave}
                                    handleIndividualPrepareForDocGen=${handleIndividualPrepareForDocGen}
                                    handleIndividualFinalize=${handleIndividualFinalize}
                                />
                            `)}
                        </div>
                    </main>
                    
                    <!-- Info Modal Component -->
                    ${showInfoModal && html`
                        <div class="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
                            <div class="bg-white p-8 rounded-lg shadow-2xl max-w-sm w-full text-center mx-4">
                                <h2 class="text-2xl font-bold mb-4">${modalContent.title}</h2>
                                <p class="text-gray-700 mb-6">${modalContent.message}</p>
                                <button
                                    onClick=${closeInfoModal}
                                    class="bg-blue-600 text-white font-bold py-2 px-8 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-300"
                                >
                                    Sluiten
                                </button>
                            </div>
                        </div>
                    `}
                    
                    <!-- Confirmation Modal for Reset -->
                    ${showConfirmModal && html`
                        <div class="fixed inset-0 bg-black bg-opacity-60 flex items-center justify-center z-50">
                            <div class="bg-white p-8 rounded-lg shadow-2xl max-w-sm w-full text-center mx-4">
                                <h2 class="text-2xl font-bold mb-2 text-gray-800">Weet u het zeker?</h2>
                                <p class="text-gray-600 mb-6">Hiermee worden alle gegevens op de pagina gewist. Deze actie kan niet ongedaan worden gemaakt.</p>
                                <div class="flex justify-center space-x-4">
                                     <button
                                        onClick=${cancelReset}
                                        class="bg-gray-300 text-gray-800 font-bold py-2 px-8 rounded-lg hover:bg-gray-400 focus:outline-none focus:ring-4 focus:ring-gray-200"
                                    >
                                        Annuleren
                                    </button>
                                    <button
                                        onClick=${confirmReset}
                                        class="bg-red-600 text-white font-bold py-2 px-8 rounded-lg hover:bg-red-700 focus:outline-none focus:ring-4 focus:ring-red-300"
                                    >
                                        Bevestigen
                                    </button>
                                </div>
                            </div>
                        </div>
                    `}
                </div>
            `;
        };

        // --- Render the App ---
        render(html`<${App} />`, document.getElementById('app'));
    </script>
</head>
<body class="bg-gray-50">
    <div id="app"></div>
</body>
</html>
