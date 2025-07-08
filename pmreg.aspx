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

        // Status choices (matching SharePoint exactly)
        const STATUS_CHOICES = [
            'Bezig met uitwerken',
            'Aangehouden',
            'Afgerond'  // Changed from 'Afgehandeld' to match SharePoint
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
                    
                    const itemWithUser = {
                        ...itemData,
                        Username: currentUser.Title || currentUser.LoginName
                    };

                    const response = await fetch(`${this.apiUrl}lists/getbytitle('${this.listName}')/items`, {
                        method: 'POST',
                        headers: {
                            'Accept': 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose',
                            'X-RequestDigest': digest
                        },
                        credentials: 'include',
                        body: JSON.stringify(itemWithUser)
                    });
                    
                    if (!response.ok) {
                        const errorText = await response.text();
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
                        body: JSON.stringify(itemData)
                    });
                    
                    if (!response.ok) {
                        const errorText = await response.text();
                        throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
                    }
                    
                    return { success: true, itemId };
                } catch (error) {
                    console.error('Error updating item:', error);
                    throw error;
                }
            }

            transformCaseToSharePoint(caseData) {
                return {
                    Title: caseData.zaaknummer || '',
                    Feitcode: caseData.feitcode || '',
                    CJIBNummer: caseData.cjibNummer || '',
                    CJIBLast4: caseData.cjibLast4 || '',
                    Feitomschrijving: caseData.feitomschrijving || '',
                    Vooronderzoek: caseData.vooronderzoek || '',
                    ReactiePMBU: caseData.reactie || '',
                    HearingDate: caseData.hearingDate || null,
                    StartTime: caseData.startTime || '',
                    EndTime: caseData.endTime || '',
                    Status: caseData.status || 'Bezig met uitwerken'
                };
            }
        }

        const sharePointService = new SharePointService();

        // --- Helper function to generate initial empty cases ---
        const createInitialCases = (count) => {
            return Array.from({ length: count }, (_, i) => ({
                id: `case-${i}`,
                sharePointId: null,
                zaaknummer: '',
                feitcode: '',
                cjibNummer: '',
                cjibLast4: '',
                feitomschrijving: '',
                vooronderzoek: '',
                reactie: '',
                hearingDate: new Date().toISOString().split('T')[0], // Today's date
                startTime: '',
                endTime: '',
                status: 'Bezig met uitwerken', // Matches SharePoint exactly
                isModified: false,
            }));
        };

        // --- CaseCard Component ---
        // Represents a single case with its input fields.
        const CaseCard = ({ caseData, index, onUpdate, onFocus, isActive, onSaveIndividual, onTempSave, connectionStatus }) => {
            const { id, zaaknummer, feitcode, cjibNummer, cjibLast4, feitomschrijving, vooronderzoek, reactie, hearingDate, startTime, endTime, status, isModified, sharePointId } = caseData;

            const handleInputChange = (e) => {
                const { name, value } = e.target;
                let updatedData = { ...caseData, [name]: value, isModified: true };
                
                // Auto-generate CJIB Last 4 when CJIB number changes
                if (name === 'cjibNummer') {
                    const last4 = value.slice(-4);
                    updatedData.cjibLast4 = last4;
                }
                
                onUpdate(index, updatedData);
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
                        <h3 class="text-xl font-bold text-gray-700">Zaak #${index + 1}</h3>
                        <div class="flex space-x-2">
                            ${hasSharePointId && html`
                                <button
                                    onClick=${handleTempSave}
                                    disabled=${connectionStatus !== 'success'}
                                    class="bg-orange-500 text-white font-bold py-1 px-4 rounded-lg hover:bg-orange-600 focus:outline-none focus:ring-2 focus:ring-orange-300 transition-all duration-300 text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                                    title="Tijdelijk opslaan voor aanpassingen"
                                >
                                    Temp. Opslaan
                                </button>
                            `}
                            <button
                                onClick=${handleSaveCase}
                                disabled=${connectionStatus !== 'success'}
                                class="bg-green-600 text-white font-bold py-1 px-4 rounded-lg hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-green-300 transition-all duration-300 text-sm disabled:opacity-50 disabled:cursor-not-allowed"
                                title=${hasSharePointId ? "Definitief opslaan" : "Nieuwe zaak opslaan"}
                            >
                                ${hasSharePointId ? "Definitief" : "Opslaan"}
                            </button>
                        </div>
                    </div>
                    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
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

                        <!-- Feitomschrijving -->
                        <div class="col-span-1 md:col-span-2 lg:col-span-3 flex flex-col">
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
                        
                        <!-- Vooronderzoek -->
                        <div class="col-span-1 md:col-span-2 lg:col-span-3 flex flex-col">
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

                        <!-- Reactie burger/gemachtigde -->
                        <div class="col-span-1 md:col-span-2 lg:col-span-3 flex flex-col">
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
                reader.onload = (e) => {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const sheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet);

                        // Map Excel data to our case format
                        const importedCases = jsonData.slice(0, 20).map((row, index) => ({
                            id: `case-${index}`,
                            sharePointId: null,
                            zaaknummer: row.Registratienummer || row.zaaknummer || '',
                            feitcode: row.Feitcode || row.feitcode || '',
                            cjibNummer: row.CJIBNummer || row.cjibNummer || '',
                            cjibLast4: (row.CJIBNummer || row.cjibNummer || '').slice(-4),
                            feitomschrijving: row.Feitomschrijving || row.feitomschrijving || '',
                            vooronderzoek: '',
                            reactie: '',
                            hearingDate: new Date().toISOString().split('T')[0],
                            startTime: '',
                            endTime: '',
                            status: 'Bezig met uitwerken',
                            isModified: true,
                        }));

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
                                feitomschrijving: '',
                                vooronderzoek: '',
                                reactie: '',
                                hearingDate: new Date().toISOString().split('T')[0],
                                startTime: '',
                                endTime: '',
                                status: 'Bezig met uitwerken',
                                isModified: false,
                            });
                        }

                        setCases(importedCases);
                        setModalContent({
                            title: 'Excel Import Voltooid',
                            message: `${Math.min(jsonData.length, 20)} zaken zijn geïmporteerd uit het Excel bestand.`
                        });
                        setShowInfoModal(true);

                    } catch (error) {
                        console.error('Error importing Excel:', error);
                        setModalContent({
                            title: 'Import Fout',
                            message: `Er is een fout opgetreden bij het importeren: ${error.message}`
                        });
                        setShowInfoModal(true);
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
                    const sharePointData = sharePointService.transformCaseToSharePoint(caseData);
                    
                    let result;
                    if (caseData.sharePointId) {
                        // Update existing item
                        result = await sharePointService.updateItem(caseData.sharePointId, sharePointData);
                    } else {
                        // Create new item
                        result = await sharePointService.createItem(sharePointData);
                        
                        // Update the case with the SharePoint ID
                        const updatedCase = { ...caseData, sharePointId: result.Id, isModified: false };
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
                    const sharePointData = sharePointService.transformCaseToSharePoint(caseData);
                    
                    // Add temporary status flag to indicate this is a work-in-progress update
                    const tempData = {
                        ...sharePointData,
                        Status: 'Bezig met uitwerken' // Force status to indicate work in progress
                    };
                    
                    await sharePointService.updateItem(caseData.sharePointId, tempData);
                    
                    // Update local state to reflect saved changes
                    const updatedCase = { ...caseData, isModified: false, status: 'Bezig met uitwerken' };
                    handleUpdateCase(index, updatedCase);
                    
                    setModalContent({
                        title: 'Tijdelijk Opgeslagen',
                        message: `Zaak #${index + 1} is tijdelijk opgeslagen. Status is ingesteld op 'Bezig met uitwerken' voor verdere bewerking.`
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
                                const sharePointData = sharePointService.transformCaseToSharePoint(caseData);
                                
                                let result;
                                if (caseData.sharePointId) {
                                    result = await sharePointService.updateItem(caseData.sharePointId, sharePointData);
                                    successes.push(`Zaak #${i + 1} bijgewerkt`);
                                } else {
                                    result = await sharePointService.createItem(sharePointData);
                                    // Update the case with the SharePoint ID
                                    const updatedCase = { ...caseData, sharePointId: result.Id, isModified: false };
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
                                <div class="flex items-center space-x-3">
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
                                        onClick=${handleSaveAll}
                                        disabled=${isLoading || connectionStatus !== 'success'}
                                        class="bg-blue-600 text-white font-bold py-2 px-6 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-300 transition-all duration-300 shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                                    >
                                        Alles Opslaan
                                    </button>
                                    <button
                                        onClick=${handleResetAll}
                                        disabled=${isLoading}
                                        class="bg-red-600 text-white font-bold py-2 px-6 rounded-lg hover:bg-red-700 focus:outline-none focus:ring-4 focus:ring-red-300 transition-all duration-300 shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                                    >
                                        Resetten
                                    </button>
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
                                    isActive=${index === activeCaseIndex}
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
