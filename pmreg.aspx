<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hoorzitting Notulen Logger</title>
    <script src="https://cdn.tailwindcss.com"></script>
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
            listUrl: 'https://som.org.om.local/sites/MulderT/T/PMREG/',
        };

        // Status choices
        const STATUS_CHOICES = [
            'Bezig met uitwerken',
            'Aangehouden',
            'Afgehandeld'
        ];

        // SharePoint Service Class
        class SharePointService {
            constructor() {
                this.siteUrl = SHAREPOINT_CONFIG.siteUrl;
                this.listName = SHAREPOINT_CONFIG.listName;
                this.apiUrl = SHAREPOINT_CONFIG.apiUrl;
                this.currentUser = null;
            }

            async getRequestDigest() {
                try {
                    const response = await fetch(`${this.apiUrl}contextinfo`, {
                        method: 'POST',
                        headers: {
                            'Accept': 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose'
                        },
                        credentials: 'include'
                    });
                    
                    if (!response.ok) {
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
                feitomschrijving: '',
                vooronderzoek: '',
                reactie: '',
                hearingDate: new Date().toISOString().split('T')[0], // Today's date
                startTime: '',
                endTime: '',
                status: 'Bezig met uitwerken',
                isModified: false,
            }));
        };

        // --- CaseCard Component ---
        // Represents a single case with its input fields.
        const CaseCard = ({ caseData, index, onUpdate, onFocus, isActive, onSaveIndividual }) => {
            const { id, zaaknummer, feitcode, feitomschrijving, vooronderzoek, reactie, hearingDate, startTime, endTime, status, isModified } = caseData;

            const handleInputChange = (e) => {
                const { name, value } = e.target;
                onUpdate(index, { ...caseData, [name]: value, isModified: true });
            };

            const handleFocus = () => {
                onFocus(index);
            };

            const handleSaveCase = () => {
                onSaveIndividual(index);
            };

            const cardBorderColor = isModified ? 'border-blue-500' : 'border-gray-200';
            const activeShadow = isActive ? 'shadow-xl' : 'shadow-md';

            return html`
                <div 
                    id=${`case-card-${index}`}
                    class="bg-white p-6 rounded-lg border-l-4 ${cardBorderColor} ${activeShadow} transition-all duration-300 mb-4"
                >
                    <div class="flex justify-between items-center mb-4">
                        <h3 class="text-xl font-bold text-gray-700">Zaak #${index + 1}</h3>
                        <button
                            onClick=${handleSaveCase}
                            class="bg-green-600 text-white font-bold py-1 px-4 rounded-lg hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-green-300 transition-all duration-300 text-sm"
                        >
                            Opslaan
                        </button>
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
                                <h1 class="text-3xl font-bold text-gray-800">Hoorzitting Notulen</h1>
                                <div class="flex items-center space-x-3">
                                    <button
                                        onClick=${handleSaveAll}
                                        disabled=${isLoading}
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
