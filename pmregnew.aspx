<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PMBU Hoorzitting Notulen - Nieuw</title>
    
    <!-- React and ReactDOM from CDN -->
    <script crossorigin src="https://unpkg.com/react@18/umd/react.development.js"></script>
    <script crossorigin src="https://unpkg.com/react-dom@18/umd/react-dom.development.js"></script>
    <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
    
    <!-- Excel library -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    
    <!-- Custom CSS -->
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
            background-color: #f5f7fa;
            color: #333;
            line-height: 1.6;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }

        .header {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }

        .header h1 {
            color: #2c3e50;
            margin-bottom: 15px;
        }

        .status-indicator {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 8px 12px;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
        }

        .status-checking {
            background: #fff3cd;
            color: #856404;
        }

        .status-success {
            background: #d1edff;
            color: #0c5460;
        }

        .status-failed {
            background: #f8d7da;
            color: #721c24;
        }

        .status-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
        }

        .status-checking .status-dot {
            background: #ffc107;
            animation: pulse 1.5s infinite;
        }

        .status-success .status-dot {
            background: #28a745;
        }

        .status-failed .status-dot {
            background: #dc3545;
        }

        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }

        .controls {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            align-items: center;
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #e9ecef;
        }

        .controls-left {
            display: flex;
            gap: 12px;
            flex: 1;
        }

        .controls-right {
            display: flex;
            gap: 12px;
        }

        .global-controls {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 6px;
            margin-top: 15px;
        }

        .global-controls-row {
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            align-items: center;
            margin-bottom: 15px;
        }

        .global-controls-row:last-child {
            margin-bottom: 0;
        }

        .control-group {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .btn {
            padding: 10px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.2s;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }

        .btn:disabled {
            opacity: 0.6;
            cursor: not-allowed;
        }

        .btn-primary {
            background: #007bff;
            color: white;
        }

        .btn-primary:hover:not(:disabled) {
            background: #0056b3;
        }

        .btn-success {
            background: #28a745;
            color: white;
        }

        .btn-success:hover:not(:disabled) {
            background: #1e7e34;
        }

        .btn-warning {
            background: #ffc107;
            color: #212529;
        }

        .btn-warning:hover:not(:disabled) {
            background: #e0a800;
        }

        .btn-danger {
            background: #dc3545;
            color: white;
        }

        .btn-danger:hover:not(:disabled) {
            background: #c82333;
        }

        .btn-secondary {
            background: #6c757d;
            color: white;
        }

        .btn-secondary:hover:not(:disabled) {
            background: #545b62;
        }

        .btn-orange {
            background: #fd7e14;
            color: white;
        }

        .btn-orange:hover:not(:disabled) {
            background: #e55a00;
        }

        .case-card {
            background: white;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 4px solid #e9ecef;
            transition: all 0.3s;
        }

        .case-card.modified {
            border-left-color: #007bff;
        }

        .case-card.active {
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }

        .case-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }

        .case-title {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .case-title h3 {
            margin: 0;
            color: #2c3e50;
        }

        .sharepoint-badge {
            background: #d1edff;
            color: #0c5460;
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 500;
            display: flex;
            align-items: center;
            gap: 4px;
        }

        .form-section {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 6px;
            margin-bottom: 20px;
        }

        .section-title {
            font-size: 12px;
            font-weight: 600;
            color: #6c757d;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 15px;
        }

        .form-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }

        .form-grid-2 {
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        }

        .form-grid-3 {
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        }

        .form-group {
            display: flex;
            flex-direction: column;
        }

        .form-group.full-width {
            grid-column: 1 / -1;
        }

        .form-label {
            font-weight: 500;
            color: #495057;
            margin-bottom: 5px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .form-control {
            padding: 10px 12px;
            border: 1px solid #ced4da;
            border-radius: 4px;
            font-size: 14px;
            transition: border-color 0.2s, box-shadow 0.2s;
        }

        .form-control:focus {
            outline: none;
            border-color: #007bff;
            box-shadow: 0 0 0 2px rgba(0, 123, 255, 0.25);
        }

        .form-control:disabled {
            background: #f8f9fa;
            color: #6c757d;
        }

        .form-control.loading {
            background: #e3f2fd;
        }

        textarea.form-control {
            resize: vertical;
            min-height: 80px;
        }

        textarea.form-control.large {
            min-height: 120px;
        }

        .readonly {
            background: #f8f9fa;
            color: #6c757d;
        }

        .loading-spinner {
            width: 16px;
            height: 16px;
            border: 2px solid #f3f3f3;
            border-top: 2px solid #007bff;
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .case-actions {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 6px;
            margin-top: 20px;
            border-top: 1px solid #e9ecef;
        }

        .actions-row {
            display: flex;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 10px;
        }

        .actions-group {
            display: flex;
            gap: 8px;
        }

        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.5);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 1000;
        }

        .modal {
            background: white;
            border-radius: 8px;
            padding: 24px;
            max-width: 500px;
            width: 90%;
            max-height: 80vh;
            overflow-y: auto;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
        }

        .modal h2 {
            margin-bottom: 16px;
            color: #2c3e50;
        }

        .modal p {
            margin-bottom: 20px;
            color: #6c757d;
            white-space: pre-line;
        }

        .modal-actions {
            display: flex;
            gap: 12px;
            justify-content: flex-end;
        }

        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.6);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 1001;
        }

        .loading-content {
            background: white;
            padding: 30px;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
        }

        .loading-content .loading-spinner {
            width: 40px;
            height: 40px;
            margin: 0 auto 15px;
        }

        .toggle {
            position: relative;
            display: inline-block;
            width: 50px;
            height: 24px;
        }

        .toggle input {
            opacity: 0;
            width: 0;
            height: 0;
        }

        .toggle-slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: #ccc;
            transition: .4s;
            border-radius: 24px;
        }

        .toggle-slider:before {
            position: absolute;
            content: "";
            height: 18px;
            width: 18px;
            left: 3px;
            bottom: 3px;
            background-color: white;
            transition: .4s;
            border-radius: 50%;
        }

        .toggle input:checked + .toggle-slider {
            background-color: #007bff;
        }

        .toggle input:checked + .toggle-slider:before {
            transform: translateX(26px);
        }

        .date-menu {
            position: relative;
            display: inline-block;
        }

        .date-dropdown {
            position: absolute;
            top: 100%;
            right: 0;
            background: white;
            border: 1px solid #e9ecef;
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            min-width: 300px;
            max-height: 400px;
            overflow-y: auto;
            z-index: 100;
            margin-top: 5px;
        }

        .date-dropdown-header {
            padding: 15px;
            border-bottom: 1px solid #e9ecef;
        }

        .date-dropdown-header h4 {
            margin: 0 0 5px 0;
            color: #2c3e50;
        }

        .date-dropdown-header p {
            margin: 0;
            color: #6c757d;
            font-size: 14px;
        }

        .date-option {
            display: block;
            width: 100%;
            padding: 12px 15px;
            border: none;
            background: none;
            text-align: left;
            cursor: pointer;
            border-bottom: 1px solid #f8f9fa;
            transition: background-color 0.2s;
        }

        .date-option:hover {
            background: #f8f9fa;
        }

        .date-option-date {
            font-weight: 500;
            color: #2c3e50;
        }

        .date-option-count {
            font-size: 14px;
            color: #6c757d;
            margin-top: 2px;
        }

        .error-boundary {
            background: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
            padding: 15px;
            border-radius: 6px;
            margin: 20px 0;
        }

        .error-boundary h3 {
            margin: 0 0 10px 0;
        }

        .file-input {
            display: none;
        }

        .text-muted {
            color: #6c757d !important;
        }

        .text-small {
            font-size: 14px;
        }

        .mb-0 {
            margin-bottom: 0 !important;
        }

        .mt-1 {
            margin-top: 8px;
        }

        @media (max-width: 768px) {
            .container {
                padding: 10px;
            }
            
            .form-grid {
                grid-template-columns: 1fr;
            }
            
            .controls {
                flex-direction: column;
                align-items: stretch;
            }
            
            .controls-left,
            .controls-right {
                width: 100%;
                justify-content: center;
            }
            
            .global-controls-row {
                flex-direction: column;
                align-items: stretch;
                gap: 10px;
            }
            
            .actions-row {
                flex-direction: column;
            }
            
            .actions-group {
                justify-content: center;
            }
        }
    </style>
</head>
<body>
    <div id="root"></div>

    <script type="text/babel">
        const { useState, useEffect, useCallback, useMemo } = React;

        // Configuration
        const SHAREPOINT_CONFIG = {
            siteUrl: 'https://som.org.om.local/sites/MulderT/T/',
            listName: 'PMREG',
            apiUrl: 'https://som.org.om.local/sites/MulderT/T/_api/web/',
            contextApiUrl: 'https://som.org.om.local/sites/MulderT/T/_api/',
            feitcodeLookup: {
                siteUrl: 'https://som.org.om.local/sites/MulderT/SBeheer/',
                apiUrl: 'https://som.org.om.local/sites/MulderT/SBeheer/_api/web/',
                listName: 'Feitcode'
            }
        };

        const STATUS_CHOICES = [
            'Nieuw',
            'Voorbereiding',
            'In behandeling',
            'Aangehouden',
            'Klaarzetten voor DocGen',
            'Afgehandeld'
        ];

        // SharePoint Service
        class SharePointService {
            constructor() {
                this.siteUrl = SHAREPOINT_CONFIG.siteUrl;
                this.listName = SHAREPOINT_CONFIG.listName;
                this.apiUrl = SHAREPOINT_CONFIG.apiUrl;
                this.contextApiUrl = SHAREPOINT_CONFIG.contextApiUrl;
                this.currentUser = null;
            }

            async getRequestDigest() {
                try {
                    const response = await fetch(`${this.contextApiUrl}contextinfo`, {
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

            async testConnection() {
                try {
                    const webResponse = await fetch(`${this.apiUrl}`, {
                        headers: {
                            'Accept': 'application/json;odata=verbose'
                        },
                        credentials: 'include'
                    });
                    
                    if (!webResponse.ok) {
                        throw new Error(`Cannot access SharePoint web: ${webResponse.status}`);
                    }

                    const listResponse = await fetch(`${this.apiUrl}lists/getbytitle('${this.listName}')`, {
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

            async getFeitomschrijvingByFeitcode(feitcode) {
                if (!feitcode || feitcode.trim() === '') {
                    return null;
                }

                try {
                    const filter = `Feitcode eq '${feitcode.replace(/'/g, "''")}'`;
                    const url = `${SHAREPOINT_CONFIG.feitcodeLookup.apiUrl}lists/getbytitle('${SHAREPOINT_CONFIG.feitcodeLookup.listName}')/items?$filter=${encodeURIComponent(filter)}&$select=Feitomschrijving&$top=1`;
                    
                    const response = await fetch(url, {
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
                        return results[0].Feitomschrijving || '';
                    }
                    
                    return null;
                } catch (error) {
                    console.error('Error fetching Feitomschrijving:', error);
                    throw error;
                }
            }

            transformCaseToSharePoint(caseData) {
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
                    CJIBNummer: caseData.cjibNummer || '',
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
                    Status: caseData.status || 'Nieuw'
                };
            }
        }

        const sharePointService = new SharePointService();

        // Helper functions
        const ensureISODate = (dateInput) => {
            if (!dateInput) return '';
            try {
                const date = new Date(dateInput);
                if (isNaN(date.getTime())) return '';
                return date.toISOString().split('T')[0];
            } catch (error) {
                console.warn('Invalid date input:', dateInput);
                return '';
            }
        };

        const calculateEndTime = (startTime) => {
            if (!startTime || !startTime.match(/^\d{1,2}:\d{2}$/)) {
                return '';
            }
            
            try {
                const [hours, minutes] = startTime.split(':').map(Number);
                
                if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
                    return '';
                }
                
                const startMinutes = hours * 60 + minutes;
                const endMinutes = startMinutes + 4;
                
                const finalMinutes = endMinutes % (24 * 60);
                const endHours = Math.floor(finalMinutes / 60);
                const endMins = finalMinutes % 60;
                
                return `${endHours.toString().padStart(2, '0')}:${endMins.toString().padStart(2, '0')}`;
            } catch (error) {
                console.warn('Error calculating end time:', error);
                return '';
            }
        };

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
                hearingDate: ensureISODate(new Date()),
                startTime: '',
                endTime: '',
                verslaglegger: '',
                gesprokenMet: '',
                bedrijfsnaam: '',
                status: 'Nieuw',
                isModified: false,
            }));
        };

        // Error Boundary Component
        class ErrorBoundary extends React.Component {
            constructor(props) {
                super(props);
                this.state = { hasError: false, error: null };
            }

            static getDerivedStateFromError(error) {
                return { hasError: true, error };
            }

            componentDidCatch(error, errorInfo) {
                console.error('Error Boundary caught an error:', error, errorInfo);
            }

            render() {
                if (this.state.hasError) {
                    return (
                        <div className="error-boundary">
                            <h3>Er is een fout opgetreden</h3>
                            <p>{this.state.error?.message || 'Onbekende fout'}</p>
                            <button 
                                className="btn btn-primary"
                                onClick={() => window.location.reload()}
                            >
                                Pagina herladen
                            </button>
                        </div>
                    );
                }

                return this.props.children;
            }
        }

        // Status Indicator Component
        const StatusIndicator = ({ status }) => {
            const getStatusConfig = () => {
                switch (status) {
                    case 'checking':
                        return { 
                            className: 'status-checking', 
                            text: 'Verbinding testen...' 
                        };
                    case 'success':
                        return { 
                            className: 'status-success', 
                            text: 'SharePoint verbonden' 
                        };
                    case 'failed':
                        return { 
                            className: 'status-failed', 
                            text: 'Verbindingsfout' 
                        };
                    default:
                        return { 
                            className: 'status-checking', 
                            text: 'Onbekend' 
                        };
                }
            };

            const config = getStatusConfig();

            return (
                <div className={`status-indicator ${config.className}`}>
                    <div className="status-dot"></div>
                    {config.text}
                </div>
            );
        };

        // Case Card Component
        const CaseCard = ({ 
            caseData, 
            index, 
            onUpdate, 
            onFocus, 
            isActive, 
            connectionStatus, 
            useGlobalGesprokenMet 
        }) => {
            const [feitcodeLookupLoading, setFeitcodeLookupLoading] = useState(false);
            const [duplicateCheckTimer, setDuplicateCheckTimer] = useState(null);

            useEffect(() => {
                return () => {
                    if (duplicateCheckTimer) {
                        clearTimeout(duplicateCheckTimer);
                    }
                };
            }, [duplicateCheckTimer]);

            const handleInputChange = async (e) => {
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

                // Auto-lookup Feitomschrijving when Feitcode changes
                if (name === 'feitcode' && value && value.trim() !== '') {
                    const currentValue = value.trim();
                    
                    // Clear existing timer and set new one
                    if (duplicateCheckTimer) {
                        clearTimeout(duplicateCheckTimer);
                    }
                    
                    const newTimer = setTimeout(async () => {
                        try {
                            setFeitcodeLookupLoading(true);
                            const feitomschrijving = await sharePointService.getFeitomschrijvingByFeitcode(currentValue);
                            if (feitomschrijving && caseData.feitcode === currentValue) {
                                const updatedCaseWithFeitomschrijving = {
                                    ...updatedData,
                                    feitomschrijving: feitomschrijving,
                                    isModified: true
                                };
                                onUpdate(index, updatedCaseWithFeitomschrijving);
                                return; // Don't call onUpdate again below
                            }
                        } catch (error) {
                            console.warn('Failed to lookup Feitomschrijving:', error);
                        } finally {
                            setFeitcodeLookupLoading(false);
                        }
                    }, 500);
                    
                    setDuplicateCheckTimer(newTimer);
                } else if (name === 'feitcode' && (!value || value.trim() === '')) {
                    // Clear feitomschrijving when feitcode is cleared
                    updatedData.feitomschrijving = '';
                }
                
                onUpdate(index, updatedData);
            };

            const handleFocus = () => {
                onFocus(index);
            };

            const hasSharePointId = caseData.sharePointId !== null && caseData.sharePointId !== undefined;

            return (
                <div 
                    className={`case-card ${caseData.isModified ? 'modified' : ''} ${isActive ? 'active' : ''}`}
                    onClick={handleFocus}
                >
                    <div className="case-header">
                        <div className="case-title">
                            <h3>Zaak #{index + 1}</h3>
                            {hasSharePointId && (
                                <div className="sharepoint-badge">
                                    ✓ SharePoint
                                </div>
                            )}
                        </div>
                    </div>

                    {/* Basic Case Information */}
                    <div className="form-section">
                        <div className="section-title">Zaak Informatie</div>
                        <div className="form-grid form-grid-3">
                            <div className="form-group">
                                <label className="form-label">Zaaknummer</label>
                                <input
                                    type="text"
                                    name="zaaknummer"
                                    value={caseData.zaaknummer || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                    placeholder="bv. 123456789"
                                />
                            </div>
                            <div className="form-group">
                                <label className="form-label">CJIB Nummer</label>
                                <input
                                    type="text"
                                    name="cjibNummer"
                                    value={caseData.cjibNummer || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                    placeholder="Volledig CJIB nummer"
                                />
                            </div>
                            <div className="form-group">
                                <label className="form-label">CJIB Laatste 4</label>
                                <input
                                    type="text"
                                    name="cjibLast4"
                                    value={caseData.cjibLast4 || ''}
                                    className="form-control readonly"
                                    placeholder="Auto"
                                    readOnly
                                />
                            </div>
                        </div>
                    </div>

                    {/* Party Information */}
                    <div className="form-section">
                        <div className="section-title">Betrokkenen</div>
                        <div className="form-grid form-grid-3">
                            <div className="form-group">
                                <label className="form-label">Betrokkene</label>
                                <input
                                    type="text"
                                    name="betrokkene"
                                    value={caseData.betrokkene || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                    placeholder="Naam van de betrokkene"
                                />
                            </div>
                            <div className="form-group">
                                <label className="form-label">Eigenaar</label>
                                <input
                                    type="text"
                                    name="eigenaar"
                                    value={caseData.eigenaar || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                    placeholder="Naam van de eigenaar"
                                />
                            </div>
                            <div className="form-group">
                                <label className="form-label">Soort</label>
                                <input
                                    type="text"
                                    name="soort"
                                    value={caseData.soort || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                    placeholder="Soort zaak/overtreding"
                                />
                            </div>
                        </div>
                    </div>

                    {/* Timing Information */}
                    <div className="form-section">
                        <div className="section-title">Tijd en Datum</div>
                        <div className="form-grid form-grid-3">
                            <div className="form-group">
                                <label className="form-label">Datum Hoorzitting</label>
                                <input
                                    type="date"
                                    name="hearingDate"
                                    value={caseData.hearingDate || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                />
                            </div>
                            <div className="form-group">
                                <label className="form-label">Starttijd</label>
                                <input
                                    type="time"
                                    name="startTime"
                                    value={caseData.startTime || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                />
                            </div>
                            <div className="form-group">
                                <label className="form-label">Eindtijd</label>
                                <input
                                    type="time"
                                    name="endTime"
                                    value={caseData.endTime || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                />
                            </div>
                        </div>
                    </div>

                    {/* Violation Information */}
                    <div className="form-section">
                        <div className="section-title">Overtreding Details</div>
                        <div className="form-grid form-grid-2">
                            <div className="form-group">
                                <label className="form-label">Feitcode</label>
                                <input
                                    type="text"
                                    name="feitcode"
                                    value={caseData.feitcode || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                    placeholder="bv. R584"
                                />
                            </div>
                            <div className="form-group">
                                <label className="form-label">Status</label>
                                <select
                                    name="status"
                                    value={caseData.status || 'Nieuw'}
                                    onChange={handleInputChange}
                                    className="form-control"
                                >
                                    {STATUS_CHOICES.map(choice => (
                                        <option key={choice} value={choice}>{choice}</option>
                                    ))}
                                </select>
                            </div>
                            <div className="form-group full-width">
                                <label className="form-label">
                                    Feitomschrijving
                                    {feitcodeLookupLoading && (
                                        <div className="loading-spinner"></div>
                                    )}
                                </label>
                                <input
                                    type="text"
                                    name="feitomschrijving"
                                    value={caseData.feitomschrijving || ''}
                                    onChange={handleInputChange}
                                    className={`form-control ${feitcodeLookupLoading ? 'loading' : ''}`}
                                    placeholder="Omschrijving van de overtreding (wordt automatisch ingevuld bij Feitcode)"
                                    disabled={feitcodeLookupLoading}
                                />
                            </div>
                        </div>
                    </div>

                    {/* Investigation & Communication */}
                    <div className="form-section">
                        <div className="section-title">Onderzoek en Communicatie</div>
                        
                        <div className="form-group">
                            <label className="form-label">Vooronderzoek</label>
                            <textarea
                                name="vooronderzoek"
                                value={caseData.vooronderzoek || ''}
                                onChange={handleInputChange}
                                className="form-control"
                                placeholder="Resultaten van het vooronderzoek..."
                            />
                        </div>

                        {!useGlobalGesprokenMet && (
                            <div className="form-group">
                                <label className="form-label">Gesproken Met</label>
                                <input
                                    type="text"
                                    name="gesprokenMet"
                                    value={caseData.gesprokenMet || ''}
                                    onChange={handleInputChange}
                                    className="form-control"
                                    placeholder="Met wie is er gesproken?"
                                />
                            </div>
                        )}

                        <div className="form-group">
                            <label className="form-label">Reactie burger/gemachtigde</label>
                            <textarea
                                name="reactie"
                                value={caseData.reactie || ''}
                                onChange={handleInputChange}
                                className="form-control large"
                                placeholder="Noteer hier het gesprek..."
                            />
                        </div>

                        <div className="form-group">
                            <label className="form-label">Aantekening Hoorverzoek</label>
                            <textarea
                                name="aantekeninghoorverzoek"
                                value={caseData.aantekeninghoorverzoek || ''}
                                onChange={handleInputChange}
                                className="form-control"
                                placeholder="Aantekeningen betreffende het hoorverzoek..."
                            />
                        </div>
                    </div>

                    {/* Case Actions */}
                    <div className="case-actions">
                        <div className="actions-row">
                            <div className="actions-group">
                                <button 
                                    className="btn btn-orange"
                                    disabled={connectionStatus !== 'success'}
                                    title="Tijdelijk opslaan - status wordt 'In behandeling'"
                                >
                                    Tijdelijk Opslaan
                                </button>
                                <button 
                                    className="btn btn-warning"
                                    disabled={connectionStatus !== 'success'}
                                    title="Klaarzetten voor DocGen"
                                >
                                    Klaarzetten DocGen
                                </button>
                                <button 
                                    className="btn btn-success"
                                    disabled={connectionStatus !== 'success'}
                                    title="Definitief afhandelen - status wordt 'Afgehandeld'"
                                >
                                    Definitief
                                </button>
                            </div>
                            <div className="actions-group">
                                {hasSharePointId && (
                                    <button 
                                        className="btn btn-secondary"
                                        disabled={connectionStatus !== 'success'}
                                        title="Oude tijdelijke opslag (behoudt huidige status)"
                                    >
                                        Temp. Opslaan (Legacy)
                                    </button>
                                )}
                                <button 
                                    className="btn btn-primary"
                                    disabled={connectionStatus !== 'success'}
                                    title={hasSharePointId ? "Opslaan zonder status wijziging" : "Nieuwe zaak opslaan"}
                                >
                                    {hasSharePointId ? "Opslaan" : "Nieuwe Zaak"}
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            );
        };

        // Modal Components
        const Modal = ({ isOpen, onClose, title, children }) => {
            if (!isOpen) return null;

            return (
                <div className="modal-overlay" onClick={onClose}>
                    <div className="modal" onClick={e => e.stopPropagation()}>
                        <h2>{title}</h2>
                        {children}
                    </div>
                </div>
            );
        };

        const InfoModal = ({ isOpen, onClose, title, message }) => (
            <Modal isOpen={isOpen} onClose={onClose} title={title}>
                <p>{message}</p>
                <div className="modal-actions">
                    <button className="btn btn-primary" onClick={onClose}>
                        Sluiten
                    </button>
                </div>
            </Modal>
        );

        const ConfirmModal = ({ isOpen, onClose, onConfirm, title, message }) => (
            <Modal isOpen={isOpen} onClose={onClose} title={title}>
                <p>{message}</p>
                <div className="modal-actions">
                    <button className="btn btn-secondary" onClick={onClose}>
                        Annuleren
                    </button>
                    <button className="btn btn-danger" onClick={onConfirm}>
                        Bevestigen
                    </button>
                </div>
            </Modal>
        );

        const LoadingOverlay = ({ isVisible, message = "Bezig met opslaan..." }) => {
            if (!isVisible) return null;

            return (
                <div className="loading-overlay">
                    <div className="loading-content">
                        <div className="loading-spinner"></div>
                        <p>{message}</p>
                    </div>
                </div>
            );
        };

        // Toggle Component
        const Toggle = ({ checked, onChange, id }) => (
            <label className="toggle" htmlFor={id}>
                <input 
                    type="checkbox" 
                    id={id} 
                    checked={checked} 
                    onChange={onChange} 
                />
                <span className="toggle-slider"></span>
            </label>
        );

        // Main App Component
        const App = () => {
            const [cases, setCases] = useState(() => createInitialCases(20));
            const [activeCaseIndex, setActiveCaseIndex] = useState(0);
            const [connectionStatus, setConnectionStatus] = useState('checking');
            const [isLoading, setIsLoading] = useState(false);
            
            // Modal states
            const [showInfoModal, setShowInfoModal] = useState(false);
            const [showConfirmModal, setShowConfirmModal] = useState(false);
            const [modalContent, setModalContent] = useState({ title: '', message: '' });
            
            // Global form states
            const [globalVerslaglegger, setGlobalVerslaglegger] = useState('');
            const [globalGesprokenMet, setGlobalGesprokenMet] = useState('');
            const [useGlobalGesprokenMet, setUseGlobalGesprokenMet] = useState(true);
            const [isGemachtigde, setIsGemachtigde] = useState(true);
            const [globalBedrijfsnaam, setGlobalBedrijfsnaam] = useState('');
            
            // Date menu states
            const [showDateMenu, setShowDateMenu] = useState(false);
            const [availableDates, setAvailableDates] = useState([]);
            const [loadingDates, setLoadingDates] = useState(false);

            // Test SharePoint connection on load
            useEffect(() => {
                const testConnection = async () => {
                    try {
                        console.log('Testing SharePoint connection...');
                        await sharePointService.testConnection();
                        setTimeout(() => {
                            setConnectionStatus('success');
                            console.log('SharePoint connection test successful');
                        }, 100);
                    } catch (error) {
                        setTimeout(() => {
                            setConnectionStatus('failed');
                            console.error('SharePoint connection test failed:', error);
                            setModalContent({
                                title: 'SharePoint Verbindingsfout',
                                message: `Kan geen verbinding maken met SharePoint: ${error.message}`
                            });
                            setShowInfoModal(true);
                        }, 100);
                    }
                };
                
                setTimeout(testConnection, 200);
            }, []);

            // Close date menu when clicking outside
            useEffect(() => {
                const handleClickOutside = (event) => {
                    if (showDateMenu && !event.target.closest('.date-menu')) {
                        setShowDateMenu(false);
                    }
                };

                if (showDateMenu) {
                    document.addEventListener('mousedown', handleClickOutside);
                    return () => document.removeEventListener('mousedown', handleClickOutside);
                }
            }, [showDateMenu]);

            // Scroll to active card
            useEffect(() => {
                const timeoutId = setTimeout(() => {
                    try {
                        const activeCard = document.querySelector('.case-card.active');
                        if (activeCard && typeof activeCard.scrollIntoView === 'function') {
                            activeCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        }
                    } catch (error) {
                        console.warn('Error scrolling to active card:', error);
                    }
                }, 100);
                
                return () => clearTimeout(timeoutId);
            }, [activeCaseIndex]);

            const handleUpdateCase = useCallback((index, updatedCase) => {
                setCases(prevCases => {
                    const newCases = [...prevCases];
                    newCases[index] = updatedCase;
                    return newCases;
                });
            }, []);

            const handleFocusCase = useCallback((index) => {
                setActiveCaseIndex(index);
            }, []);

            const handleExcelImport = (event) => {
                const file = event.target.files[0];
                if (!file) return;

                const reader = new FileReader();
                reader.onload = async (e) => {
                    setIsLoading(true);
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const sheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet);

                        // Transform Excel data to case format
                        const importedCases = jsonData.slice(0, 20).map((row, index) => ({
                            id: `case-${index}`,
                            sharePointId: null,
                            zaaknummer: row.Registratienummer || row.zaaknummer || row.Zaaknummer || '',
                            feitcode: row.Feitcode || row.feitcode || '',
                            cjibNummer: row['CJIB-Nummer'] || row.CJIBNummer || row.cjibNummer || '',
                            cjibLast4: (row['CJIB-Nummer'] || row.CJIBNummer || row.cjibNummer || '').slice(-4),
                            betrokkene: row.Betrokkene || row.betrokkene || '',
                            eigenaar: row.Eigenaar || row.eigenaar || '',
                            soort: row.Soort || row.soort || '',
                            aantekeninghoorverzoek: row['Aantekening hoorverzoek'] || row.AantekeningHoorverzoek || '',
                            feitomschrijving: '',
                            vooronderzoek: row.Vooronderzoek || row.vooronderzoek || '',
                            reactie: '',
                            hearingDate: ensureISODate(new Date()),
                            startTime: '',
                            endTime: '',
                            verslaglegger: '',
                            gesprokenMet: '',
                            bedrijfsnaam: row.Bedrijfsnaam || row.bedrijfsnaam || '',
                            status: 'Nieuw',
                            isModified: false,
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
                    } finally {
                        setIsLoading(false);
                    }
                };
                reader.readAsArrayBuffer(file);
                
                // Reset file input
                event.target.value = '';
            };

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

            return (
                <ErrorBoundary>
                    <div className="container">
                        {/* Header */}
                        <div className="header">
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
                                    <h1>PMBU Hoorzitting Notulen</h1>
                                    <StatusIndicator status={connectionStatus} />
                                </div>
                            </div>

                            {/* Controls */}
                            <div className="controls">
                                <div className="controls-left">
                                    <div className="date-menu">
                                        <button 
                                            className="btn btn-primary"
                                            disabled={isLoading || connectionStatus !== 'success'}
                                            onClick={() => setShowDateMenu(!showDateMenu)}
                                        >
                                            📅 Laden per Datum
                                        </button>
                                        {showDateMenu && (
                                            <div className="date-dropdown">
                                                <div className="date-dropdown-header">
                                                    <h4>Beschikbare Datums</h4>
                                                    <p>Selecteer een datum om onafgeronde zaken te laden</p>
                                                </div>
                                                {loadingDates ? (
                                                    <div style={{ padding: '20px', textAlign: 'center' }}>
                                                        <div className="loading-spinner" style={{ margin: '0 auto 10px' }}></div>
                                                        <p>Datums laden...</p>
                                                    </div>
                                                ) : availableDates.length === 0 ? (
                                                    <div style={{ padding: '20px', textAlign: 'center' }}>
                                                        <p className="text-muted">Geen onafgeronde zaken gevonden</p>
                                                    </div>
                                                ) : (
                                                    availableDates.map(dateInfo => (
                                                        <button
                                                            key={dateInfo.date}
                                                            className="date-option"
                                                            onClick={() => console.log('Load date:', dateInfo.date)}
                                                        >
                                                            <div className="date-option-date">{dateInfo.displayDate}</div>
                                                            <div className="date-option-count">
                                                                {dateInfo.count} {dateInfo.count === 1 ? 'zaak' : 'zaken'} te voltooien
                                                            </div>
                                                        </button>
                                                    ))
                                                )}
                                                <button 
                                                    className="date-option"
                                                    onClick={() => setShowDateMenu(false)}
                                                    style={{ borderTop: '1px solid #e9ecef', background: '#f8f9fa' }}
                                                >
                                                    Sluiten
                                                </button>
                                            </div>
                                        )}
                                    </div>
                                    
                                    <input
                                        type="file"
                                        accept=".xlsx,.xls"
                                        onChange={handleExcelImport}
                                        className="file-input"
                                        id="excel-import"
                                    />
                                    <button
                                        className="btn btn-secondary"
                                        disabled={isLoading || connectionStatus !== 'success'}
                                        onClick={() => document.getElementById('excel-import').click()}
                                    >
                                        📊 Excel Import
                                    </button>
                                    
                                    <button
                                        className="btn btn-danger"
                                        disabled={isLoading}
                                        onClick={handleResetAll}
                                    >
                                        🔄 Resetten
                                    </button>
                                </div>

                                <div className="controls-right">
                                    <button 
                                        className="btn btn-orange"
                                        disabled={isLoading || connectionStatus !== 'success'}
                                    >
                                        Alles Tijdelijk Opslaan
                                    </button>
                                    <button 
                                        className="btn btn-warning"
                                        disabled={isLoading || connectionStatus !== 'success'}
                                    >
                                        Alles Klaarzetten DocGen
                                    </button>
                                    <button 
                                        className="btn btn-success"
                                        disabled={isLoading || connectionStatus !== 'success'}
                                    >
                                        Alles Definitief
                                    </button>
                                </div>
                            </div>

                            {/* Global Controls */}
                            <div className="global-controls">
                                <div className="global-controls-row">
                                    <div className="control-group">
                                        <label htmlFor="global-verslaglegger" className="form-label">Verslaglegger:</label>
                                        <input
                                            type="text"
                                            id="global-verslaglegger"
                                            value={globalVerslaglegger}
                                            onChange={(e) => setGlobalVerslaglegger(e.target.value)}
                                            className="form-control"
                                            style={{ width: '200px' }}
                                            placeholder="Naam van de verslaglegger"
                                        />
                                    </div>

                                    <div className="control-group">
                                        <label className="form-label">Gesproken Met:</label>
                                        <Toggle
                                            id="gesproken-met-toggle"
                                            checked={useGlobalGesprokenMet}
                                            onChange={(e) => setUseGlobalGesprokenMet(e.target.checked)}
                                        />
                                        <span className="text-small text-muted">
                                            {useGlobalGesprokenMet ? 'Globaal' : 'Per zaak'}
                                        </span>
                                    </div>

                                    {useGlobalGesprokenMet && (
                                        <div className="control-group">
                                            <input
                                                type="text"
                                                value={globalGesprokenMet}
                                                onChange={(e) => setGlobalGesprokenMet(e.target.value)}
                                                className="form-control"
                                                style={{ width: '200px' }}
                                                placeholder="Met wie gesproken"
                                            />
                                        </div>
                                    )}
                                </div>

                                <div className="global-controls-row">
                                    <div className="control-group">
                                        <label className="form-label">Type:</label>
                                        <Toggle
                                            id="type-toggle"
                                            checked={isGemachtigde}
                                            onChange={(e) => setIsGemachtigde(e.target.checked)}
                                        />
                                        <span className="text-small text-muted">
                                            {isGemachtigde ? 'Gemachtigde' : 'Burger'}
                                        </span>
                                    </div>

                                    {isGemachtigde && (
                                        <div className="control-group">
                                            <label className="form-label">Bedrijfsnaam:</label>
                                            <input
                                                type="text"
                                                value={globalBedrijfsnaam}
                                                onChange={(e) => setGlobalBedrijfsnaam(e.target.value)}
                                                className="form-control"
                                                style={{ width: '200px' }}
                                                placeholder="Naam van het bedrijf"
                                            />
                                        </div>
                                    )}
                                </div>
                            </div>
                        </div>

                        {/* Cases */}
                        <div>
                            {cases.map((caseItem, index) => (
                                <CaseCard
                                    key={caseItem.id}
                                    caseData={caseItem}
                                    index={index}
                                    onUpdate={handleUpdateCase}
                                    onFocus={handleFocusCase}
                                    isActive={index === activeCaseIndex}
                                    connectionStatus={connectionStatus}
                                    useGlobalGesprokenMet={useGlobalGesprokenMet}
                                />
                            ))}
                        </div>

                        {/* Modals */}
                        <InfoModal
                            isOpen={showInfoModal}
                            onClose={closeInfoModal}
                            title={modalContent.title}
                            message={modalContent.message}
                        />

                        <ConfirmModal
                            isOpen={showConfirmModal}
                            onClose={cancelReset}
                            onConfirm={confirmReset}
                            title="Weet u het zeker?"
                            message="Hiermee worden alle gegevens op de pagina gewist. Deze actie kan niet ongedaan worden gemaakt."
                        />

                        <LoadingOverlay isVisible={isLoading} />
                    </div>
                </ErrorBoundary>
            );
        };

        // Render the app
        const container = document.getElementById('root');
        const root = ReactDOM.createRoot(container);
        root.render(<App />);
    </script>
</body>
</html>
