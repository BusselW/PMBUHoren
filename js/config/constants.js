// SharePoint Configuration
export const SHAREPOINT_CONFIG = {
    siteUrl: 'https://som.org.om.local/sites/MulderT/T/',
    listName: 'PMREG',
    apiUrl: 'https://som.org.om.local/sites/MulderT/T/_api/web/',
    contextApiUrl: 'https://som.org.om.local/sites/MulderT/T/_api/', // Separate URL for contextinfo
    listUrl: 'https://som.org.om.local/sites/MulderT/T/PMREG/',
    
    // Feitcode lookup configuration
    feitcodeLookup: {
        siteUrl: 'https://som.org.om.local/sites/MulderT/SBeheer/',
        apiUrl: 'https://som.org.om.local/sites/MulderT/SBeheer/_api/web/',
        listName: 'Feitcode'
    }
};

// Status choices (updated workflow)
export const STATUS_CHOICES = [
    'Nieuw',
    'Voorbereiding',
    'In behandeling',
    'Aangehouden',
    'Klaarzetten voor DocGen',
    'Afgehandeld'
];

// Field definitions and validation
export const FIELD_DEFINITIONS = {
    zaaknummer: { required: true, label: 'Zaaknummer' },
    feitcode: { required: true, label: 'Feitcode' },
    cjibNummer: { required: false, label: 'CJIB Nummer' },
    betrokkene: { required: false, label: 'Betrokkene' },
    eigenaar: { required: false, label: 'Eigenaar' },
    soort: { required: false, label: 'Soort' },
    aantekeninghoorverzoek: { required: false, label: 'Aantekenening Hoorverzoek' },
    feitomschrijving: { required: false, label: 'Feitomschrijving' },
    vooronderzoek: { required: false, label: 'Vooronderzoek' },
    reactie: { required: false, label: 'Reactie PMBU' },
    hearingDate: { required: true, label: 'Hoorzitting Datum', type: 'date' },
    startTime: { required: true, label: 'Start Tijd', type: 'time' },
    endTime: { required: true, label: 'Eind Tijd', type: 'time' },
    verslaglegger: { required: false, label: 'Verslaglegger' },
    gesprokenMet: { required: false, label: 'Gesproken Met' },
    bedrijfsnaam: { required: false, label: 'Bedrijfsnaam' },
    status: { required: true, label: 'Status', type: 'select', options: STATUS_CHOICES }
};

// Excel import column mapping
export const EXCEL_COLUMN_MAPPING = {
    'Zaaknummer': 'zaaknummer',
    'Feitcode': 'feitcode',
    'CJIB Nummer': 'cjibNummer',
    'Betrokkene': 'betrokkene',
    'Eigenaar': 'eigenaar',
    'Soort': 'soort',
    'Aantekenening Hoorverzoek': 'aantekeninghoorverzoek',
    'Feitomschrijving': 'feitomschrijving',
    'Vooronderzoek': 'vooronderzoek',
    'Reactie PMBU': 'reactie',
    'Hoorzitting Datum': 'hearingDate',
    'Datum Tijd': 'dateTime', // Special field that will be split
    'Verslaglegger': 'verslaglegger',
    'Gesproken Met': 'gesprokenMet',
    'Bedrijfsnaam': 'bedrijfsnaam'
};

// Default values for new cases
export const DEFAULT_CASE_VALUES = {
    zaaknummer: '',
    feitcode: '',
    cjibNummer: '',
    betrokkene: '',
    eigenaar: '',
    soort: '',
    aantekeninghoorverzoek: '',
    feitomschrijving: '',
    vooronderzoek: '',
    reactie: '',
    hearingDate: '',
    startTime: '',
    endTime: '',
    verslaglegger: '',
    gesprokenMet: '',
    bedrijfsnaam: '',
    status: 'Nieuw'
};
