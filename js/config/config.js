const SHAREPOINT_CONFIG = {
    siteUrl: 'http://som.org.om.local/sites/MulderT/',
    listName: 'PMREG',
    apiUrl: 'http://som.org.om.local/sites/MulderT/_api/web/',
    contextApiUrl: 'http://som.org.om.local/sites/MulderT/_api/', // Separate URL for contextinfo
    listUrl: 'http://som.org.om.local/sites/MulderT/PMREG/',
    
    // Feitcode lookup configuration
    feitcodeLookup: {
        siteUrl: 'http://som.org.om.local/sites/MulderT/SBeheer/',
        apiUrl: 'http://som.org.om.local/sites/MulderT/SBeheer/_api/web/',
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

export { SHAREPOINT_CONFIG, STATUS_CHOICES };