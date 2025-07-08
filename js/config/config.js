// SharePoint Configuration
export const SHAREPOINT_CONFIG = {
    siteUrl: 'https://som.org.om.local/sites/MulderT/T/',
    listName: 'PMREG',
    apiUrl: 'https://som.org.om.local/sites/MulderT/T/_api/web/',
    contextApiUrl: 'https://som.org.om.local/sites/MulderT/T/_api/', // Separate URL for contextinfo
    listUrl: 'https://som.org.om.local/sites/MulderT/T/PMREG/',
};

// Field mapping configuration
export const FIELD_CONFIG = [
    {
        displayName: 'Titel',
        internalName: 'Title',
        fieldType: 'Text',
        description: 'Standaard SharePoint-veld. Je kunt dit gebruiken voor het Zaaknummer.'
    },
    {
        displayName: 'Feitcode',
        internalName: 'Feitcode',
        fieldType: 'Text',
        description: 'De code die correspondeert met de overtreding.'
    },
    {
        displayName: 'CJIB Nummer',
        internalName: 'CJIBNummer',
        fieldType: 'Text',
        description: 'Het volledige CJIB nummer van de zaak.'
    },
    {
        displayName: 'CJIB Laatste 4',
        internalName: 'CJIBLast4',
        fieldType: 'Text',
        description: 'De laatste 4 cijfers van het CJIB nummer (automatisch gegenereerd).'
    },
    {
        displayName: 'Betrokkene',
        internalName: 'Betrokkene',
        fieldType: 'Text',
        description: 'De betrokkene bij de zaak.'
    },
    {
        displayName: 'Eigenaar',
        internalName: 'Eigenaar',
        fieldType: 'Text',
        description: 'De eigenaar van het voertuig of object in de zaak.'
    },
    {
        displayName: 'Soort',
        internalName: 'Soort',
        fieldType: 'Text',
        description: 'Het soort zaak of overtreding.'
    },
    {
        displayName: 'Aantekening Hoorverzoek',
        internalName: 'AantekeningHoorverzoek',
        fieldType: 'Note',
        description: 'Aantekeningen betreffende het hoorverzoek.'
    },
    {
        displayName: 'Verslaglegger',
        internalName: 'Verslaglegger',
        fieldType: 'Text',
        description: 'De naam van de persoon die het verslag opstelt.'
    },
    {
        displayName: 'Gesproken Met',
        internalName: 'GesprokenMet',
        fieldType: 'Text',
        description: 'Met wie er is gesproken tijdens de hoorzitting.'
    },
    {
        displayName: 'Feitomschrijving',
        internalName: 'Feitomschrijving',
        fieldType: 'Note',
        description: 'De volledige omschrijving van de overtreding.'
    },
    {
        displayName: 'Vooronderzoek',
        internalName: 'Vooronderzoek',
        fieldType: 'Note',
        description: 'Notities en bevindingen van het onderzoek voorafgaand aan de hoorzitting.'
    },
    {
        displayName: 'Reactie PMBU',
        internalName: 'ReactiePMBU',
        fieldType: 'Note',
        description: 'De letterlijke reactie of de samenvatting van het gesprek met de burger/gemachtigde.'
    },
    {
        displayName: 'Datum Hoorzitting',
        internalName: 'HearingDate',
        fieldType: 'DateTime',
        description: 'De datum waarop de hoorzitting plaatsvindt. Handig voor sorteren en filteren.'
    },
    {
        displayName: 'Starttijd',
        internalName: 'StartTime',
        fieldType: 'Text',
        description: 'De starttijd van de specifieke zaakbehandeling (bv. 14:32).'
    },
    {
        displayName: 'Eindtijd',
        internalName: 'EndTime',
        fieldType: 'Text',
        description: 'De eindtijd van de specifieke zaakbehandeling (bv. 14:38).'
    },
    {
        displayName: 'Status',
        internalName: 'Status',
        fieldType: 'Choice',
        description: 'Status van de zaak (bijv. \'Nieuw\', \'In behandeling\', \'Afgerond\').',
        choices: ['Bezig met uitwerken', 'Aangehouden', 'Afgerond']  // Updated to match SharePoint
    },
    {
        displayName: 'Gebruiker',
        internalName: 'Username',
        fieldType: 'Text',
        description: 'De naam van de medewerker die de notulen heeft ingevoerd. Automatisch gevuld.'
    }
];

// Status choices (matching SharePoint exactly)
export const STATUS_CHOICES = [
    'Bezig met uitwerken',
    'Aangehouden',
    'Afgerond'  // Changed from 'Afgehandeld' to match SharePoint
];
