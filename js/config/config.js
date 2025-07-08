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
        displayName: 'Bedrijfsnaam',
        internalName: 'Bedrijfsnaam',
        fieldType: 'Text',
        description: 'De naam van het bedrijf (bij gemachtigde).'
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
        description: 'De starttijd van de specifieke zaakbehandeling (HH:MM formaat, bv. 14:32). Bij Excel import wordt dit automatisch gesplitst uit het "Datum en Tijd hoorzitting" veld. Bij handmatige invoer wordt de eindtijd automatisch berekend.'
    },
    {
        displayName: 'Eindtijd',
        internalName: 'EndTime',
        fieldType: 'Text',
        description: 'De eindtijd van de specifieke zaakbehandeling (HH:MM formaat, bv. 14:36). Wordt automatisch berekend als StartTime + 4 minuten bij zowel Excel import als handmatige invoer.'
    },
    {
        displayName: 'Status',
        internalName: 'Status',
        fieldType: 'Choice',
        description: 'Status van de zaak doorheen de workflow.',
        choices: ['Nieuw', 'Voorbereiding', 'In behandeling', 'Aangehouden', 'Klaarzetten voor DocGen', 'Afgehandeld']
    },
    {
        displayName: 'Gebruiker',
        internalName: 'Username',
        fieldType: 'Text',
        description: 'De naam van de medewerker die de notulen heeft ingevoerd. Automatisch gevuld.'
    }
];

// Status choices (updated workflow)
export const STATUS_CHOICES = [
    'Nieuw',
    'Voorbereiding',
    'In behandeling',
    'Aangehouden',
    'Klaarzetten voor DocGen',
    'Afgehandeld'
];

// Excel Import Field Mapping
// This section documents how Excel columns are mapped to SharePoint fields
export const EXCEL_FIELD_MAPPING = {
    'Registratienummer': 'Title', // Also accepts: 'zaaknummer', 'Zaaknummer'
    'Feitcode': 'Feitcode', // Also accepts: 'feitcode'
    'CJIB-Nummer': 'CJIBNummer', // Also accepts: 'CJIBNummer', 'cjibNummer', 'CJIB Nummer'
    'Betrokkene': 'Betrokkene', // Also accepts: 'betrokkene'
    'Eigenaar': 'Eigenaar', // Also accepts: 'eigenaar'
    'Soort': 'Soort', // Also accepts: 'soort'
    'Aantekening hoorverzoek': 'AantekeningHoorverzoek', // Also accepts: 'AantekeningHoorverzoek', 'aantekeninghoorverzoek'
    'Vooronderzoek': 'Vooronderzoek', // Also accepts: 'vooronderzoek'
    'Verslaglegger': 'Verslaglegger', // Also accepts: 'verslaglegger'
    'Bedrijfsnaam': 'Bedrijfsnaam', // Also accepts: 'bedrijfsnaam', 'Bedrijf'
    
    // Special field: Date and Time splitting
    'Datum en Tijd hoorzitting': {
        description: 'Expected format: dd-mm-yyyy hh:mm (e.g., 15-03-2024 14:30)',
        splits_to: {
            'HearingDate': 'Date part (converted to YYYY-MM-DD)',
            'StartTime': 'Time part (HH:MM format)',
            'EndTime': 'Calculated as StartTime + 4 minutes'
        },
        alternatives: ['Datum en tijd hoorzitting', 'Datum_en_Tijd_hoorzitting']
    },
    
    // Fields always set to defaults during import
    automatic_fields: {
        'Feitomschrijving': 'Set to blank during import',
        'ReactiePMBU': 'Set to blank during import (reactie)',
        'GesprokenMet': 'Set to blank during import',
        'Status': 'Set to "Nieuw"',
        'CJIBLast4': 'Auto-calculated from CJIBNummer'
    }
};
