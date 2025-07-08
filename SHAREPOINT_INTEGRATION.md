# PMREG SharePoint Integration Documentation

## Overview
This application integrates with SharePoint to manage hearing notes (Hoorzitting Notulen) for the PMBU department.

## SharePoint Configuration

### Site Details
- **Site URL**: https://som.org.om.local/sites/MulderT/T/
- **List Name**: PMREG
- **Library**: PMREG
- **API Endpoint**: https://som.org.om.local/sites/MulderT/T/_api/web/

### SharePoint List Fields

| Display Name (Dutch) | Internal Name | SharePoint Field Type | Description |
|----------------------|---------------|----------------------|-------------|
| Titel | Title | Single line of text | Standard SharePoint field. Used for case number (Zaaknummer) |
| Feitcode | Feitcode | Single line of text | Code corresponding to the violation |
| Feitomschrijving | Feitomschrijving | Multiple lines of text | Full description of the violation |
| Vooronderzoek | Vooronderzoek | Multiple lines of text | Notes and findings from pre-hearing investigation |
| Reactie PMBU | ReactiePMBU | Multiple lines of text | Literal response or summary of conversation with citizen/representative |
| Datum Hoorzitting | HearingDate | Date and time | Date when the hearing takes place |
| Starttijd | StartTime | Single line of text | Start time of specific case handling (e.g., 14:32) |
| Eindtijd | EndTime | Single line of text | End time of specific case handling (e.g., 14:38) |
| Status | Status | Choice | Case status with predefined options |
| Gebruiker | Username | Single line of text | Name of the employee who entered the notes (auto-filled) |

### Status Field Choices (Updated to match SharePoint)
- Bezig met uitwerken
- Aangehouden
- Afgerond  <!-- Changed from 'Afgehandeld' to match SharePoint exactly -->

## Features

### Connection Status
- **Real-time connection indicator** in the header
- **Green dot**: SharePoint connection successful
- **Yellow dot**: Testing connection
- **Red dot**: Connection failed
- **Smart button disabling**: Save functions disabled when connection fails

### Individual Case Management
- Each case card has its own "Opslaan" (Save) button for final saves
- **"Temp. Opslaan" (Temporary Save)** button for work-in-progress updates
- Cases can be saved individually to SharePoint
- Real-time feedback on save operations

### Temporary Save Feature
- **Orange "Temp. Opslaan" button** appears only for existing SharePoint cases
- Allows updating case data before the hearing without finalizing
- Automatically sets status to "Bezig met uitwerken" to indicate work in progress
- Perfect for pre-hearing preparation and data updates
- Only available for cases that have been saved to SharePoint at least once

### Bulk Operations
- "Alles Opslaan" (Save All) button saves all modified cases
- Only cases with data or modifications are saved
- Comprehensive error handling and reporting

### User Management
- Automatically fetches current SharePoint user information
- Auto-populates Username field when saving

### Time Management
- Start and End time fields use HTML5 time pickers
- Format: HH:MM (24-hour format)
- Stored as text in SharePoint for flexibility

### Status Management
- Dropdown selection with predefined choices
- Default status: "Bezig met uitwerken"

## Usage

### Creating New Cases
1. Fill in case details in the form fields
2. Use the green "Opslaan" button to create a new case in SharePoint
3. Cases are automatically assigned a SharePoint ID upon creation
4. Once saved, the "Temp. Opslaan" button becomes available

### Updating Existing Cases
1. Modify any field in an existing case
2. The case border turns blue to indicate modifications
3. **For temporary updates**: Use orange "Temp. Opslaan" button
   - Updates data in SharePoint immediately
   - Sets status to "Bezig met uitwerken" automatically
   - Perfect for pre-hearing data preparation
4. **For final updates**: Use green "Definitief" button
   - Saves with current status selection
   - Use when case handling is complete

### Temporary Save Workflow
1. **Initial Save**: Create case with green "Opslaan" button
2. **Pre-hearing Updates**: Use orange "Temp. Opslaan" for ongoing changes
3. **Final Save**: Use green "Definitief" when case is completed
4. **Status Management**: Temp saves always mark as "Bezig met uitwerken"

### Time Entry
- Use the time picker controls for Start and End times
- Format is automatically handled (HH:MM)
- Leave empty if times are not applicable

### Status Updates
- Use the Status dropdown to change case status
- Default is "Bezig met uitwerken" for new cases

## Error Handling
- Network errors are caught and displayed to the user
- Individual case save errors don't prevent other cases from saving
- Detailed error messages help troubleshoot issues

## Security
- Uses SharePoint's built-in authentication
- All API calls include credentials for authentication
- Request digest tokens are automatically managed

## Files Structure
```
/
├── pmreg.aspx                 # Main application file
├── js/
│   ├── config/
│   │   └── config.js         # Configuration settings
│   └── sharepoint-service.js # SharePoint CRUD operations
└── README.md                 # This documentation
```

## Browser Compatibility
- Modern browsers with ES6 module support
- Requires JavaScript enabled
- SharePoint authentication cookies required
