import { EXCEL_COLUMN_MAPPING, DEFAULT_CASE_VALUES } from '../config/constants.js';
import { parseExcelDate, splitDateTimeToFields, addMinutesToTime } from './date-utils.js';

// Excel import/export utilities
export const importFromExcel = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (jsonData.length === 0) {
                    reject(new Error('Het Excel bestand is leeg'));
                    return;
                }
                
                const headers = jsonData[0];
                const rows = jsonData.slice(1);
                
                console.log('Excel headers:', headers);
                console.log('Excel rows:', rows.length);
                
                const cases = [];
                const errors = [];
                
                rows.forEach((row, index) => {
                    try {
                        // Skip empty rows
                        if (row.every(cell => !cell || cell === '')) {
                            return;
                        }
                        
                        const caseData = { ...DEFAULT_CASE_VALUES };
                        
                        headers.forEach((header, colIndex) => {
                            const mappedField = EXCEL_COLUMN_MAPPING[header];
                            if (mappedField && row[colIndex] !== undefined && row[colIndex] !== null && row[colIndex] !== '') {
                                if (mappedField === 'dateTime') {
                                    // Special handling for combined date/time field
                                    const dateTime = parseExcelDate(row[colIndex]);
                                    if (dateTime) {
                                        const { date, time } = splitDateTimeToFields(dateTime);
                                        caseData.hearingDate = date;
                                        caseData.startTime = time;
                                        // Auto-calculate end time (StartTime + 4 minutes)
                                        caseData.endTime = addMinutesToTime(time, 4);
                                    }
                                } else if (mappedField === 'hearingDate') {
                                    // Handle separate date field
                                    const date = parseExcelDate(row[colIndex]);
                                    if (date) {
                                        caseData.hearingDate = date.toISOString().split('T')[0];
                                    }
                                } else {
                                    // Convert to string and trim
                                    caseData[mappedField] = String(row[colIndex]).trim();
                                }
                            }
                        });
                        
                        // Validate required fields
                        if (!caseData.zaaknummer) {
                            errors.push(`Rij ${index + 2}: Zaaknummer is verplicht`);
                            return;
                        }
                        
                        if (!caseData.hearingDate) {
                            errors.push(`Rij ${index + 2}: Hoorzitting datum is verplicht`);
                            return;
                        }
                        
                        if (!caseData.startTime) {
                            errors.push(`Rij ${index + 2}: Start tijd is verplicht`);
                            return;
                        }
                        
                        cases.push({
                            ...caseData,
                            isFromExcel: true,
                            excelRowIndex: index + 2
                        });
                        
                    } catch (error) {
                        errors.push(`Rij ${index + 2}: ${error.message}`);
                    }
                });
                
                console.log('Imported cases:', cases);
                console.log('Import errors:', errors);
                
                resolve({ cases, errors });
                
            } catch (error) {
                reject(new Error(`Fout bij het lezen van Excel bestand: ${error.message}`));
            }
        };
        
        reader.onerror = () => {
            reject(new Error('Fout bij het lezen van het bestand'));
        };
        
        reader.readAsArrayBuffer(file);
    });
};

export const exportToExcel = (cases) => {
    try {
        // Define export headers (excluding display-only fields)
        const headers = [
            'Zaaknummer',
            'Feitcode', 
            'CJIB Nummer',
            'Betrokkene',
            'Eigenaar',
            'Soort',
            'Aantekenening Hoorverzoek',
            'Feitomschrijving',
            'Vooronderzoek',
            'Reactie PMBU',
            'Hoorzitting Datum',
            'Start Tijd',
            'Eind Tijd',
            'Verslaglegger',
            'Gesproken Met',
            'Bedrijfsnaam',
            'Status'
        ];
        
        // Map cases to export format
        const exportData = cases.map(caseData => [
            caseData.zaaknummer || '',
            caseData.feitcode || '',
            caseData.cjibNummer || '',
            caseData.betrokkene || '',
            caseData.eigenaar || '',
            caseData.soort || '',
            caseData.aantekeninghoorverzoek || '',
            caseData.feitomschrijving || '',
            caseData.vooronderzoek || '',
            caseData.reactie || '',
            caseData.hearingDate || '',
            caseData.startTime || '',
            caseData.endTime || '',
            caseData.verslaglegger || '',
            caseData.gesprokenMet || '',
            caseData.bedrijfsnaam || '',
            caseData.status || ''
        ]);
        
        // Create worksheet
        const worksheet = XLSX.utils.aoa_to_sheet([headers, ...exportData]);
        
        // Set column widths
        const columnWidths = headers.map(() => ({ wch: 15 }));
        worksheet['!cols'] = columnWidths;
        
        // Create workbook
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'PMBU Hoorzittingen');
        
        // Generate filename with current date
        const now = new Date();
        const filename = `PMBU_Hoorzittingen_${now.getFullYear()}-${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}.xlsx`;
        
        // Download file
        XLSX.writeFile(workbook, filename);
        
        return { success: true, filename };
    } catch (error) {
        console.error('Export error:', error);
        throw new Error(`Fout bij exporteren: ${error.message}`);
    }
};
