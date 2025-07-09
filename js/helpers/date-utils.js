// Date and time utility functions
export const formatDateForDisplay = (dateStr) => {
    if (!dateStr) return '';
    try {
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return '';
        return date.toLocaleDateString('nl-NL');
    } catch {
        return '';
    }
};

export const formatDateForInput = (dateStr) => {
    if (!dateStr) return '';
    try {
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return '';
        return date.toISOString().split('T')[0];
    } catch {
        return '';
    }
};

export const formatDateForSharePoint = (dateStr) => {
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

export const formatTimeForDisplay = (timeStr) => {
    if (!timeStr) return '';
    try {
        // Handle both "HH:mm" and full datetime formats
        if (timeStr.includes('T')) {
            const date = new Date(timeStr);
            return date.toLocaleTimeString('nl-NL', { 
                hour: '2-digit', 
                minute: '2-digit',
                hour12: false 
            });
        }
        return timeStr;
    } catch {
        return timeStr || '';
    }
};

export const addMinutesToTime = (timeStr, minutes) => {
    if (!timeStr || !minutes) return timeStr;
    
    try {
        // Parse time string (HH:MM format)
        const [hours, mins] = timeStr.split(':').map(Number);
        const totalMinutes = (hours * 60) + mins + minutes;
        
        const newHours = Math.floor(totalMinutes / 60) % 24;
        const newMins = totalMinutes % 60;
        
        return `${newHours.toString().padStart(2, '0')}:${newMins.toString().padStart(2, '0')}`;
    } catch {
        return timeStr;
    }
};

export const parseExcelDate = (excelDate) => {
    if (!excelDate) return null;
    
    try {
        // Handle Excel serial date numbers
        if (typeof excelDate === 'number') {
            // Excel date serial number (days since 1900-01-01, with 1900 incorrectly treated as leap year)
            const excelEpoch = new Date(1900, 0, 1);
            const days = excelDate - 1; // Excel counts from 1, JavaScript from 0
            const date = new Date(excelEpoch.getTime() + (days * 24 * 60 * 60 * 1000));
            return date;
        }
        
        // Handle string dates
        if (typeof excelDate === 'string') {
            const date = new Date(excelDate);
            if (isNaN(date.getTime())) return null;
            return date;
        }
        
        // Handle Date objects
        if (excelDate instanceof Date) {
            return excelDate;
        }
        
        return null;
    } catch {
        return null;
    }
};

export const splitDateTimeToFields = (dateTime) => {
    if (!dateTime) return { date: '', time: '' };
    
    try {
        const date = parseExcelDate(dateTime);
        if (!date) return { date: '', time: '' };
        
        const dateStr = date.toISOString().split('T')[0]; // YYYY-MM-DD
        const timeStr = date.toLocaleTimeString('nl-NL', { 
            hour: '2-digit', 
            minute: '2-digit',
            hour12: false 
        }); // HH:MM
        
        return { date: dateStr, time: timeStr };
    } catch {
        return { date: '', time: '' };
    }
};
