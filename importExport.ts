import type { CellFormat } from '../backend';

interface CellData {
    value: string;
    formula?: string;
    displayValue?: string;
    format?: CellFormat;
}

function getColumnLabel(index: number): string {
    let label = '';
    let num = index;
    while (num >= 0) {
        label = String.fromCharCode(65 + (num % 26)) + label;
        num = Math.floor(num / 26) - 1;
    }
    return label;
}

export function exportToCSV(cells: Map<string, CellData>, filename: string) {
    const rows: string[][] = [];
    let maxRow = 0;
    let maxCol = 0;

    cells.forEach((_, cellId) => {
        const match = cellId.match(/^([A-Z]+)(\d+)$/);
        if (match) {
            const col = match[1].split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
            const row = parseInt(match[2]) - 1;
            maxRow = Math.max(maxRow, row);
            maxCol = Math.max(maxCol, col);
        }
    });

    for (let row = 0; row <= maxRow; row++) {
        const rowData: string[] = [];
        for (let col = 0; col <= maxCol; col++) {
            const cellId = `${getColumnLabel(col)}${row + 1}`;
            const cellData = cells.get(cellId);
            const value = cellData?.displayValue || cellData?.value || '';
            rowData.push(`"${value.replace(/"/g, '""')}"`);
        }
        rows.push(rowData);
    }

    const csv = rows.map(row => row.join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}.csv`;
    a.click();
    URL.revokeObjectURL(url);
}

export function exportToXLSX(cells: Map<string, CellData>, sheets: string[], filename: string) {
    // Since xlsx library is not available, export as CSV instead
    // In a production environment, you would need to add the xlsx package
    exportToCSV(cells, filename);
}

export function exportToJSON(cells: Map<string, CellData>, filename: string) {
    const data: Record<string, any> = {};
    
    cells.forEach((cellData, cellId) => {
        data[cellId] = {
            value: cellData.value,
            formula: cellData.formula,
            displayValue: cellData.displayValue,
            format: cellData.format,
        };
    });

    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}.json`;
    a.click();
    URL.revokeObjectURL(url);
}

export async function importFromFile(file: File): Promise<Map<string, CellData>> {
    const cells = new Map<string, CellData>();

    if (file.name.endsWith('.csv')) {
        const text = await file.text();
        const rows = text.split('\n');
        rows.forEach((row, rowIndex) => {
            if (!row.trim()) return;
            
            // Simple CSV parser (handles quoted values)
            const values: string[] = [];
            let current = '';
            let inQuotes = false;
            
            for (let i = 0; i < row.length; i++) {
                const char = row[i];
                const nextChar = row[i + 1];
                
                if (char === '"' && inQuotes && nextChar === '"') {
                    current += '"';
                    i++; // Skip next quote
                } else if (char === '"') {
                    inQuotes = !inQuotes;
                } else if (char === ',' && !inQuotes) {
                    values.push(current);
                    current = '';
                } else {
                    current += char;
                }
            }
            values.push(current); // Add last value
            
            values.forEach((value, colIndex) => {
                const trimmedValue = value.trim();
                if (trimmedValue) {
                    const cellId = `${getColumnLabel(colIndex)}${rowIndex + 1}`;
                    cells.set(cellId, { value: trimmedValue });
                }
            });
        });
    } else if (file.name.endsWith('.json')) {
        const text = await file.text();
        const data = JSON.parse(text);
        Object.entries(data).forEach(([cellId, cellData]: [string, any]) => {
            cells.set(cellId, {
                value: cellData.value || '',
                formula: cellData.formula,
                displayValue: cellData.displayValue,
                format: cellData.format,
            });
        });
    }

    return cells;
}

export async function importFromXLSX(file: File): Promise<Map<string, CellData>> {
    const cells = new Map<string, CellData>();

    try {
        // Try to dynamically import xlsx library
        let XLSX: any;
        try {
            // Use a function to avoid static analysis by bundler
            const importXLSX = new Function('return import("https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs")');
            XLSX = await importXLSX();
        } catch (importError) {
            throw new Error('Excel import requires the xlsx library. The library could not be loaded. Please use CSV or JSON format instead.');
        }
        
        // Read file as array buffer
        const arrayBuffer = await file.arrayBuffer();
        
        // Parse workbook
        const workbook = XLSX.read(arrayBuffer, { 
            type: 'array',
            cellStyles: true,
            cellHTML: false,
        });
        
        // Get first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Get sheet range
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        // Iterate through cells
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                
                if (!cell) continue;
                
                // Extract cell value
                let value = '';
                if (cell.v !== undefined) {
                    value = String(cell.v);
                }
                
                // Extract formatting
                const format: CellFormat = {};
                
                if (cell.s) {
                    // Bold
                    if (cell.s.font?.bold) {
                        format.bold = true;
                    }
                    
                    // Italic
                    if (cell.s.font?.italic) {
                        format.italic = true;
                    }
                    
                    // Alignment
                    if (cell.s.alignment?.horizontal) {
                        const hAlign = cell.s.alignment.horizontal;
                        if (hAlign === 'left' || hAlign === 'center' || hAlign === 'right') {
                            format.alignment = hAlign;
                        }
                    }
                }
                
                // Convert to our cell ID format (A1, B2, etc.)
                const cellId = `${getColumnLabel(col)}${row + 1}`;
                
                cells.set(cellId, {
                    value,
                    format: Object.keys(format).length > 0 ? format : undefined,
                });
            }
        }
        
        // Handle merged cells
        if (worksheet['!merges']) {
            // Note: Merged cells are stored in the worksheet['!merges'] array
            // Each merge is an object with s (start) and e (end) properties
            // For now, we'll just extract the values without special merge handling
            // as the full merge implementation would require more complex state management
        }
        
        return cells;
    } catch (error: any) {
        console.error('Excel import error:', error);
        
        // Provide user-friendly error message
        if (error.message && error.message.includes('xlsx library')) {
            throw error;
        }
        
        throw new Error('Failed to parse Excel file. Please ensure it is a valid .xlsx file, or use CSV/JSON format instead.');
    }
}

