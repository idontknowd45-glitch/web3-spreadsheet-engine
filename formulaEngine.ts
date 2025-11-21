interface CellData {
    value: string;
    formula?: string;
    displayValue?: string;
}

function parseCellReference(ref: string): { col: number; row: number } | null {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;

    const colStr = match[1];
    const rowStr = match[2];

    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    col -= 1;

    const row = parseInt(rowStr) - 1;

    return { col, row };
}

function parseRange(range: string): string[] {
    const [start, end] = range.split(':');
    if (!start || !end) return [];

    const startRef = parseCellReference(start);
    const endRef = parseCellReference(end);

    if (!startRef || !endRef) return [];

    const cells: string[] = [];
    for (let row = startRef.row; row <= endRef.row; row++) {
        for (let col = startRef.col; col <= endRef.col; col++) {
            let colLabel = '';
            let num = col;
            while (num >= 0) {
                colLabel = String.fromCharCode(65 + (num % 26)) + colLabel;
                num = Math.floor(num / 26) - 1;
            }
            cells.push(`${colLabel}${row + 1}`);
        }
    }

    return cells;
}

function getCellValue(cellId: string, cells: Map<string, CellData>): number {
    const cellData = cells.get(cellId);
    if (!cellData) return 0;

    const value = cellData.displayValue || cellData.value;
    const num = parseFloat(value);
    return isNaN(num) ? 0 : num;
}

function getCellText(cellId: string, cells: Map<string, CellData>): string {
    const cellData = cells.get(cellId);
    if (!cellData) return '';
    return cellData.displayValue || cellData.value;
}

export function evaluateFormula(
    formula: string,
    currentCellId: string,
    cells: Map<string, CellData>
): string {
    if (!formula.startsWith('=')) return formula;

    const expression = formula.substring(1).trim().toUpperCase();

    try {
        // SUM function
        if (expression.startsWith('SUM(') && expression.endsWith(')')) {
            const range = expression.slice(4, -1);
            const cellIds = parseRange(range);
            const sum = cellIds.reduce((acc, cellId) => acc + getCellValue(cellId, cells), 0);
            return sum.toString();
        }

        // AVERAGE function
        if (expression.startsWith('AVERAGE(') && expression.endsWith(')')) {
            const range = expression.slice(8, -1);
            const cellIds = parseRange(range);
            if (cellIds.length === 0) return '0';
            const sum = cellIds.reduce((acc, cellId) => acc + getCellValue(cellId, cells), 0);
            return (sum / cellIds.length).toString();
        }

        // MIN function
        if (expression.startsWith('MIN(') && expression.endsWith(')')) {
            const range = expression.slice(4, -1);
            const cellIds = parseRange(range);
            if (cellIds.length === 0) return '0';
            const values = cellIds.map(cellId => getCellValue(cellId, cells));
            return Math.min(...values).toString();
        }

        // MAX function
        if (expression.startsWith('MAX(') && expression.endsWith(')')) {
            const range = expression.slice(4, -1);
            const cellIds = parseRange(range);
            if (cellIds.length === 0) return '0';
            const values = cellIds.map(cellId => getCellValue(cellId, cells));
            return Math.max(...values).toString();
        }

        // COUNT function
        if (expression.startsWith('COUNT(') && expression.endsWith(')')) {
            const range = expression.slice(6, -1);
            const cellIds = parseRange(range);
            return cellIds.filter(cellId => {
                const val = getCellText(cellId, cells);
                return val !== '' && !isNaN(parseFloat(val));
            }).length.toString();
        }

        // COUNTA function
        if (expression.startsWith('COUNTA(') && expression.endsWith(')')) {
            const range = expression.slice(7, -1);
            const cellIds = parseRange(range);
            return cellIds.filter(cellId => getCellText(cellId, cells) !== '').length.toString();
        }

        // ROUND function
        if (expression.startsWith('ROUND(') && expression.endsWith(')')) {
            const args = expression.slice(6, -1).split(',');
            if (args.length !== 2) return '#VALUE!';
            const value = parseFloat(args[0]);
            const decimals = parseInt(args[1]);
            if (isNaN(value) || isNaN(decimals)) return '#VALUE!';
            return value.toFixed(decimals);
        }

        // CONCAT function
        if (expression.startsWith('CONCAT(') && expression.endsWith(')')) {
            const args = expression.slice(7, -1).split(',');
            return args.map(arg => {
                const cellRef = parseCellReference(arg.trim());
                if (cellRef) {
                    return getCellText(arg.trim(), cells);
                }
                return arg.trim().replace(/^["']|["']$/g, '');
            }).join('');
        }

        // LEN function
        if (expression.startsWith('LEN(') && expression.endsWith(')')) {
            const arg = expression.slice(4, -1).trim();
            const cellRef = parseCellReference(arg);
            const text = cellRef ? getCellText(arg, cells) : arg.replace(/^["']|["']$/g, '');
            return text.length.toString();
        }

        // UPPER function
        if (expression.startsWith('UPPER(') && expression.endsWith(')')) {
            const arg = expression.slice(6, -1).trim();
            const cellRef = parseCellReference(arg);
            const text = cellRef ? getCellText(arg, cells) : arg.replace(/^["']|["']$/g, '');
            return text.toUpperCase();
        }

        // LOWER function
        if (expression.startsWith('LOWER(') && expression.endsWith(')')) {
            const arg = expression.slice(6, -1).trim();
            const cellRef = parseCellReference(arg);
            const text = cellRef ? getCellText(arg, cells) : arg.replace(/^["']|["']$/g, '');
            return text.toLowerCase();
        }

        // TODAY function
        if (expression === 'TODAY()') {
            return new Date().toLocaleDateString();
        }

        // NOW function
        if (expression === 'NOW()') {
            return new Date().toLocaleString();
        }

        // IF function
        const ifMatch = expression.match(/^IF\((.+),(.+),(.+)\)$/);
        if (ifMatch) {
            const condition = ifMatch[1].trim();
            const trueValue = ifMatch[2].trim();
            const falseValue = ifMatch[3].trim();

            const compMatch = condition.match(/([A-Z0-9]+)\s*([><=!]+)\s*(.+)/);
            if (compMatch) {
                const cellRef = compMatch[1];
                const operator = compMatch[2];
                const compareValue = parseFloat(compMatch[3]);

                const cellValue = getCellValue(cellRef, cells);

                let result = false;
                switch (operator) {
                    case '>':
                        result = cellValue > compareValue;
                        break;
                    case '<':
                        result = cellValue < compareValue;
                        break;
                    case '>=':
                        result = cellValue >= compareValue;
                        break;
                    case '<=':
                        result = cellValue <= compareValue;
                        break;
                    case '==':
                    case '=':
                        result = cellValue === compareValue;
                        break;
                    case '!=':
                        result = cellValue !== compareValue;
                        break;
                }

                return result ? trueValue.replace(/^["']|["']$/g, '') : falseValue.replace(/^["']|["']$/g, '');
            }
        }

        // VLOOKUP function
        const vlookupMatch = expression.match(/^VLOOKUP\((.+),([A-Z0-9:]+),(\d+),(.+)\)$/);
        if (vlookupMatch) {
            const lookupValue = vlookupMatch[1].trim().replace(/^["']|["']$/g, '');
            const tableRange = vlookupMatch[2];
            const columnIndex = parseInt(vlookupMatch[3]);

            const cellIds = parseRange(tableRange);
            if (cellIds.length === 0) return '#N/A';

            const startRef = parseCellReference(tableRange.split(':')[0]);
            const endRef = parseCellReference(tableRange.split(':')[1]);
            if (!startRef || !endRef) return '#N/A';

            const numCols = endRef.col - startRef.col + 1;
            const numRows = Math.floor(cellIds.length / numCols);

            for (let row = 0; row < numRows; row++) {
                const firstCellId = cellIds[row * numCols];
                const cellValue = getCellText(firstCellId, cells);

                if (cellValue === lookupValue) {
                    const targetCellId = cellIds[row * numCols + columnIndex - 1];
                    return getCellText(targetCellId, cells);
                }
            }

            return '#N/A';
        }

        // Simple cell reference
        const cellRef = parseCellReference(expression);
        if (cellRef) {
            return getCellValue(expression, cells).toString();
        }

        // Try to evaluate as arithmetic expression
        const sanitized = expression.replace(/[^0-9+\-*/().]/g, '');
        if (sanitized) {
            const result = eval(sanitized);
            return result.toString();
        }

        return '#NAME?';
    } catch (error) {
        return '#ERROR!';
    }
}
