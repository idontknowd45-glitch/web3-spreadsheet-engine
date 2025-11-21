import { useState, useEffect, useCallback, useRef } from 'react';
import { toast } from 'sonner';
import { useGetSpreadsheet, useGetSheet, useSaveCell, useAddSheet, useDeleteSheet, useShareSpreadsheet, useDeleteSpreadsheet, useApplyFormatToSelection, useApplyFontColor, useApplyFillColor, useAddImage, useUpdateImage, useDeleteImage, useSwitchSheet } from '../hooks/useQueries';
import { evaluateFormula } from '../lib/formulaEngine';
import { exportToCSV, exportToXLSX, exportToJSON, importFromFile, importFromXLSX } from '../lib/importExport';
import type { Sheet, SpreadsheetPermission, CellFormat, ImageData } from '../backend';
import ExcelRibbon from './ExcelRibbon';
import ExcelGrid from './ExcelGrid';
import ExcelStatusBar from './ExcelStatusBar';
import ImageLayer from './ImageLayer';

const INITIAL_ROWS = 10000;
const INITIAL_COLS = 50;

function getColumnLabel(index: number): string {
    let label = '';
    let num = index;
    while (num >= 0) {
        label = String.fromCharCode(65 + (num % 26)) + label;
        num = Math.floor(num / 26) - 1;
    }
    return label;
}

function getCellId(row: number, col: number): string {
    return `${getColumnLabel(col)}${row + 1}`;
}

interface CellData {
    value: string;
    formula?: string;
    displayValue?: string;
    format?: CellFormat;
}

interface LocalImageData {
    id: string;
    src: string;
    x: number;
    y: number;
    width: number;
    height: number;
    anchorCell: string;
}

interface HistoryEntry {
    type: 'edit' | 'format' | 'fontColor' | 'fillColor' | 'image-add' | 'image-update' | 'image-delete';
    cellId?: string;
    cellIds?: string[];
    oldValue?: string;
    oldFormula?: string;
    newValue?: string;
    newFormula?: string;
    oldFormats?: Map<string, CellFormat | undefined>;
    newFormat?: CellFormat;
    color?: string;
    oldColors?: Map<string, string | undefined>;
    image?: LocalImageData;
    oldImage?: LocalImageData;
}

// Store per-sheet state
interface SheetState {
    cells: Map<string, CellData>;
    images: LocalImageData[];
    history: HistoryEntry[];
    historyIndex: number;
}

export default function SpreadsheetGrid({ spreadsheetId }: { spreadsheetId: string }) {
    const [activeSheet, setActiveSheet] = useState<string>('Sheet1');
    const [selectedCell, setSelectedCell] = useState<{ row: number; col: number } | null>(null);
    const [selectionStart, setSelectionStart] = useState<{ row: number; col: number } | null>(null);
    const [selectionEnd, setSelectionEnd] = useState<{ row: number; col: number } | null>(null);
    const [formulaBarValue, setFormulaBarValue] = useState('');
    const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
    const [copiedCells, setCopiedCells] = useState<Map<string, CellData> | null>(null);
    const [isEditingCell, setIsEditingCell] = useState(false);
    const [zoom, setZoom] = useState(100);
    const [currentFormat, setCurrentFormat] = useState<CellFormat>({});
    const [gridOffset, setGridOffset] = useState({ top: 0, left: 0 });

    // Store state per sheet
    const [sheetStates, setSheetStates] = useState<Map<string, SheetState>>(new Map());

    const { data: spreadsheet } = useGetSpreadsheet(spreadsheetId);
    const { data: sheet } = useGetSheet(spreadsheetId, activeSheet);
    const saveCell = useSaveCell();
    const addSheet = useAddSheet();
    const deleteSheet = useDeleteSheet();
    const shareSpreadsheet = useShareSpreadsheet();
    const deleteSpreadsheet = useDeleteSpreadsheet();
    const applyFormatMutation = useApplyFormatToSelection();
    const applyFontColorMutation = useApplyFontColor();
    const applyFillColorMutation = useApplyFillColor();
    const addImageMutation = useAddImage();
    const updateImageMutation = useUpdateImage();
    const deleteImageMutation = useDeleteImage();
    const switchSheetMutation = useSwitchSheet();

    const inputRef = useRef<HTMLInputElement>(null);
    const gridContainerRef = useRef<HTMLDivElement>(null);

    // Get current sheet state
    const currentSheetState = sheetStates.get(activeSheet) || {
        cells: new Map(),
        images: [],
        history: [],
        historyIndex: -1,
    };

    const cells = currentSheetState.cells;
    const images = currentSheetState.images;
    const history = currentSheetState.history;
    const historyIndex = currentSheetState.historyIndex;

    // Update current sheet state
    const updateSheetState = useCallback((updates: Partial<SheetState>) => {
        setSheetStates(prev => {
            const newStates = new Map(prev);
            const currentState = newStates.get(activeSheet) || {
                cells: new Map(),
                images: [],
                history: [],
                historyIndex: -1,
            };
            newStates.set(activeSheet, { ...currentState, ...updates });
            return newStates;
        });
    }, [activeSheet]);

    // Initialize active sheet from backend when spreadsheet loads
    useEffect(() => {
        if (spreadsheet && spreadsheet.activeSheet) {
            setActiveSheet(spreadsheet.activeSheet);
        }
    }, [spreadsheet?.id]); // Only run when spreadsheet ID changes

    // Load sheet data when switching sheets or when sheet data changes
    useEffect(() => {
        if (sheet) {
            // Check if we already have state for this sheet
            const existingState = sheetStates.get(activeSheet);
            
            // Only load from backend if we don't have local state yet
            if (!existingState) {
                const newCells = new Map<string, CellData>();
                const cellEntries = extractCellsFromSheet(sheet);

                cellEntries.forEach(([cellId, cell]) => {
                    newCells.set(cellId, {
                        value: cell.value,
                        formula: cell.formula,
                        format: cell.format,
                    });
                });

                // Load images
                const imageEntries = extractImagesFromSheet(sheet);
                const loadedImages: LocalImageData[] = imageEntries.map(([_, img]) => ({
                    id: img.id,
                    src: img.src,
                    x: Number(img.x),
                    y: Number(img.y),
                    width: Number(img.width),
                    height: Number(img.height),
                    anchorCell: img.anchorCell,
                }));

                // Initialize state for this sheet
                setSheetStates(prev => {
                    const newStates = new Map(prev);
                    newStates.set(activeSheet, {
                        cells: newCells,
                        images: loadedImages,
                        history: [],
                        historyIndex: -1,
                    });
                    return newStates;
                });

                setHasUnsavedChanges(false);
            }
        }
    }, [sheet, activeSheet]);

    // Recalculate formulas when cells change
    useEffect(() => {
        const newCells = new Map(cells);
        let hasChanges = false;

        cells.forEach((cellData, cellId) => {
            if (cellData.formula) {
                const displayValue = evaluateFormula(cellData.formula, cellId, cells);
                if (cellData.displayValue !== displayValue) {
                    newCells.set(cellId, { ...cellData, displayValue });
                    hasChanges = true;
                }
            }
        });

        if (hasChanges) {
            updateSheetState({ cells: newCells });
        }
    }, [cells, updateSheetState]);

    // Update current format when selection changes
    useEffect(() => {
        if (selectedCell) {
            const cellId = getCellId(selectedCell.row, selectedCell.col);
            const cellData = cells.get(cellId);
            setCurrentFormat(cellData?.format || {});
        }
    }, [selectedCell, cells]);

    // Get selected cell IDs
    const getSelectedCellIds = useCallback((): string[] => {
        const selectedCellIds: string[] = [];

        if (selectionEnd && selectionStart) {
            const minRow = Math.min(selectionStart.row, selectionEnd.row);
            const maxRow = Math.max(selectionStart.row, selectionEnd.row);
            const minCol = Math.min(selectionStart.col, selectionEnd.col);
            const maxCol = Math.max(selectionStart.col, selectionEnd.col);

            for (let row = minRow; row <= maxRow; row++) {
                for (let col = minCol; col <= maxCol; col++) {
                    selectedCellIds.push(getCellId(row, col));
                }
            }
        } else if (selectedCell) {
            selectedCellIds.push(getCellId(selectedCell.row, selectedCell.col));
        }

        return selectedCellIds;
    }, [selectedCell, selectionStart, selectionEnd]);

    // Centralized formatting function for bold, italic, underline, alignment
    const applyFormatToSelection = useCallback((formatObj: CellFormat) => {
        if (!selectedCell && !selectionEnd) return;

        const selectedCellIds = getSelectedCellIds();
        const oldFormats = new Map<string, CellFormat | undefined>();

        selectedCellIds.forEach(cellId => {
            const cellData = cells.get(cellId);
            oldFormats.set(cellId, cellData?.format);
        });

        // Update local state immediately
        const newCells = new Map(cells);
        selectedCellIds.forEach(cellId => {
            const existingCell = newCells.get(cellId) || { value: '', format: {} };
            const mergedFormat = { ...existingCell.format, ...formatObj };
            newCells.set(cellId, {
                ...existingCell,
                format: mergedFormat,
            });
        });

        // Add to history
        const historyEntry: HistoryEntry = {
            type: 'format',
            cellIds: selectedCellIds,
            oldFormats,
            newFormat: formatObj,
        };

        updateSheetState({
            cells: newCells,
            history: [...history.slice(0, historyIndex + 1), historyEntry],
            historyIndex: historyIndex + 1,
        });

        // Persist to backend
        applyFormatMutation.mutate(
            {
                spreadsheetId,
                sheetName: activeSheet,
                cellIds: selectedCellIds,
                formatObj,
            },
            {
                onSuccess: () => {
                    toast.success('Formatting applied');
                },
                onError: () => {
                    toast.error('Failed to apply formatting');
                },
            }
        );
    }, [selectedCell, selectionStart, selectionEnd, cells, spreadsheetId, activeSheet, applyFormatMutation, history, historyIndex, updateSheetState, getSelectedCellIds]);

    // Font color handler
    const handleApplyFontColor = useCallback((color: string) => {
        if (!selectedCell && !selectionEnd) return;

        const selectedCellIds = getSelectedCellIds();
        const oldColors = new Map<string, string | undefined>();

        selectedCellIds.forEach(cellId => {
            const cellData = cells.get(cellId);
            oldColors.set(cellId, cellData?.format?.fontColor);
        });

        // Update local state immediately
        const newCells = new Map(cells);
        selectedCellIds.forEach(cellId => {
            const existingCell = newCells.get(cellId) || { value: '', format: {} };
            const mergedFormat = { ...existingCell.format, fontColor: color };
            newCells.set(cellId, {
                ...existingCell,
                format: mergedFormat,
            });
        });

        // Add to history
        const historyEntry: HistoryEntry = {
            type: 'fontColor',
            cellIds: selectedCellIds,
            oldColors,
            color,
        };

        updateSheetState({
            cells: newCells,
            history: [...history.slice(0, historyIndex + 1), historyEntry],
            historyIndex: historyIndex + 1,
        });

        // Persist to backend
        applyFontColorMutation.mutate(
            {
                spreadsheetId,
                sheetName: activeSheet,
                cellIds: selectedCellIds,
                color,
            },
            {
                onSuccess: () => {
                    toast.success('Font color applied');
                },
                onError: () => {
                    toast.error('Failed to apply font color');
                },
            }
        );
    }, [selectedCell, selectionEnd, cells, spreadsheetId, activeSheet, applyFontColorMutation, history, historyIndex, updateSheetState, getSelectedCellIds]);

    // Fill color handler
    const handleApplyFillColor = useCallback((color: string) => {
        if (!selectedCell && !selectionEnd) return;

        const selectedCellIds = getSelectedCellIds();
        const oldColors = new Map<string, string | undefined>();

        selectedCellIds.forEach(cellId => {
            const cellData = cells.get(cellId);
            oldColors.set(cellId, cellData?.format?.fillColor);
        });

        // Update local state immediately
        const newCells = new Map(cells);
        selectedCellIds.forEach(cellId => {
            const existingCell = newCells.get(cellId) || { value: '', format: {} };
            const mergedFormat = { ...existingCell.format, fillColor: color };
            newCells.set(cellId, {
                ...existingCell,
                format: mergedFormat,
            });
        });

        // Add to history
        const historyEntry: HistoryEntry = {
            type: 'fillColor',
            cellIds: selectedCellIds,
            oldColors,
            color,
        };

        updateSheetState({
            cells: newCells,
            history: [...history.slice(0, historyIndex + 1), historyEntry],
            historyIndex: historyIndex + 1,
        });

        // Persist to backend
        applyFillColorMutation.mutate(
            {
                spreadsheetId,
                sheetName: activeSheet,
                cellIds: selectedCellIds,
                color,
            },
            {
                onSuccess: () => {
                    toast.success('Fill color applied');
                },
                onError: () => {
                    toast.error('Failed to apply fill color');
                },
            }
        );
    }, [selectedCell, selectionEnd, cells, spreadsheetId, activeSheet, applyFillColorMutation, history, historyIndex, updateSheetState, getSelectedCellIds]);

    // Image handling functions
    const handleInsertImage = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        if (!file.type.match(/^image\/(png|jpeg|jpg)$/)) {
            toast.error('Please select a PNG or JPG image');
            e.target.value = '';
            return;
        }

        try {
            const reader = new FileReader();
            reader.onload = async (event) => {
                const base64 = event.target?.result as string;
                
                // Calculate position based on selected cell
                const cellPosition = getCellPosition(selectedCell || { row: 0, col: 0 });
                
                const newImage: LocalImageData = {
                    id: `img-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
                    src: base64,
                    x: cellPosition.x,
                    y: cellPosition.y,
                    width: 200,
                    height: 150,
                    anchorCell: selectedCell ? getCellId(selectedCell.row, selectedCell.col) : 'A1',
                };

                // Add to history
                const historyEntry: HistoryEntry = {
                    type: 'image-add',
                    image: newImage,
                };

                updateSheetState({
                    images: [...images, newImage],
                    history: [...history.slice(0, historyIndex + 1), historyEntry],
                    historyIndex: historyIndex + 1,
                });

                // Persist to backend
                const backendImage: ImageData = {
                    id: newImage.id,
                    src: newImage.src,
                    x: BigInt(newImage.x),
                    y: BigInt(newImage.y),
                    width: BigInt(newImage.width),
                    height: BigInt(newImage.height),
                    anchorCell: newImage.anchorCell,
                };

                addImageMutation.mutate(
                    {
                        spreadsheetId,
                        sheetName: activeSheet,
                        image: backendImage,
                    },
                    {
                        onSuccess: () => {
                            toast.success('Image inserted');
                        },
                        onError: () => {
                            toast.error('Failed to insert image');
                        },
                    }
                );
            };
            reader.readAsDataURL(file);
        } catch (error) {
            toast.error('Failed to read image file');
        }

        e.target.value = '';
    };

    const handleImageUpdate = (updatedImage: LocalImageData) => {
        const oldImage = images.find(img => img.id === updatedImage.id);
        if (!oldImage) return;

        // Add to history
        const historyEntry: HistoryEntry = {
            type: 'image-update',
            image: updatedImage,
            oldImage,
        };

        updateSheetState({
            images: images.map(img => img.id === updatedImage.id ? updatedImage : img),
            history: [...history.slice(0, historyIndex + 1), historyEntry],
            historyIndex: historyIndex + 1,
        });

        // Persist to backend
        const backendImage: ImageData = {
            id: updatedImage.id,
            src: updatedImage.src,
            x: BigInt(updatedImage.x),
            y: BigInt(updatedImage.y),
            width: BigInt(updatedImage.width),
            height: BigInt(updatedImage.height),
            anchorCell: updatedImage.anchorCell,
        };

        updateImageMutation.mutate(
            {
                spreadsheetId,
                sheetName: activeSheet,
                image: backendImage,
            },
            {
                onError: () => {
                    toast.error('Failed to update image');
                },
            }
        );
    };

    const handleImageDelete = (imageId: string) => {
        const deletedImage = images.find(img => img.id === imageId);
        if (!deletedImage) return;

        // Add to history
        const historyEntry: HistoryEntry = {
            type: 'image-delete',
            image: deletedImage,
        };

        updateSheetState({
            images: images.filter(img => img.id !== imageId),
            history: [...history.slice(0, historyIndex + 1), historyEntry],
            historyIndex: historyIndex + 1,
        });

        // Persist to backend
        deleteImageMutation.mutate(
            {
                spreadsheetId,
                sheetName: activeSheet,
                imageId,
            },
            {
                onSuccess: () => {
                    toast.success('Image deleted');
                },
                onError: () => {
                    toast.error('Failed to delete image');
                },
            }
        );
    };

    const getCellPosition = (cell: { row: number; col: number }): { x: number; y: number } => {
        const rowHeight = 24;
        const colWidth = 100;
        const headerHeight = 24;
        const headerWidth = 48;

        return {
            x: cell.col * colWidth + headerWidth,
            y: cell.row * rowHeight + headerHeight,
        };
    };

    // Comprehensive keyboard shortcuts
    useEffect(() => {
        const handleKeyDown = (e: KeyboardEvent) => {
            const isCtrlOrCmd = e.ctrlKey || e.metaKey;

            // Save: Ctrl+S
            if (isCtrlOrCmd && e.key === 's') {
                e.preventDefault();
                if (hasUnsavedChanges) {
                    handleFormulaBarSubmit();
                }
            }
            // Undo: Ctrl+Z
            else if (isCtrlOrCmd && e.key === 'z' && !e.shiftKey) {
                e.preventDefault();
                handleUndo();
            }
            // Redo: Ctrl+Y or Ctrl+Shift+Z
            else if (isCtrlOrCmd && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
                e.preventDefault();
                handleRedo();
            }
            // Copy: Ctrl+C
            else if (isCtrlOrCmd && e.key === 'c' && !isEditingCell) {
                e.preventDefault();
                handleCopy();
            }
            // Cut: Ctrl+X
            else if (isCtrlOrCmd && e.key === 'x' && !isEditingCell) {
                e.preventDefault();
                handleCut();
            }
            // Paste: Ctrl+V
            else if (isCtrlOrCmd && e.key === 'v' && !isEditingCell) {
                e.preventDefault();
                handlePaste();
            }
            // Delete: Delete key
            else if (e.key === 'Delete' && !isEditingCell && selectedCell) {
                e.preventDefault();
                handleDelete();
            }
            // Select All: Ctrl+A
            else if (isCtrlOrCmd && e.key === 'a' && !isEditingCell) {
                e.preventDefault();
                handleSelectAll();
            }
            // Find: Ctrl+F
            else if (isCtrlOrCmd && e.key === 'f') {
                e.preventDefault();
                toast.info('Find functionality coming soon');
            }
            // Replace: Ctrl+H
            else if (isCtrlOrCmd && e.key === 'h') {
                e.preventDefault();
                toast.info('Replace functionality coming soon');
            }
            // Bold: Ctrl+B
            else if (isCtrlOrCmd && e.key === 'b' && !isEditingCell) {
                e.preventDefault();
                applyFormatToSelection({ bold: !currentFormat.bold });
            }
            // Italic: Ctrl+I
            else if (isCtrlOrCmd && e.key === 'i' && !isEditingCell) {
                e.preventDefault();
                applyFormatToSelection({ italic: !currentFormat.italic });
            }
            // Underline: Ctrl+U
            else if (isCtrlOrCmd && e.key === 'u' && !isEditingCell) {
                e.preventDefault();
                applyFormatToSelection({ underline: !currentFormat.underline });
            }
            // Format Cells: Ctrl+1
            else if (isCtrlOrCmd && e.key === '1' && !isEditingCell) {
                e.preventDefault();
                toast.info('Format cells dialog coming soon');
            }
            // Navigation and editing when not editing
            else if (!isEditingCell && selectedCell) {
                // Arrow keys
                if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (e.shiftKey) {
                        handleShiftArrow('up');
                    } else if (isCtrlOrCmd) {
                        handleCtrlArrow('up');
                    } else {
                        moveSelection('up');
                    }
                } else if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    if (e.shiftKey) {
                        handleShiftArrow('down');
                    } else if (isCtrlOrCmd) {
                        handleCtrlArrow('down');
                    } else {
                        moveSelection('down');
                    }
                } else if (e.key === 'ArrowLeft') {
                    e.preventDefault();
                    if (e.shiftKey) {
                        handleShiftArrow('left');
                    } else if (isCtrlOrCmd) {
                        handleCtrlArrow('left');
                    } else {
                        moveSelection('left');
                    }
                } else if (e.key === 'ArrowRight') {
                    e.preventDefault();
                    if (e.shiftKey) {
                        handleShiftArrow('right');
                    } else if (isCtrlOrCmd) {
                        handleCtrlArrow('right');
                    } else {
                        moveSelection('right');
                    }
                }
                // Home: Ctrl+Home (go to A1)
                else if (isCtrlOrCmd && e.key === 'Home') {
                    e.preventDefault();
                    setSelectedCell({ row: 0, col: 0 });
                    setSelectionStart({ row: 0, col: 0 });
                    setSelectionEnd(null);
                }
                // Enter: move down or start editing
                else if (e.key === 'Enter') {
                    e.preventDefault();
                    if (e.shiftKey) {
                        moveSelection('up');
                    } else {
                        setIsEditingCell(true);
                        inputRef.current?.focus();
                    }
                }
                // Tab: move right
                else if (e.key === 'Tab') {
                    e.preventDefault();
                    if (e.shiftKey) {
                        moveSelection('left');
                    } else {
                        moveSelection('right');
                    }
                }
                // F2: edit cell
                else if (e.key === 'F2') {
                    e.preventDefault();
                    setIsEditingCell(true);
                    inputRef.current?.focus();
                }
                // Escape: clear selection
                else if (e.key === 'Escape') {
                    e.preventDefault();
                    setSelectionEnd(null);
                }
                // Space: Ctrl+Space (select column), Shift+Space (select row)
                else if (e.key === ' ') {
                    e.preventDefault();
                    if (isCtrlOrCmd) {
                        handleSelectColumn();
                    } else if (e.shiftKey) {
                        handleSelectRow();
                    }
                }
                // PageUp/PageDown: Ctrl+PageUp/PageDown (switch sheets)
                else if (isCtrlOrCmd && e.key === 'PageUp') {
                    e.preventDefault();
                    handlePreviousSheet();
                } else if (isCtrlOrCmd && e.key === 'PageDown') {
                    e.preventDefault();
                    handleNextSheet();
                }
            }
            // Escape while editing: exit edit mode
            else if (isEditingCell && e.key === 'Escape') {
                e.preventDefault();
                setIsEditingCell(false);
                inputRef.current?.blur();
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [selectedCell, hasUnsavedChanges, isEditingCell, historyIndex, history, selectionStart, selectionEnd, activeSheet, currentFormat, applyFormatToSelection]);

    // Update formula bar when selected cell changes
    useEffect(() => {
        if (selectedCell) {
            const cellId = getCellId(selectedCell.row, selectedCell.col);
            const cellData = cells.get(cellId);
            setFormulaBarValue(cellData?.formula || cellData?.value || '');
        }
    }, [selectedCell, cells]);

    const extractCellsFromSheet = (sheet: Sheet): [string, any][] => {
        const entries: [string, any][] = [];
        const traverse = (node: any) => {
            if (!node) return;
            if (node.__kind__ === 'leaf') return;

            const [left, key, value, right] = node[node.__kind__];
            traverse(left);
            entries.push([key, value]);
            traverse(right);
        };

        traverse(sheet.cells.root);
        return entries;
    };

    const extractImagesFromSheet = (sheet: Sheet): [string, any][] => {
        const entries: [string, any][] = [];
        const traverse = (node: any) => {
            if (!node) return;
            if (node.__kind__ === 'leaf') return;

            const [left, key, value, right] = node[node.__kind__];
            traverse(left);
            entries.push([key, value]);
            traverse(right);
        };

        traverse(sheet.images.root);
        return entries;
    };

    const moveSelection = (direction: 'up' | 'down' | 'left' | 'right') => {
        if (!selectedCell) return;

        let newRow = selectedCell.row;
        let newCol = selectedCell.col;

        switch (direction) {
            case 'up':
                newRow = Math.max(0, selectedCell.row - 1);
                break;
            case 'down':
                newRow = Math.min(INITIAL_ROWS - 1, selectedCell.row + 1);
                break;
            case 'left':
                newCol = Math.max(0, selectedCell.col - 1);
                break;
            case 'right':
                newCol = Math.min(INITIAL_COLS - 1, selectedCell.col + 1);
                break;
        }

        setSelectedCell({ row: newRow, col: newCol });
        setSelectionStart({ row: newRow, col: newCol });
        setSelectionEnd(null);
    };

    const handleShiftArrow = (direction: 'up' | 'down' | 'left' | 'right') => {
        if (!selectedCell || !selectionStart) return;

        let newRow = selectionEnd?.row ?? selectedCell.row;
        let newCol = selectionEnd?.col ?? selectedCell.col;

        switch (direction) {
            case 'up':
                newRow = Math.max(0, newRow - 1);
                break;
            case 'down':
                newRow = Math.min(INITIAL_ROWS - 1, newRow + 1);
                break;
            case 'left':
                newCol = Math.max(0, newCol - 1);
                break;
            case 'right':
                newCol = Math.min(INITIAL_COLS - 1, newCol + 1);
                break;
        }

        setSelectionEnd({ row: newRow, col: newCol });
    };

    const handleCtrlArrow = (direction: 'up' | 'down' | 'left' | 'right') => {
        if (!selectedCell) return;

        let newRow = selectedCell.row;
        let newCol = selectedCell.col;

        // Jump to edge of data region or edge of sheet
        switch (direction) {
            case 'up':
                newRow = 0;
                break;
            case 'down':
                newRow = INITIAL_ROWS - 1;
                break;
            case 'left':
                newCol = 0;
                break;
            case 'right':
                newCol = INITIAL_COLS - 1;
                break;
        }

        setSelectedCell({ row: newRow, col: newCol });
        setSelectionStart({ row: newRow, col: newCol });
        setSelectionEnd(null);
    };

    const handleSelectAll = () => {
        setSelectionStart({ row: 0, col: 0 });
        setSelectionEnd({ row: INITIAL_ROWS - 1, col: INITIAL_COLS - 1 });
        setSelectedCell({ row: 0, col: 0 });
    };

    const handleSelectColumn = () => {
        if (!selectedCell) return;
        setSelectionStart({ row: 0, col: selectedCell.col });
        setSelectionEnd({ row: INITIAL_ROWS - 1, col: selectedCell.col });
    };

    const handleSelectRow = () => {
        if (!selectedCell) return;
        setSelectionStart({ row: selectedCell.row, col: 0 });
        setSelectionEnd({ row: selectedCell.row, col: INITIAL_COLS - 1 });
    };

    const handlePreviousSheet = () => {
        const currentIndex = sheets.indexOf(activeSheet);
        if (currentIndex > 0) {
            handleSheetChange(sheets[currentIndex - 1]);
        }
    };

    const handleNextSheet = () => {
        const currentIndex = sheets.indexOf(activeSheet);
        if (currentIndex < sheets.length - 1) {
            handleSheetChange(sheets[currentIndex + 1]);
        }
    };

    const handleCellClick = (row: number, col: number, e?: React.MouseEvent) => {
        if (e?.shiftKey && selectedCell) {
            setSelectionEnd({ row, col });
        } else {
            setSelectedCell({ row, col });
            setSelectionStart({ row, col });
            setSelectionEnd(null);
        }
        setIsEditingCell(false);
    };

    const handleSelectionChange = (start: { row: number; col: number }, end: { row: number; col: number } | null) => {
        setSelectionStart(start);
        setSelectionEnd(end);
        setSelectedCell(start);
    };

    const handleCellEdit = (row: number, col: number) => {
        setSelectedCell({ row, col });
        setIsEditingCell(true);
        setTimeout(() => inputRef.current?.focus(), 0);
    };

    const handleFormulaBarChange = (value: string) => {
        setFormulaBarValue(value);
        setHasUnsavedChanges(true);
    };

    const handleFormulaBarSubmit = useCallback(() => {
        if (!selectedCell) return;

        const cellId = getCellId(selectedCell.row, selectedCell.col);
        const oldCellData = cells.get(cellId);
        const isFormula = formulaBarValue.startsWith('=');
        const formula = isFormula ? formulaBarValue : null;
        const value = isFormula ? '' : formulaBarValue;

        // Add to history
        const historyEntry: HistoryEntry = {
            type: 'edit',
            cellId,
            oldValue: oldCellData?.value || '',
            oldFormula: oldCellData?.formula,
            newValue: value,
            newFormula: formula || undefined,
        };

        const newCells = new Map(cells);
        newCells.set(cellId, {
            value,
            formula: formula || undefined,
            format: oldCellData?.format,
        });

        updateSheetState({
            cells: newCells,
            history: [...history.slice(0, historyIndex + 1), historyEntry],
            historyIndex: historyIndex + 1,
        });

        saveCell.mutate(
            {
                spreadsheetId,
                sheetName: activeSheet,
                cellId,
                value,
                formula,
            },
            {
                onSuccess: () => {
                    setHasUnsavedChanges(false);
                    toast.success('Cell saved');
                },
                onError: () => {
                    toast.error('Failed to save cell');
                },
            }
        );
        setIsEditingCell(false);
    }, [selectedCell, formulaBarValue, cells, spreadsheetId, activeSheet, saveCell, history, historyIndex, updateSheetState]);

    const handleUndo = () => {
        if (historyIndex >= 0) {
            const entry = history[historyIndex];
            const newCells = new Map(cells);

            if (entry.type === 'edit' && entry.cellId) {
                newCells.set(entry.cellId, {
                    value: entry.oldValue || '',
                    formula: entry.oldFormula,
                    format: newCells.get(entry.cellId)?.format,
                });
                updateSheetState({ cells: newCells, historyIndex: historyIndex - 1 });
            } else if (entry.type === 'format' && entry.cellIds && entry.oldFormats) {
                entry.cellIds.forEach(cellId => {
                    const existingCell = newCells.get(cellId) || { value: '' };
                    newCells.set(cellId, {
                        ...existingCell,
                        format: entry.oldFormats?.get(cellId),
                    });
                });
                updateSheetState({ cells: newCells, historyIndex: historyIndex - 1 });
            } else if (entry.type === 'fontColor' && entry.cellIds && entry.oldColors) {
                entry.cellIds.forEach(cellId => {
                    const existingCell = newCells.get(cellId) || { value: '' };
                    const mergedFormat = { ...existingCell.format, fontColor: entry.oldColors?.get(cellId) };
                    newCells.set(cellId, {
                        ...existingCell,
                        format: mergedFormat,
                    });
                });
                updateSheetState({ cells: newCells, historyIndex: historyIndex - 1 });
            } else if (entry.type === 'fillColor' && entry.cellIds && entry.oldColors) {
                entry.cellIds.forEach(cellId => {
                    const existingCell = newCells.get(cellId) || { value: '' };
                    const mergedFormat = { ...existingCell.format, fillColor: entry.oldColors?.get(cellId) };
                    newCells.set(cellId, {
                        ...existingCell,
                        format: mergedFormat,
                    });
                });
                updateSheetState({ cells: newCells, historyIndex: historyIndex - 1 });
            } else if (entry.type === 'image-add' && entry.image) {
                updateSheetState({
                    images: images.filter(img => img.id !== entry.image!.id),
                    historyIndex: historyIndex - 1,
                });
            } else if (entry.type === 'image-update' && entry.oldImage) {
                updateSheetState({
                    images: images.map(img => img.id === entry.oldImage!.id ? entry.oldImage! : img),
                    historyIndex: historyIndex - 1,
                });
            } else if (entry.type === 'image-delete' && entry.image) {
                updateSheetState({
                    images: [...images, entry.image],
                    historyIndex: historyIndex - 1,
                });
            }

            toast.success('Undo');
        }
    };

    const handleRedo = () => {
        if (historyIndex < history.length - 1) {
            const entry = history[historyIndex + 1];
            const newCells = new Map(cells);

            if (entry.type === 'edit' && entry.cellId) {
                newCells.set(entry.cellId, {
                    value: entry.newValue || '',
                    formula: entry.newFormula,
                    format: newCells.get(entry.cellId)?.format,
                });
                updateSheetState({ cells: newCells, historyIndex: historyIndex + 1 });
            } else if (entry.type === 'format' && entry.cellIds && entry.newFormat) {
                entry.cellIds.forEach(cellId => {
                    const existingCell = newCells.get(cellId) || { value: '' };
                    const mergedFormat = { ...existingCell.format, ...entry.newFormat };
                    newCells.set(cellId, {
                        ...existingCell,
                        format: mergedFormat,
                    });
                });
                updateSheetState({ cells: newCells, historyIndex: historyIndex + 1 });
            } else if (entry.type === 'fontColor' && entry.cellIds && entry.color) {
                entry.cellIds.forEach(cellId => {
                    const existingCell = newCells.get(cellId) || { value: '' };
                    const mergedFormat = { ...existingCell.format, fontColor: entry.color };
                    newCells.set(cellId, {
                        ...existingCell,
                        format: mergedFormat,
                    });
                });
                updateSheetState({ cells: newCells, historyIndex: historyIndex + 1 });
            } else if (entry.type === 'fillColor' && entry.cellIds && entry.color) {
                entry.cellIds.forEach(cellId => {
                    const existingCell = newCells.get(cellId) || { value: '' };
                    const mergedFormat = { ...existingCell.format, fillColor: entry.color };
                    newCells.set(cellId, {
                        ...existingCell,
                        format: mergedFormat,
                    });
                });
                updateSheetState({ cells: newCells, historyIndex: historyIndex + 1 });
            } else if (entry.type === 'image-add' && entry.image) {
                updateSheetState({
                    images: [...images, entry.image],
                    historyIndex: historyIndex + 1,
                });
            } else if (entry.type === 'image-update' && entry.image) {
                updateSheetState({
                    images: images.map(img => img.id === entry.image!.id ? entry.image! : img),
                    historyIndex: historyIndex + 1,
                });
            } else if (entry.type === 'image-delete' && entry.image) {
                updateSheetState({
                    images: images.filter(img => img.id !== entry.image!.id),
                    historyIndex: historyIndex + 1,
                });
            }

            toast.success('Redo');
        }
    };

    const handleCopy = () => {
        if (!selectedCell) return;
        const cellId = getCellId(selectedCell.row, selectedCell.col);
        const cellData = cells.get(cellId);
        if (cellData) {
            const copiedData = new Map<string, CellData>();
            copiedData.set(cellId, cellData);
            setCopiedCells(copiedData);
            toast.success('Cell copied');
        }
    };

    const handleCut = () => {
        if (!selectedCell) return;
        const cellId = getCellId(selectedCell.row, selectedCell.col);
        const cellData = cells.get(cellId);
        if (cellData) {
            const copiedData = new Map<string, CellData>();
            copiedData.set(cellId, cellData);
            setCopiedCells(copiedData);
            
            const newCells = new Map(cells);
            newCells.delete(cellId);
            updateSheetState({ cells: newCells });
            toast.success('Cell cut');
        }
    };

    const handlePaste = () => {
        if (!selectedCell || !copiedCells) return;
        const targetCellId = getCellId(selectedCell.row, selectedCell.col);
        const copiedData = Array.from(copiedCells.values())[0];
        
        const newCells = new Map(cells);
        newCells.set(targetCellId, copiedData);
        updateSheetState({ cells: newCells });
        toast.success('Cell pasted');
    };

    const handleDelete = () => {
        if (!selectedCell) return;
        const cellId = getCellId(selectedCell.row, selectedCell.col);
        
        const newCells = new Map(cells);
        newCells.delete(cellId);
        updateSheetState({ cells: newCells });
        setFormulaBarValue('');
        toast.success('Cell cleared');
    };

    const handleDragFill = (startRow: number, startCol: number, endRow: number, endCol: number) => {
        // Simple drag fill implementation
        const sourceCellId = getCellId(startRow, startCol);
        const sourceCell = cells.get(sourceCellId);
        
        if (!sourceCell) return;

        const newCells = new Map(cells);
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);

        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                if (row === startRow && col === startCol) continue;
                const targetCellId = getCellId(row, col);
                newCells.set(targetCellId, { ...sourceCell });
            }
        }

        updateSheetState({ cells: newCells });
        toast.success('Cells filled');
    };

    const handleAddSheet = async () => {
        try {
            const newSheetName = await addSheet.mutateAsync({
                spreadsheetId,
            });
            handleSheetChange(newSheetName);
            toast.success(`Sheet "${newSheetName}" created`);
        } catch (error) {
            toast.error('Failed to add sheet');
        }
    };

    const handleDeleteSheet = async (sheetName: string) => {
        if (sheets.length <= 1) {
            toast.error('Cannot delete the last sheet');
            return;
        }

        try {
            await deleteSheet.mutateAsync({
                spreadsheetId,
                sheetName,
            });
            if (activeSheet === sheetName) {
                const newActiveSheet = sheets[0] === sheetName ? sheets[1] : sheets[0];
                handleSheetChange(newActiveSheet);
            }
            toast.success('Sheet deleted');
        } catch (error) {
            toast.error('Failed to delete sheet');
        }
    };

    const handleSheetChange = (sheetName: string) => {
        if (sheetName === activeSheet) return;

        // Switch to the new sheet
        setActiveSheet(sheetName);

        // Persist the active sheet to backend
        switchSheetMutation.mutate(
            {
                spreadsheetId,
                sheetName,
            },
            {
                onError: () => {
                    toast.error('Failed to switch sheet');
                },
            }
        );
    };

    const handleExportCSV = () => {
        exportToCSV(cells, spreadsheet?.name || 'spreadsheet');
        toast.success('Exported to CSV');
    };

    const handleExportXLSX = () => {
        exportToXLSX(cells, sheets, spreadsheet?.name || 'spreadsheet');
        toast.success('Exported (CSV format - XLSX requires additional library)');
    };

    const handleExportJSON = () => {
        exportToJSON(cells, spreadsheet?.name || 'spreadsheet');
        toast.success('Exported to JSON');
    };

    const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        try {
            const importedCells = await importFromFile(file);
            updateSheetState({ cells: importedCells });
            toast.success('File imported successfully');
        } catch (error: any) {
            toast.error(error.message || 'Failed to import file');
        }
        
        e.target.value = '';
    };

    const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        if (!file.name.endsWith('.xlsx')) {
            toast.error('Please select a valid .xlsx file');
            e.target.value = '';
            return;
        }

        try {
            const importedCells = await importFromXLSX(file);
            
            // Replace current sheet data
            updateSheetState({
                cells: importedCells,
                history: [],
                historyIndex: -1,
            });
            
            toast.success('Excel file imported successfully');
        } catch (error: any) {
            toast.error(error.message || 'Invalid or corrupted Excel file');
        }
        
        e.target.value = '';
    };

    const handleShare = async (email: string, permission: SpreadsheetPermission) => {
        try {
            await shareSpreadsheet.mutateAsync({
                spreadsheetId,
                user: email,
                permission,
            });
            toast.success('Spreadsheet shared successfully');
        } catch (error) {
            toast.error('Failed to share spreadsheet');
        }
    };

    const handleDeleteSpreadsheet = async () => {
        try {
            await deleteSpreadsheet.mutateAsync(spreadsheetId);
            toast.success('Spreadsheet deleted');
        } catch (error) {
            toast.error('Failed to delete spreadsheet');
        }
    };

    const sheets = spreadsheet
        ? (() => {
              const sheetNames: string[] = [];
              const traverse = (node: any) => {
                  if (!node) return;
                  if (node.__kind__ === 'leaf') return;

                  const [left, key, , right] = node[node.__kind__];
                  traverse(left);
                  sheetNames.push(key);
                  traverse(right);
              };

              traverse(spreadsheet.sheets.root);
              return sheetNames;
          })()
        : [];

    const getCellDisplay = (row: number, col: number): string => {
        const cellId = getCellId(row, col);
        const cellData = cells.get(cellId);
        if (!cellData) return '';
        if (cellData.formula) {
            return cellData.displayValue || '';
        }
        return cellData.value;
    };

    const getCellFormat = (row: number, col: number): CellFormat | undefined => {
        const cellId = getCellId(row, col);
        const cellData = cells.get(cellId);
        return cellData?.format;
    };

    const isCellSelected = (row: number, col: number): boolean => {
        if (!selectedCell) return false;
        if (!selectionEnd) {
            return selectedCell.row === row && selectedCell.col === col;
        }
        const minRow = Math.min(selectionStart?.row || 0, selectionEnd.row);
        const maxRow = Math.max(selectionStart?.row || 0, selectionEnd.row);
        const minCol = Math.min(selectionStart?.col || 0, selectionEnd.col);
        const maxCol = Math.max(selectionStart?.col || 0, selectionEnd.col);
        return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
    };

    return (
        <div className="flex h-full flex-col bg-background" role="application" aria-label="Spreadsheet application">
            {/* Excel Ribbon */}
            <ExcelRibbon
                selectedCell={selectedCell}
                formulaBarValue={formulaBarValue}
                onFormulaBarChange={handleFormulaBarChange}
                onFormulaBarSubmit={handleFormulaBarSubmit}
                onUndo={handleUndo}
                onRedo={handleRedo}
                onSave={handleFormulaBarSubmit}
                onExportCSV={handleExportCSV}
                onExportXLSX={handleExportXLSX}
                onExportJSON={handleExportJSON}
                onImport={handleImport}
                onImportExcel={handleImportExcel}
                onShare={handleShare}
                onDeleteSpreadsheet={handleDeleteSpreadsheet}
                canUndo={historyIndex >= 0}
                canRedo={historyIndex < history.length - 1}
                hasUnsavedChanges={hasUnsavedChanges}
                isSaving={saveCell.isPending}
                inputRef={inputRef}
                setIsEditingCell={setIsEditingCell}
                spreadsheetId={spreadsheetId}
                spreadsheetName={spreadsheet?.name || 'spreadsheet'}
                currentFormat={currentFormat}
                onApplyFormat={applyFormatToSelection}
                onApplyFontColor={handleApplyFontColor}
                onApplyFillColor={handleApplyFillColor}
                onInsertImage={handleInsertImage}
            />

            {/* Grid with Image Layer */}
            <div ref={gridContainerRef} className="flex-1 relative overflow-hidden">
                <ExcelGrid
                    rows={INITIAL_ROWS}
                    cols={INITIAL_COLS}
                    selectedCell={selectedCell}
                    onCellClick={handleCellClick}
                    getCellDisplay={getCellDisplay}
                    getCellFormat={getCellFormat}
                    isCellSelected={isCellSelected}
                    onCopy={handleCopy}
                    onCut={handleCut}
                    onPaste={handlePaste}
                    zoom={zoom}
                    onCellEdit={handleCellEdit}
                    onSelectionChange={handleSelectionChange}
                    onDragFill={handleDragFill}
                />
                <ImageLayer
                    images={images}
                    onImageUpdate={handleImageUpdate}
                    onImageDelete={handleImageDelete}
                    gridOffset={gridOffset}
                    zoom={zoom}
                />
            </div>

            {/* Status Bar with Sheet Tabs */}
            <ExcelStatusBar
                sheets={sheets}
                activeSheet={activeSheet}
                onSheetChange={handleSheetChange}
                onAddSheet={handleAddSheet}
                onDeleteSheet={handleDeleteSheet}
                zoom={zoom}
                onZoomChange={setZoom}
                cellCount={cells.size}
                isAddingSheet={addSheet.isPending}
            />
        </div>
    );
}
