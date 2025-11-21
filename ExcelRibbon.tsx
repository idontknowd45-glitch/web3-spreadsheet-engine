import { useState, useRef, useEffect } from 'react';
import { Save, Undo, Redo, Printer, Share2, Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, AlignJustify, Download, Upload, Trash2, Search, BarChart3, FileText, Settings, HelpCircle, ChevronDown, Merge, Palette, Type, Hash, Grid3x3, Filter, ArrowUpDown, Keyboard, Image as ImageIcon, FileSpreadsheet, Code } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Tabs, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Separator } from '@/components/ui/separator';
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger, DropdownMenuSeparator } from '@/components/ui/dropdown-menu';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog';
import { Label } from '@/components/ui/label';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip';
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover';
import { KeyTipBadge } from './KeyTipBadge';
import { KeyTipOverlay } from './KeyTipOverlay';
import { KeyboardShortcutsDialog } from './KeyboardShortcutsDialog';
import { KeyTipSettingsDialog } from './KeyTipSettingsDialog';
import { useKeyTips } from '../hooks/useKeyTips';
import { exportUnifiedSourceCode } from '../lib/unifiedExport';
import type { SpreadsheetPermission, CellFormat } from '../backend';

interface ExcelRibbonProps {
    selectedCell: { row: number; col: number } | null;
    formulaBarValue: string;
    onFormulaBarChange: (value: string) => void;
    onFormulaBarSubmit: () => void;
    onUndo: () => void;
    onRedo: () => void;
    onSave: () => void;
    onExportCSV: () => void;
    onExportXLSX: () => void;
    onExportJSON: () => void;
    onImport: (e: React.ChangeEvent<HTMLInputElement>) => void;
    onImportExcel: (e: React.ChangeEvent<HTMLInputElement>) => void;
    onShare: (email: string, permission: SpreadsheetPermission) => void;
    onDeleteSpreadsheet: () => void;
    canUndo: boolean;
    canRedo: boolean;
    hasUnsavedChanges: boolean;
    isSaving: boolean;
    inputRef: React.RefObject<HTMLInputElement | null>;
    setIsEditingCell: (value: boolean) => void;
    spreadsheetId: string;
    spreadsheetName: string;
    currentFormat: CellFormat;
    onApplyFormat: (formatObj: CellFormat) => void;
    onApplyFontColor: (color: string) => void;
    onApplyFillColor: (color: string) => void;
    onInsertImage: (e: React.ChangeEvent<HTMLInputElement>) => void;
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

const COLOR_PALETTE = [
    { name: 'White', value: '#FFFFFF' },
    { name: 'Black', value: '#000000' },
    { name: 'Red', value: '#FF0000' },
    { name: 'Dark Blue', value: '#00008B' },
    { name: 'Green', value: '#008000' },
];

export default function ExcelRibbon({
    selectedCell,
    formulaBarValue,
    onFormulaBarChange,
    onFormulaBarSubmit,
    onUndo,
    onRedo,
    onSave,
    onExportCSV,
    onExportXLSX,
    onExportJSON,
    onImport,
    onImportExcel,
    onShare,
    onDeleteSpreadsheet,
    canUndo,
    canRedo,
    hasUnsavedChanges,
    isSaving,
    inputRef,
    setIsEditingCell,
    spreadsheetId,
    spreadsheetName,
    currentFormat,
    onApplyFormat,
    onApplyFontColor,
    onApplyFillColor,
    onInsertImage,
}: ExcelRibbonProps) {
    const [activeTab, setActiveTab] = useState('home');
    const [showShareDialog, setShowShareDialog] = useState(false);
    const [shareEmail, setShareEmail] = useState('');
    const [sharePermission, setSharePermission] = useState<SpreadsheetPermission>('editor' as any);
    const [showShortcutsDialog, setShowShortcutsDialog] = useState(false);
    const [showSettingsDialog, setShowSettingsDialog] = useState(false);
    const [fontColorOpen, setFontColorOpen] = useState(false);
    const [fillColorOpen, setFillColorOpen] = useState(false);

    const keyTips = useKeyTips();

    // Refs for KeyTip badges
    const fileTabRef = useRef<HTMLButtonElement>(null);
    const homeTabRef = useRef<HTMLButtonElement>(null);
    const insertTabRef = useRef<HTMLButtonElement>(null);
    const formulasTabRef = useRef<HTMLButtonElement>(null);
    const dataTabRef = useRef<HTMLButtonElement>(null);
    const viewTabRef = useRef<HTMLButtonElement>(null);
    const helpTabRef = useRef<HTMLButtonElement>(null);

    const boldBtnRef = useRef<HTMLButtonElement>(null);
    const italicBtnRef = useRef<HTMLButtonElement>(null);
    const underlineBtnRef = useRef<HTMLButtonElement>(null);
    const fontSelectRef = useRef<HTMLButtonElement>(null);
    const sizeSelectRef = useRef<HTMLButtonElement>(null);
    const alignLeftBtnRef = useRef<HTMLButtonElement>(null);
    const alignCenterBtnRef = useRef<HTMLButtonElement>(null);
    const alignRightBtnRef = useRef<HTMLButtonElement>(null);
    const mergeBtnRef = useRef<HTMLButtonElement>(null);
    const undoBtnRef = useRef<HTMLButtonElement>(null);
    const redoBtnRef = useRef<HTMLButtonElement>(null);
    const fontColorBtnRef = useRef<HTMLButtonElement>(null);
    const fillColorBtnRef = useRef<HTMLButtonElement>(null);

    const chartBtnRef = useRef<HTMLButtonElement>(null);
    const imageBtnRef = useRef<HTMLButtonElement>(null);
    const filterBtnRef = useRef<HTMLButtonElement>(null);
    const sortBtnRef = useRef<HTMLButtonElement>(null);

    const imageInputRef = useRef<HTMLInputElement>(null);
    const excelInputRef = useRef<HTMLInputElement>(null);

    // Register KeyTips
    useEffect(() => {
        keyTips.clearKeyTips();

        // Tab KeyTips
        keyTips.registerKeyTip('tab-file', { key: 'F', label: 'File', handler: () => setActiveTab('file') });
        keyTips.registerKeyTip('tab-home', { key: 'H', label: 'Home', handler: () => setActiveTab('home') });
        keyTips.registerKeyTip('tab-insert', { key: 'N', label: 'Insert', handler: () => setActiveTab('insert') });
        keyTips.registerKeyTip('tab-formulas', { key: 'M', label: 'Formulas', handler: () => setActiveTab('formulas') });
        keyTips.registerKeyTip('tab-data', { key: 'A', label: 'Data', handler: () => setActiveTab('data') });
        keyTips.registerKeyTip('tab-view', { key: 'W', label: 'View', handler: () => setActiveTab('view') });
        keyTips.registerKeyTip('tab-help', { key: 'E', label: 'Help', handler: () => setActiveTab('help') });

        // Home tab commands
        keyTips.registerKeyTip('home-bold', { key: 'B', label: 'Bold', tab: 'H', handler: () => onApplyFormat({ bold: !currentFormat.bold }) });
        keyTips.registerKeyTip('home-italic', { key: 'I', label: 'Italic', tab: 'H', handler: () => onApplyFormat({ italic: !currentFormat.italic }) });
        keyTips.registerKeyTip('home-underline', { key: 'U', label: 'Underline', tab: 'H', handler: () => onApplyFormat({ underline: !currentFormat.underline }) });
        keyTips.registerKeyTip('home-font', { key: 'F', label: 'Font', tab: 'H', handler: () => fontSelectRef.current?.click() });
        keyTips.registerKeyTip('home-size', { key: 'S', label: 'Size', tab: 'H', handler: () => sizeSelectRef.current?.click() });
        keyTips.registerKeyTip('home-align-left', { key: 'L', label: 'Left', tab: 'H', handler: () => onApplyFormat({ alignment: 'left' }) });
        keyTips.registerKeyTip('home-align-center', { key: 'C', label: 'Center', tab: 'H', handler: () => onApplyFormat({ alignment: 'center' }) });
        keyTips.registerKeyTip('home-align-right', { key: 'R', label: 'Right', tab: 'H', handler: () => onApplyFormat({ alignment: 'right' }) });
        keyTips.registerKeyTip('home-merge', { key: 'M', label: 'Merge', tab: 'H', handler: () => {} });
        keyTips.registerKeyTip('home-undo', { key: 'Z', label: 'Undo', tab: 'H', handler: onUndo });
        keyTips.registerKeyTip('home-redo', { key: 'Y', label: 'Redo', tab: 'H', handler: onRedo });
        keyTips.registerKeyTip('home-font-color', { key: 'O', label: 'Font Color', tab: 'H', handler: () => fontColorBtnRef.current?.click() });
        keyTips.registerKeyTip('home-fill-color', { key: 'G', label: 'Fill Color', tab: 'H', handler: () => fillColorBtnRef.current?.click() });

        // Insert tab commands
        keyTips.registerKeyTip('insert-chart', { key: 'C', label: 'Chart', tab: 'N', handler: () => {} });
        keyTips.registerKeyTip('insert-image', { key: 'I', label: 'Image', tab: 'N', handler: () => imageInputRef.current?.click() });

        // Data tab commands
        keyTips.registerKeyTip('data-filter', { key: 'F', label: 'Filter', tab: 'A', handler: () => {} });
        keyTips.registerKeyTip('data-sort', { key: 'S', label: 'Sort', tab: 'A', handler: () => {} });

        return () => {
            keyTips.clearKeyTips();
        };
    }, [keyTips, onUndo, onRedo, onApplyFormat, currentFormat]);

    const handleShareSubmit = () => {
        if (shareEmail.trim()) {
            onShare(shareEmail, sharePermission);
            setShowShareDialog(false);
            setShareEmail('');
        }
    };

    const handleFontColorSelect = (color: string) => {
        onApplyFontColor(color);
        setFontColorOpen(false);
    };

    const handleFillColorSelect = (color: string) => {
        onApplyFillColor(color);
        setFillColorOpen(false);
    };

    const handleExportUnifiedSource = () => {
        exportUnifiedSourceCode(spreadsheetName || 'spreadsheet');
    };

    const cellReference = selectedCell
        ? `${getColumnLabel(selectedCell.col)}${selectedCell.row + 1}`
        : '';

    const showTabKeyTips = keyTips.isActive && keyTips.state === 'awaitingTab';
    const showCommandKeyTips = keyTips.isActive && keyTips.state === 'awaitingCommand';

    return (
        <TooltipProvider>
            <div className="border-b bg-card">
                {/* KeyTip Overlay */}
                <KeyTipOverlay sequence={keyTips.currentSequence} visible={keyTips.isActive} />

                {/* Quick Access Toolbar & Tabs */}
                <div className="flex items-center justify-between px-2 py-1 border-b">
                    <div className="flex items-center gap-1">
                        <Tooltip>
                            <TooltipTrigger asChild>
                                <Button
                                    variant="ghost"
                                    size="icon"
                                    className="h-7 w-7"
                                    onClick={onSave}
                                    disabled={!hasUnsavedChanges || isSaving}
                                    aria-label="Save"
                                >
                                    <Save className="h-4 w-4" />
                                </Button>
                            </TooltipTrigger>
                            <TooltipContent>Save (Ctrl+S)</TooltipContent>
                        </Tooltip>
                        <Tooltip>
                            <TooltipTrigger asChild>
                                <Button
                                    ref={undoBtnRef}
                                    variant="ghost"
                                    size="icon"
                                    className="h-7 w-7"
                                    onClick={onUndo}
                                    disabled={!canUndo}
                                    aria-label="Undo"
                                >
                                    <Undo className="h-4 w-4" />
                                </Button>
                            </TooltipTrigger>
                            <TooltipContent>Undo (Ctrl+Z)</TooltipContent>
                        </Tooltip>
                        <Tooltip>
                            <TooltipTrigger asChild>
                                <Button
                                    ref={redoBtnRef}
                                    variant="ghost"
                                    size="icon"
                                    className="h-7 w-7"
                                    onClick={onRedo}
                                    disabled={!canRedo}
                                    aria-label="Redo"
                                >
                                    <Redo className="h-4 w-4" />
                                </Button>
                            </TooltipTrigger>
                            <TooltipContent>Redo (Ctrl+Y)</TooltipContent>
                        </Tooltip>
                        <Tooltip>
                            <TooltipTrigger asChild>
                                <Button
                                    variant="ghost"
                                    size="icon"
                                    className="h-7 w-7"
                                    onClick={() => window.print()}
                                    aria-label="Print"
                                >
                                    <Printer className="h-4 w-4" />
                                </Button>
                            </TooltipTrigger>
                            <TooltipContent>Print (Ctrl+P)</TooltipContent>
                        </Tooltip>
                        <Dialog open={showShareDialog} onOpenChange={setShowShareDialog}>
                            <DialogTrigger asChild>
                                <Tooltip>
                                    <TooltipTrigger asChild>
                                        <Button
                                            variant="ghost"
                                            size="icon"
                                            className="h-7 w-7"
                                            aria-label="Share"
                                        >
                                            <Share2 className="h-4 w-4" />
                                        </Button>
                                    </TooltipTrigger>
                                    <TooltipContent>Share</TooltipContent>
                                </Tooltip>
                            </DialogTrigger>
                            <DialogContent>
                                <DialogHeader>
                                    <DialogTitle>Share Spreadsheet</DialogTitle>
                                </DialogHeader>
                                <div className="space-y-4 py-4">
                                    <div className="space-y-2">
                                        <Label htmlFor="principal">Principal ID</Label>
                                        <Input
                                            id="principal"
                                            value={shareEmail}
                                            onChange={(e) => setShareEmail(e.target.value)}
                                            placeholder="Enter principal ID"
                                        />
                                    </div>
                                    <div className="space-y-2">
                                        <Label htmlFor="permission">Permission</Label>
                                        <Select
                                            value={sharePermission}
                                            onValueChange={(value) => setSharePermission(value as any)}
                                        >
                                            <SelectTrigger id="permission">
                                                <SelectValue />
                                            </SelectTrigger>
                                            <SelectContent>
                                                <SelectItem value="viewer">Viewer</SelectItem>
                                                <SelectItem value="editor">Editor</SelectItem>
                                            </SelectContent>
                                        </Select>
                                    </div>
                                    <Button onClick={handleShareSubmit} className="w-full">
                                        Share
                                    </Button>
                                </div>
                            </DialogContent>
                        </Dialog>
                    </div>

                    <Tabs value={activeTab} onValueChange={setActiveTab} className="w-auto">
                        <TabsList className="h-8 bg-transparent relative">
                            <TabsTrigger ref={fileTabRef} value="file" className="text-xs px-3 relative">
                                File
                                <KeyTipBadge keyTip="F" visible={showTabKeyTips} targetRef={fileTabRef} />
                            </TabsTrigger>
                            <TabsTrigger ref={homeTabRef} value="home" className="text-xs px-3 relative">
                                Home
                                <KeyTipBadge keyTip="H" visible={showTabKeyTips} targetRef={homeTabRef} />
                            </TabsTrigger>
                            <TabsTrigger ref={insertTabRef} value="insert" className="text-xs px-3 relative">
                                Insert
                                <KeyTipBadge keyTip="N" visible={showTabKeyTips} targetRef={insertTabRef} />
                            </TabsTrigger>
                            <TabsTrigger ref={formulasTabRef} value="formulas" className="text-xs px-3 relative">
                                Formulas
                                <KeyTipBadge keyTip="M" visible={showTabKeyTips} targetRef={formulasTabRef} />
                            </TabsTrigger>
                            <TabsTrigger ref={dataTabRef} value="data" className="text-xs px-3 relative">
                                Data
                                <KeyTipBadge keyTip="A" visible={showTabKeyTips} targetRef={dataTabRef} />
                            </TabsTrigger>
                            <TabsTrigger ref={viewTabRef} value="view" className="text-xs px-3 relative">
                                View
                                <KeyTipBadge keyTip="W" visible={showTabKeyTips} targetRef={viewTabRef} />
                            </TabsTrigger>
                            <TabsTrigger ref={helpTabRef} value="help" className="text-xs px-3 relative">
                                Help
                                <KeyTipBadge keyTip="E" visible={showTabKeyTips} targetRef={helpTabRef} />
                            </TabsTrigger>
                        </TabsList>
                    </Tabs>
                </div>

                {/* Ribbon Content */}
                <div className="px-4 py-2">
                    {activeTab === 'file' && (
                        <div className="flex items-center gap-2">
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Export</span>
                                <div className="flex gap-1">
                                    <Button variant="outline" size="sm" onClick={onExportCSV}>
                                        <Download className="mr-2 h-4 w-4" />
                                        CSV
                                    </Button>
                                    <Button variant="outline" size="sm" onClick={onExportXLSX}>
                                        <Download className="mr-2 h-4 w-4" />
                                        XLSX
                                    </Button>
                                    <Button variant="outline" size="sm" onClick={onExportJSON}>
                                        <Download className="mr-2 h-4 w-4" />
                                        JSON
                                    </Button>
                                    <Button variant="outline" size="sm" onClick={handleExportUnifiedSource}>
                                        <Code className="mr-2 h-4 w-4" />
                                        Source Code
                                    </Button>
                                </div>
                            </div>
                            <Separator orientation="vertical" className="h-12" />
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Import</span>
                                <div className="flex gap-1">
                                    <label>
                                        <Button variant="outline" size="sm" asChild>
                                            <span>
                                                <Upload className="mr-2 h-4 w-4" />
                                                Import File
                                            </span>
                                        </Button>
                                        <input
                                            type="file"
                                            accept=".csv,.json"
                                            onChange={onImport}
                                            className="hidden"
                                        />
                                    </label>
                                    <label>
                                        <Button variant="outline" size="sm" asChild>
                                            <span>
                                                <FileSpreadsheet className="mr-2 h-4 w-4" />
                                                Upload Excel
                                            </span>
                                        </Button>
                                        <input
                                            ref={excelInputRef}
                                            type="file"
                                            accept=".xlsx"
                                            onChange={onImportExcel}
                                            className="hidden"
                                        />
                                    </label>
                                </div>
                            </div>
                            <Separator orientation="vertical" className="h-12" />
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Delete</span>
                                <Button variant="destructive" size="sm" onClick={onDeleteSpreadsheet}>
                                    <Trash2 className="mr-2 h-4 w-4" />
                                    Delete Spreadsheet
                                </Button>
                            </div>
                        </div>
                    )}

                    {activeTab === 'home' && (
                        <div className="flex items-center gap-2">
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Clipboard</span>
                                <div className="flex gap-1">
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button variant="outline" size="sm">
                                                <FileText className="h-4 w-4" />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Copy (Ctrl+C)</TooltipContent>
                                    </Tooltip>
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button variant="outline" size="sm">
                                                <FileText className="h-4 w-4" />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Paste (Ctrl+V)</TooltipContent>
                                    </Tooltip>
                                </div>
                            </div>
                            <Separator orientation="vertical" className="h-12" />
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Font</span>
                                <div className="flex gap-1">
                                    <Select defaultValue="arial">
                                        <SelectTrigger ref={fontSelectRef} className="h-8 w-32 relative">
                                            <SelectValue />
                                            <KeyTipBadge keyTip="F" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={fontSelectRef} />
                                        </SelectTrigger>
                                        <SelectContent>
                                            <SelectItem value="arial">Arial</SelectItem>
                                            <SelectItem value="times">Times New Roman</SelectItem>
                                            <SelectItem value="courier">Courier New</SelectItem>
                                        </SelectContent>
                                    </Select>
                                    <Select defaultValue="11">
                                        <SelectTrigger ref={sizeSelectRef} className="h-8 w-16 relative">
                                            <SelectValue />
                                            <KeyTipBadge keyTip="S" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={sizeSelectRef} />
                                        </SelectTrigger>
                                        <SelectContent>
                                            <SelectItem value="8">8</SelectItem>
                                            <SelectItem value="9">9</SelectItem>
                                            <SelectItem value="10">10</SelectItem>
                                            <SelectItem value="11">11</SelectItem>
                                            <SelectItem value="12">12</SelectItem>
                                            <SelectItem value="14">14</SelectItem>
                                            <SelectItem value="16">16</SelectItem>
                                        </SelectContent>
                                    </Select>
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button 
                                                ref={boldBtnRef} 
                                                variant={currentFormat.bold ? "default" : "outline"} 
                                                size="icon" 
                                                className="h-8 w-8 relative"
                                                onClick={() => onApplyFormat({ bold: !currentFormat.bold })}
                                            >
                                                <Bold className="h-4 w-4" />
                                                <KeyTipBadge keyTip="B" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={boldBtnRef} />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Bold (Ctrl+B)</TooltipContent>
                                    </Tooltip>
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button 
                                                ref={italicBtnRef} 
                                                variant={currentFormat.italic ? "default" : "outline"} 
                                                size="icon" 
                                                className="h-8 w-8 relative"
                                                onClick={() => onApplyFormat({ italic: !currentFormat.italic })}
                                            >
                                                <Italic className="h-4 w-4" />
                                                <KeyTipBadge keyTip="I" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={italicBtnRef} />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Italic (Ctrl+I)</TooltipContent>
                                    </Tooltip>
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button 
                                                ref={underlineBtnRef} 
                                                variant={currentFormat.underline ? "default" : "outline"} 
                                                size="icon" 
                                                className="h-8 w-8 relative"
                                                onClick={() => onApplyFormat({ underline: !currentFormat.underline })}
                                            >
                                                <Underline className="h-4 w-4" />
                                                <KeyTipBadge keyTip="U" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={underlineBtnRef} />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Underline (Ctrl+U)</TooltipContent>
                                    </Tooltip>
                                    <Popover open={fontColorOpen} onOpenChange={setFontColorOpen}>
                                        <PopoverTrigger asChild>
                                            <Tooltip>
                                                <TooltipTrigger asChild>
                                                    <Button ref={fontColorBtnRef} variant="outline" size="icon" className="h-8 w-8 relative">
                                                        <img src="/assets/generated/font-color-icon.dim_24x24.png" alt="Font Color" className="h-4 w-4" />
                                                        <KeyTipBadge keyTip="O" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={fontColorBtnRef} />
                                                    </Button>
                                                </TooltipTrigger>
                                                <TooltipContent>Font Color</TooltipContent>
                                            </Tooltip>
                                        </PopoverTrigger>
                                        <PopoverContent className="w-auto p-2" align="start">
                                            <div className="grid grid-cols-5 gap-1">
                                                {COLOR_PALETTE.map((color) => (
                                                    <button
                                                        key={color.value}
                                                        className="h-8 w-8 rounded border border-border hover:ring-2 hover:ring-primary transition-all"
                                                        style={{ backgroundColor: color.value }}
                                                        onClick={() => handleFontColorSelect(color.value)}
                                                        title={color.name}
                                                        aria-label={`Font color ${color.name}`}
                                                    />
                                                ))}
                                            </div>
                                        </PopoverContent>
                                    </Popover>
                                    <Popover open={fillColorOpen} onOpenChange={setFillColorOpen}>
                                        <PopoverTrigger asChild>
                                            <Tooltip>
                                                <TooltipTrigger asChild>
                                                    <Button ref={fillColorBtnRef} variant="outline" size="icon" className="h-8 w-8 relative">
                                                        <img src="/assets/generated/fill-color-icon.dim_24x24.png" alt="Fill Color" className="h-4 w-4" />
                                                        <KeyTipBadge keyTip="G" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={fillColorBtnRef} />
                                                    </Button>
                                                </TooltipTrigger>
                                                <TooltipContent>Fill Color</TooltipContent>
                                            </Tooltip>
                                        </PopoverTrigger>
                                        <PopoverContent className="w-auto p-2" align="start">
                                            <div className="grid grid-cols-5 gap-1">
                                                {COLOR_PALETTE.map((color) => (
                                                    <button
                                                        key={color.value}
                                                        className="h-8 w-8 rounded border border-border hover:ring-2 hover:ring-primary transition-all"
                                                        style={{ backgroundColor: color.value }}
                                                        onClick={() => handleFillColorSelect(color.value)}
                                                        title={color.name}
                                                        aria-label={`Fill color ${color.name}`}
                                                    />
                                                ))}
                                            </div>
                                        </PopoverContent>
                                    </Popover>
                                </div>
                            </div>
                            <Separator orientation="vertical" className="h-12" />
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Alignment</span>
                                <div className="flex gap-1">
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button 
                                                ref={alignLeftBtnRef} 
                                                variant={currentFormat.alignment === 'left' ? "default" : "outline"} 
                                                size="icon" 
                                                className="h-8 w-8 relative"
                                                onClick={() => onApplyFormat({ alignment: 'left' })}
                                            >
                                                <AlignLeft className="h-4 w-4" />
                                                <KeyTipBadge keyTip="L" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={alignLeftBtnRef} />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Align Left</TooltipContent>
                                    </Tooltip>
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button 
                                                ref={alignCenterBtnRef} 
                                                variant={currentFormat.alignment === 'center' ? "default" : "outline"} 
                                                size="icon" 
                                                className="h-8 w-8 relative"
                                                onClick={() => onApplyFormat({ alignment: 'center' })}
                                            >
                                                <AlignCenter className="h-4 w-4" />
                                                <KeyTipBadge keyTip="C" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={alignCenterBtnRef} />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Align Center</TooltipContent>
                                    </Tooltip>
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button 
                                                ref={alignRightBtnRef} 
                                                variant={currentFormat.alignment === 'right' ? "default" : "outline"} 
                                                size="icon" 
                                                className="h-8 w-8 relative"
                                                onClick={() => onApplyFormat({ alignment: 'right' })}
                                            >
                                                <AlignRight className="h-4 w-4" />
                                                <KeyTipBadge keyTip="R" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={alignRightBtnRef} />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Align Right</TooltipContent>
                                    </Tooltip>
                                    <Tooltip>
                                        <TooltipTrigger asChild>
                                            <Button ref={mergeBtnRef} variant="outline" size="icon" className="h-8 w-8 relative">
                                                <Merge className="h-4 w-4" />
                                                <KeyTipBadge keyTip="M" visible={showCommandKeyTips && keyTips.activeTab === 'H'} targetRef={mergeBtnRef} />
                                            </Button>
                                        </TooltipTrigger>
                                        <TooltipContent>Merge Cells</TooltipContent>
                                    </Tooltip>
                                </div>
                            </div>
                            <Separator orientation="vertical" className="h-12" />
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Number</span>
                                <Select defaultValue="general">
                                    <SelectTrigger className="h-8 w-32">
                                        <SelectValue />
                                    </SelectTrigger>
                                    <SelectContent>
                                        <SelectItem value="general">General</SelectItem>
                                        <SelectItem value="number">Number</SelectItem>
                                        <SelectItem value="currency">Currency</SelectItem>
                                        <SelectItem value="percentage">Percentage</SelectItem>
                                        <SelectItem value="date">Date</SelectItem>
                                    </SelectContent>
                                </Select>
                            </div>
                        </div>
                    )}

                    {activeTab === 'insert' && (
                        <div className="flex items-center gap-2">
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Illustrations</span>
                                <div className="flex gap-1">
                                    <label>
                                        <Button ref={imageBtnRef} variant="outline" size="sm" asChild className="relative">
                                            <span>
                                                <ImageIcon className="mr-2 h-4 w-4" />
                                                Insert Image
                                                <KeyTipBadge keyTip="I" visible={showCommandKeyTips && keyTips.activeTab === 'N'} targetRef={imageBtnRef} />
                                            </span>
                                        </Button>
                                        <input
                                            ref={imageInputRef}
                                            type="file"
                                            accept="image/png,image/jpeg,image/jpg"
                                            onChange={onInsertImage}
                                            className="hidden"
                                        />
                                    </label>
                                </div>
                            </div>
                            <Separator orientation="vertical" className="h-12" />
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Charts</span>
                                <Button ref={chartBtnRef} variant="outline" size="sm" className="relative">
                                    <BarChart3 className="mr-2 h-4 w-4" />
                                    Insert Chart
                                    <KeyTipBadge keyTip="C" visible={showCommandKeyTips && keyTips.activeTab === 'N'} targetRef={chartBtnRef} />
                                </Button>
                            </div>
                        </div>
                    )}

                    {activeTab === 'formulas' && (
                        <div className="flex items-center gap-2">
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Function Library</span>
                                <div className="flex gap-1">
                                    <DropdownMenu>
                                        <DropdownMenuTrigger asChild>
                                            <Button variant="outline" size="sm">
                                                <Type className="mr-2 h-4 w-4" />
                                                Text
                                                <ChevronDown className="ml-2 h-4 w-4" />
                                            </Button>
                                        </DropdownMenuTrigger>
                                        <DropdownMenuContent>
                                            <DropdownMenuItem>CONCAT</DropdownMenuItem>
                                            <DropdownMenuItem>LEN</DropdownMenuItem>
                                            <DropdownMenuItem>UPPER</DropdownMenuItem>
                                            <DropdownMenuItem>LOWER</DropdownMenuItem>
                                        </DropdownMenuContent>
                                    </DropdownMenu>
                                    <DropdownMenu>
                                        <DropdownMenuTrigger asChild>
                                            <Button variant="outline" size="sm">
                                                <Hash className="mr-2 h-4 w-4" />
                                                Math
                                                <ChevronDown className="ml-2 h-4 w-4" />
                                            </Button>
                                        </DropdownMenuTrigger>
                                        <DropdownMenuContent>
                                            <DropdownMenuItem>SUM</DropdownMenuItem>
                                            <DropdownMenuItem>AVERAGE</DropdownMenuItem>
                                            <DropdownMenuItem>MIN</DropdownMenuItem>
                                            <DropdownMenuItem>MAX</DropdownMenuItem>
                                            <DropdownMenuItem>ROUND</DropdownMenuItem>
                                        </DropdownMenuContent>
                                    </DropdownMenu>
                                </div>
                            </div>
                        </div>
                    )}

                    {activeTab === 'data' && (
                        <div className="flex items-center gap-2">
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Sort & Filter</span>
                                <div className="flex gap-1">
                                    <Button ref={sortBtnRef} variant="outline" size="sm" className="relative">
                                        <ArrowUpDown className="mr-2 h-4 w-4" />
                                        Sort
                                        <KeyTipBadge keyTip="S" visible={showCommandKeyTips && keyTips.activeTab === 'A'} targetRef={sortBtnRef} />
                                    </Button>
                                    <Button ref={filterBtnRef} variant="outline" size="sm" className="relative">
                                        <Filter className="mr-2 h-4 w-4" />
                                        Filter
                                        <KeyTipBadge keyTip="F" visible={showCommandKeyTips && keyTips.activeTab === 'A'} targetRef={filterBtnRef} />
                                    </Button>
                                </div>
                            </div>
                        </div>
                    )}

                    {activeTab === 'view' && (
                        <div className="flex items-center gap-2">
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Show</span>
                                <div className="flex gap-1">
                                    <Button variant="outline" size="sm">
                                        <Grid3x3 className="mr-2 h-4 w-4" />
                                        Gridlines
                                    </Button>
                                </div>
                            </div>
                        </div>
                    )}

                    {activeTab === 'help' && (
                        <div className="flex items-center gap-2">
                            <div className="flex flex-col gap-1">
                                <span className="text-xs text-muted-foreground">Help</span>
                                <div className="flex gap-1">
                                    <Button variant="outline" size="sm" onClick={() => setShowShortcutsDialog(true)}>
                                        <Keyboard className="mr-2 h-4 w-4" />
                                        Keyboard Shortcuts
                                    </Button>
                                    <Button variant="outline" size="sm" onClick={() => setShowSettingsDialog(true)}>
                                        <Settings className="mr-2 h-4 w-4" />
                                        KeyTip Settings
                                    </Button>
                                </div>
                            </div>
                        </div>
                    )}
                </div>

                {/* Name Box and Formula Bar */}
                <div className="flex items-center gap-2 border-t px-4 py-2 bg-muted/30">
                    <div className="flex items-center gap-2">
                        <Label htmlFor="name-box" className="text-xs text-muted-foreground sr-only">
                            Name Box
                        </Label>
                        <Input
                            id="name-box"
                            value={cellReference}
                            readOnly
                            className="h-8 w-24 text-sm font-medium bg-background"
                            aria-label="Cell reference"
                        />
                        <Separator orientation="vertical" className="h-6" />
                    </div>
                    <div className="flex-1 flex items-center gap-2">
                        <Label htmlFor="formula-bar" className="text-xs text-muted-foreground sr-only">
                            Formula Bar
                        </Label>
                        <Input
                            id="formula-bar"
                            ref={inputRef}
                            value={formulaBarValue}
                            onChange={(e) => onFormulaBarChange(e.target.value)}
                            onFocus={() => setIsEditingCell(true)}
                            onBlur={() => setIsEditingCell(false)}
                            onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                    onFormulaBarSubmit();
                                } else if (e.key === 'Escape') {
                                    setIsEditingCell(false);
                                    inputRef.current?.blur();
                                }
                            }}
                            placeholder="Enter value or formula (e.g. =SUM(A1:A10))"
                            className="flex-1 h-8 text-sm bg-background"
                            aria-label="Formula bar"
                        />
                        <Button
                            size="sm"
                            onClick={onFormulaBarSubmit}
                            disabled={!hasUnsavedChanges || isSaving}
                            className="h-8"
                        >
                            <Save className="mr-2 h-4 w-4" />
                            {isSaving ? 'Saving...' : 'Save'}
                        </Button>
                    </div>
                </div>
            </div>

            {/* Dialogs */}
            <KeyboardShortcutsDialog open={showShortcutsDialog} onOpenChange={setShowShortcutsDialog} />
            <KeyTipSettingsDialog
                open={showSettingsDialog}
                onOpenChange={setShowSettingsDialog}
                settings={keyTips.settings}
                onSettingsChange={keyTips.saveSettings}
            />
        </TooltipProvider>
    );
}
