import { Plus, Minus, ZoomIn, ChevronLeft, ChevronRight } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Slider } from '@/components/ui/slider';
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from '@/components/ui/dropdown-menu';

interface ExcelStatusBarProps {
    sheets: string[];
    activeSheet: string;
    onSheetChange: (sheet: string) => void;
    onAddSheet: () => void;
    onDeleteSheet: (sheet: string) => void;
    zoom: number;
    onZoomChange: (zoom: number) => void;
    cellCount: number;
    isAddingSheet: boolean;
}

export default function ExcelStatusBar({
    sheets,
    activeSheet,
    onSheetChange,
    onAddSheet,
    onDeleteSheet,
    zoom,
    onZoomChange,
    cellCount,
    isAddingSheet,
}: ExcelStatusBarProps) {
    return (
        <div className="border-t bg-card">
            {/* Sheet Tabs */}
            <div className="flex items-center gap-2 px-4 py-2 border-b">
                <div className="flex items-center gap-1">
                    <Button variant="ghost" size="icon" className="h-6 w-6">
                        <ChevronLeft className="h-4 w-4" />
                    </Button>
                    <Button variant="ghost" size="icon" className="h-6 w-6">
                        <ChevronRight className="h-4 w-4" />
                    </Button>
                </div>

                <div className="flex-1 flex items-center gap-2 overflow-x-auto">
                    {sheets.map((sheetName) => (
                        <DropdownMenu key={sheetName}>
                            <DropdownMenuTrigger asChild>
                                <Button
                                    variant={activeSheet === sheetName ? 'secondary' : 'ghost'}
                                    size="sm"
                                    className={`h-7 px-4 ${
                                        activeSheet === sheetName
                                            ? 'bg-secondary text-secondary-foreground font-medium border-b-2 border-primary'
                                            : 'hover:bg-accent'
                                    }`}
                                    onClick={(e) => {
                                        // Only switch sheet on left click, not on context menu
                                        if (e.button === 0) {
                                            onSheetChange(sheetName);
                                        }
                                    }}
                                    onContextMenu={(e) => {
                                        e.preventDefault();
                                    }}
                                >
                                    {sheetName}
                                </Button>
                            </DropdownMenuTrigger>
                            <DropdownMenuContent>
                                <DropdownMenuItem onClick={() => onDeleteSheet(sheetName)}>
                                    Delete Sheet
                                </DropdownMenuItem>
                                <DropdownMenuItem>Rename Sheet</DropdownMenuItem>
                            </DropdownMenuContent>
                        </DropdownMenu>
                    ))}

                    <Button
                        variant="ghost"
                        size="icon"
                        className="h-7 w-7"
                        onClick={onAddSheet}
                        disabled={isAddingSheet}
                        title="Add new sheet"
                    >
                        <Plus className="h-4 w-4" />
                    </Button>
                </div>
            </div>

            {/* Status Bar */}
            <div className="flex items-center justify-between px-4 py-1 text-xs text-muted-foreground">
                <div className="flex items-center gap-4">
                    <span>Cells: {cellCount}</span>
                </div>

                <div className="flex items-center gap-2">
                    <Button
                        variant="ghost"
                        size="icon"
                        className="h-6 w-6"
                        onClick={() => onZoomChange(Math.max(50, zoom - 10))}
                    >
                        <Minus className="h-3 w-3" />
                    </Button>
                    <Slider
                        value={[zoom]}
                        onValueChange={(value) => onZoomChange(value[0])}
                        min={50}
                        max={200}
                        step={10}
                        className="w-24"
                    />
                    <Button
                        variant="ghost"
                        size="icon"
                        className="h-6 w-6"
                        onClick={() => onZoomChange(Math.min(200, zoom + 10))}
                    >
                        <Plus className="h-3 w-3" />
                    </Button>
                    <span className="min-w-[3rem] text-center">{zoom}%</span>
                    <Button
                        variant="ghost"
                        size="icon"
                        className="h-6 w-6"
                        onClick={() => onZoomChange(100)}
                        title="Reset zoom"
                    >
                        <ZoomIn className="h-3 w-3" />
                    </Button>
                </div>
            </div>
        </div>
    );
}
