import { useRef, useEffect, useState, useCallback } from 'react';
import { ContextMenu, ContextMenuContent, ContextMenuItem, ContextMenuTrigger, ContextMenuSeparator } from '@/components/ui/context-menu';
import type { CellFormat } from '../backend';

interface ExcelGridProps {
    rows: number;
    cols: number;
    selectedCell: { row: number; col: number } | null;
    onCellClick: (row: number, col: number, e?: React.MouseEvent) => void;
    getCellDisplay: (row: number, col: number) => string;
    getCellFormat?: (row: number, col: number) => CellFormat | undefined;
    isCellSelected: (row: number, col: number) => boolean;
    onCopy: () => void;
    onCut: () => void;
    onPaste: () => void;
    zoom: number;
    onCellEdit?: (row: number, col: number) => void;
    onSelectionChange?: (start: { row: number; col: number }, end: { row: number; col: number } | null) => void;
    onDragFill?: (startRow: number, startCol: number, endRow: number, endCol: number) => void;
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

/**
 * Converts CellFormat to CSS styles for inline application.
 * Applies fontColor to color and fillColor to backgroundColor.
 */
function formatToCss(format?: CellFormat): React.CSSProperties {
    if (!format) return {};

    const style: React.CSSProperties = {};

    if (format.bold) style.fontWeight = 'bold';
    if (format.italic) style.fontStyle = 'italic';
    if (format.underline) style.textDecoration = 'underline';
    if (format.fontFamily) style.fontFamily = format.fontFamily;
    if (format.fontSize) style.fontSize = `${format.fontSize}px`;
    if (format.fontColor) style.color = format.fontColor;
    if (format.fillColor) style.backgroundColor = format.fillColor;
    if (format.alignment) style.textAlign = format.alignment as any;

    return style;
}

export default function ExcelGrid({
    rows,
    cols,
    selectedCell,
    onCellClick,
    getCellDisplay,
    getCellFormat,
    isCellSelected,
    onCopy,
    onCut,
    onPaste,
    zoom,
    onCellEdit,
    onSelectionChange,
    onDragFill,
}: ExcelGridProps) {
    const scale = zoom / 100;
    const containerRef = useRef<HTMLDivElement>(null);
    const scrollContainerRef = useRef<HTMLDivElement>(null);
    const [isDragging, setIsDragging] = useState(false);
    const [dragStart, setDragStart] = useState<{ row: number; col: number } | null>(null);
    const [dragEnd, setDragEnd] = useState<{ row: number; col: number } | null>(null);
    const [isDragFilling, setIsDragFilling] = useState(false);
    const [dragFillStart, setDragFillStart] = useState<{ row: number; col: number } | null>(null);
    const [hoveredCell, setHoveredCell] = useState<{ row: number; col: number } | null>(null);
    const [scrollOffset, setScrollOffset] = useState({ top: 0, left: 0 });
    const [visibleRange, setVisibleRange] = useState({ startRow: 0, endRow: 50, startCol: 0, endCol: 26 });
    
    // Touch gesture support
    const [touchStartDistance, setTouchStartDistance] = useState<number | null>(null);
    const [lastTapTime, setLastTapTime] = useState(0);
    const [lastTapCell, setLastTapCell] = useState<{ row: number; col: number } | null>(null);

    // Virtualization: Calculate visible cells based on scroll position
    const updateVisibleRange = useCallback(() => {
        if (!scrollContainerRef.current) return;

        const container = scrollContainerRef.current;
        const scrollTop = container.scrollTop;
        const scrollLeft = container.scrollLeft;
        const containerHeight = container.clientHeight;
        const containerWidth = container.clientWidth;

        const rowHeight = 24; // 6 * 4 (h-6 in Tailwind)
        const colWidth = 100; // min-w-[100px]

        const startRow = Math.max(0, Math.floor(scrollTop / rowHeight) - 5);
        const endRow = Math.min(rows, Math.ceil((scrollTop + containerHeight) / rowHeight) + 5);
        const startCol = Math.max(0, Math.floor(scrollLeft / colWidth) - 2);
        const endCol = Math.min(cols, Math.ceil((scrollLeft + containerWidth) / colWidth) + 2);

        setVisibleRange({ startRow, endRow, startCol, endCol });
        setScrollOffset({ top: scrollTop, left: scrollLeft });
    }, [rows, cols]);

    useEffect(() => {
        const container = scrollContainerRef.current;
        if (!container) return;

        const handleScroll = () => {
            requestAnimationFrame(updateVisibleRange);
        };

        container.addEventListener('scroll', handleScroll, { passive: true });
        updateVisibleRange();

        return () => container.removeEventListener('scroll', handleScroll);
    }, [updateVisibleRange]);

    // Mouse drag selection
    const handleMouseDown = (row: number, col: number, e: React.MouseEvent) => {
        if (e.button !== 0) return; // Only left click

        if (e.shiftKey && selectedCell) {
            // Shift+click: extend selection
            onSelectionChange?.(selectedCell, { row, col });
        } else if (e.ctrlKey || e.metaKey) {
            // Ctrl/Cmd+click: multi-range selection (simplified to single range for now)
            onCellClick(row, col, e);
        } else {
            // Normal click: start new selection
            setIsDragging(true);
            setDragStart({ row, col });
            setDragEnd(null);
            onCellClick(row, col, e);
        }
    };

    const handleMouseMove = (row: number, col: number) => {
        setHoveredCell({ row, col });

        if (isDragging && dragStart) {
            setDragEnd({ row, col });
            onSelectionChange?.(dragStart, { row, col });
        } else if (isDragFilling && dragFillStart) {
            setDragEnd({ row, col });
        }
    };

    const handleMouseUp = () => {
        if (isDragFilling && dragFillStart && dragEnd) {
            onDragFill?.(dragFillStart.row, dragFillStart.col, dragEnd.row, dragEnd.col);
        }
        setIsDragging(false);
        setIsDragFilling(false);
        setDragFillStart(null);
    };

    useEffect(() => {
        const handleGlobalMouseUp = () => handleMouseUp();
        window.addEventListener('mouseup', handleGlobalMouseUp);
        return () => window.removeEventListener('mouseup', handleGlobalMouseUp);
    }, [isDragFilling, dragFillStart, dragEnd]);

    // Drag fill handle
    const handleFillHandleMouseDown = (e: React.MouseEvent) => {
        e.stopPropagation();
        if (selectedCell) {
            setIsDragFilling(true);
            setDragFillStart(selectedCell);
        }
    };

    // Touch gestures
    const handleTouchStart = (e: React.TouchEvent) => {
        if (e.touches.length === 2) {
            // Pinch to zoom
            const touch1 = e.touches[0];
            const touch2 = e.touches[1];
            const distance = Math.hypot(
                touch2.clientX - touch1.clientX,
                touch2.clientY - touch1.clientY
            );
            setTouchStartDistance(distance);
        } else if (e.touches.length === 1) {
            // Double tap to edit
            const now = Date.now();
            const target = e.target as HTMLElement;
            const cellData = target.dataset.cell;
            
            if (cellData) {
                const [row, col] = cellData.split(',').map(Number);
                
                if (
                    lastTapCell &&
                    lastTapCell.row === row &&
                    lastTapCell.col === col &&
                    now - lastTapTime < 300
                ) {
                    // Double tap detected
                    onCellEdit?.(row, col);
                    setLastTapTime(0);
                    setLastTapCell(null);
                } else {
                    setLastTapTime(now);
                    setLastTapCell({ row, col });
                }
            }
        }
    };

    const handleTouchMove = (e: React.TouchEvent) => {
        if (e.touches.length === 2 && touchStartDistance !== null) {
            // Pinch zoom (would need to be connected to zoom prop)
            const touch1 = e.touches[0];
            const touch2 = e.touches[1];
            const distance = Math.hypot(
                touch2.clientX - touch1.clientX,
                touch2.clientY - touch1.clientY
            );
            const zoomDelta = (distance - touchStartDistance) / 10;
            // This would trigger zoom change in parent component
            console.log('Zoom delta:', zoomDelta);
        }
    };

    const handleTouchEnd = () => {
        setTouchStartDistance(null);
    };

    // Scroll to selected cell
    useEffect(() => {
        if (selectedCell && scrollContainerRef.current) {
            const container = scrollContainerRef.current;
            const rowHeight = 24;
            const colWidth = 100;
            const headerHeight = 24;
            const headerWidth = 48;

            const cellTop = selectedCell.row * rowHeight;
            const cellLeft = selectedCell.col * colWidth;
            const cellBottom = cellTop + rowHeight;
            const cellRight = cellLeft + colWidth;

            const scrollTop = container.scrollTop;
            const scrollLeft = container.scrollLeft;
            const viewHeight = container.clientHeight - headerHeight;
            const viewWidth = container.clientWidth - headerWidth;

            // Scroll vertically if needed
            if (cellTop < scrollTop) {
                container.scrollTop = cellTop;
            } else if (cellBottom > scrollTop + viewHeight) {
                container.scrollTop = cellBottom - viewHeight;
            }

            // Scroll horizontally if needed
            if (cellLeft < scrollLeft) {
                container.scrollLeft = cellLeft;
            } else if (cellRight > scrollLeft + viewWidth) {
                container.scrollLeft = cellRight - viewWidth;
            }
        }
    }, [selectedCell]);

    const rowHeight = 24;
    const colWidth = 100;
    const headerHeight = 24;
    const headerWidth = 48;

    return (
        <div 
            ref={containerRef}
            className="flex-1 overflow-hidden bg-background"
            onTouchStart={handleTouchStart}
            onTouchMove={handleTouchMove}
            onTouchEnd={handleTouchEnd}
        >
            <div
                ref={scrollContainerRef}
                className="h-full w-full overflow-auto"
                style={{ 
                    scrollBehavior: 'smooth',
                    willChange: 'scroll-position'
                }}
            >
                <div 
                    className="relative"
                    style={{ 
                        height: rows * rowHeight + headerHeight,
                        width: cols * colWidth + headerWidth,
                        transform: `scale(${scale})`,
                        transformOrigin: 'top left'
                    }}
                >
                    {/* Column headers */}
                    <div 
                        className="sticky top-0 left-0 z-30 flex"
                        style={{ height: headerHeight }}
                    >
                        <div 
                            className="sticky left-0 z-40 border-b border-r bg-muted"
                            style={{ width: headerWidth, height: headerHeight }}
                        />
                        {Array.from({ length: visibleRange.endCol - visibleRange.startCol }, (_, i) => {
                            const col = visibleRange.startCol + i;
                            return (
                                <div
                                    key={col}
                                    className="border-b border-r bg-muted text-xs font-bold text-center text-muted-foreground flex items-center justify-center"
                                    style={{ 
                                        width: colWidth,
                                        height: headerHeight,
                                        position: 'absolute',
                                        left: col * colWidth + headerWidth,
                                        top: 0
                                    }}
                                >
                                    {getColumnLabel(col)}
                                </div>
                            );
                        })}
                    </div>

                    {/* Row headers and cells */}
                    {Array.from({ length: visibleRange.endRow - visibleRange.startRow }, (_, i) => {
                        const row = visibleRange.startRow + i;
                        return (
                            <div 
                                key={row}
                                className="flex"
                                style={{ 
                                    position: 'absolute',
                                    top: row * rowHeight + headerHeight,
                                    left: 0,
                                    height: rowHeight
                                }}
                            >
                                {/* Row header */}
                                <div
                                    className="sticky left-0 z-20 border-b border-r bg-muted text-center text-xs font-bold text-muted-foreground flex items-center justify-center"
                                    style={{ width: headerWidth, height: rowHeight }}
                                >
                                    {row + 1}
                                </div>

                                {/* Cells */}
                                {Array.from({ length: visibleRange.endCol - visibleRange.startCol }, (_, j) => {
                                    const col = visibleRange.startCol + j;
                                    const isSelected = isCellSelected(row, col);
                                    const isActiveCell = selectedCell?.row === row && selectedCell?.col === col;
                                    const isHovered = hoveredCell?.row === row && hoveredCell?.col === col;
                                    const isDragSelected = isDragging && dragStart && dragEnd &&
                                        row >= Math.min(dragStart.row, dragEnd.row) &&
                                        row <= Math.max(dragStart.row, dragEnd.row) &&
                                        col >= Math.min(dragStart.col, dragEnd.col) &&
                                        col <= Math.max(dragStart.col, dragEnd.col);

                                    const cellFormat = getCellFormat?.(row, col);
                                    const cellStyle = formatToCss(cellFormat);
                                    const cellDisplay = getCellDisplay(row, col);

                                    return (
                                        <ContextMenu key={col}>
                                            <ContextMenuTrigger asChild>
                                                <div
                                                    data-cell={`${row},${col}`}
                                                    className={`border-b border-r px-2 text-sm flex items-center relative ${
                                                        isActiveCell
                                                            ? 'ring-2 ring-primary ring-inset z-10'
                                                            : isSelected || isDragSelected
                                                            ? 'bg-primary/10'
                                                            : isHovered
                                                            ? 'bg-accent/50'
                                                            : ''
                                                    }`}
                                                    style={{ 
                                                        width: colWidth,
                                                        height: rowHeight,
                                                        position: 'absolute',
                                                        left: col * colWidth + headerWidth,
                                                        cursor: 'cell',
                                                        ...cellStyle
                                                    }}
                                                    onMouseDown={(e) => handleMouseDown(row, col, e)}
                                                    onMouseEnter={() => handleMouseMove(row, col)}
                                                    onDoubleClick={() => onCellEdit?.(row, col)}
                                                    role="gridcell"
                                                    aria-selected={isSelected}
                                                    tabIndex={isActiveCell ? 0 : -1}
                                                >
                                                    {cellDisplay}
                                                    
                                                    {/* Drag fill handle */}
                                                    {isActiveCell && (
                                                        <div
                                                            className="absolute bottom-0 right-0 w-2 h-2 bg-primary cursor-crosshair z-20"
                                                            style={{ 
                                                                transform: 'translate(50%, 50%)',
                                                                borderRadius: '50%'
                                                            }}
                                                            onMouseDown={handleFillHandleMouseDown}
                                                            role="button"
                                                            aria-label="Drag to fill"
                                                        />
                                                    )}
                                                </div>
                                            </ContextMenuTrigger>
                                            <ContextMenuContent>
                                                <ContextMenuItem onClick={onCut}>
                                                    Cut
                                                    <span className="ml-auto text-xs text-muted-foreground">Ctrl+X</span>
                                                </ContextMenuItem>
                                                <ContextMenuItem onClick={onCopy}>
                                                    Copy
                                                    <span className="ml-auto text-xs text-muted-foreground">Ctrl+C</span>
                                                </ContextMenuItem>
                                                <ContextMenuItem onClick={onPaste}>
                                                    Paste
                                                    <span className="ml-auto text-xs text-muted-foreground">Ctrl+V</span>
                                                </ContextMenuItem>
                                                <ContextMenuSeparator />
                                                <ContextMenuItem>Insert Row</ContextMenuItem>
                                                <ContextMenuItem>Insert Column</ContextMenuItem>
                                                <ContextMenuItem>Delete Row</ContextMenuItem>
                                                <ContextMenuItem>Delete Column</ContextMenuItem>
                                                <ContextMenuSeparator />
                                                <ContextMenuItem>Format Cells...</ContextMenuItem>
                                                <ContextMenuItem>Insert Comment</ContextMenuItem>
                                                <ContextMenuItem>Clear Contents</ContextMenuItem>
                                                <ContextMenuItem>Cell Protection</ContextMenuItem>
                                            </ContextMenuContent>
                                        </ContextMenu>
                                    );
                                })}
                            </div>
                        );
                    })}
                </div>
            </div>
        </div>
    );
}
