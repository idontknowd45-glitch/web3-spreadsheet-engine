import { useState, useRef, useEffect } from 'react';
import { X, Maximize2 } from 'lucide-react';
import { Button } from '@/components/ui/button';

export interface ImageData {
    id: string;
    src: string;
    x: number;
    y: number;
    width: number;
    height: number;
    anchorCell: string;
}

interface ImageLayerProps {
    images: ImageData[];
    onImageUpdate: (image: ImageData) => void;
    onImageDelete: (imageId: string) => void;
    gridOffset: { top: number; left: number };
    zoom: number;
}

export default function ImageLayer({
    images,
    onImageUpdate,
    onImageDelete,
    gridOffset,
    zoom,
}: ImageLayerProps) {
    const [selectedImageId, setSelectedImageId] = useState<string | null>(null);
    const [isDragging, setIsDragging] = useState(false);
    const [isResizing, setIsResizing] = useState(false);
    const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
    const [resizeStart, setResizeStart] = useState({ x: 0, y: 0, width: 0, height: 0 });

    const scale = zoom / 100;

    const handleImageMouseDown = (e: React.MouseEvent, image: ImageData) => {
        e.stopPropagation();
        setSelectedImageId(image.id);
        setIsDragging(true);
        setDragStart({
            x: e.clientX - image.x,
            y: e.clientY - image.y,
        });
    };

    const handleResizeMouseDown = (e: React.MouseEvent, image: ImageData) => {
        e.stopPropagation();
        setIsResizing(true);
        setResizeStart({
            x: e.clientX,
            y: e.clientY,
            width: image.width,
            height: image.height,
        });
    };

    const handleMouseMove = (e: MouseEvent) => {
        if (isDragging && selectedImageId) {
            const image = images.find((img) => img.id === selectedImageId);
            if (!image) return;

            const newX = e.clientX - dragStart.x;
            const newY = e.clientY - dragStart.y;

            onImageUpdate({
                ...image,
                x: newX,
                y: newY,
            });
        } else if (isResizing && selectedImageId) {
            const image = images.find((img) => img.id === selectedImageId);
            if (!image) return;

            const deltaX = e.clientX - resizeStart.x;
            const deltaY = e.clientY - resizeStart.y;

            const newWidth = Math.max(50, resizeStart.width + deltaX);
            const newHeight = Math.max(50, resizeStart.height + deltaY);

            onImageUpdate({
                ...image,
                width: newWidth,
                height: newHeight,
            });
        }
    };

    const handleMouseUp = () => {
        setIsDragging(false);
        setIsResizing(false);
    };

    useEffect(() => {
        if (isDragging || isResizing) {
            window.addEventListener('mousemove', handleMouseMove);
            window.addEventListener('mouseup', handleMouseUp);

            return () => {
                window.removeEventListener('mousemove', handleMouseMove);
                window.removeEventListener('mouseup', handleMouseUp);
            };
        }
    }, [isDragging, isResizing, selectedImageId, dragStart, resizeStart, images]);

    const handleBackgroundClick = () => {
        setSelectedImageId(null);
    };

    return (
        <div
            className="absolute inset-0 pointer-events-none z-10"
            onClick={handleBackgroundClick}
        >
            {images.map((image) => {
                const isSelected = selectedImageId === image.id;

                return (
                    <div
                        key={image.id}
                        className="absolute pointer-events-auto"
                        style={{
                            left: `${image.x}px`,
                            top: `${image.y}px`,
                            width: `${image.width}px`,
                            height: `${image.height}px`,
                            cursor: isDragging ? 'grabbing' : 'grab',
                            border: isSelected ? '2px solid oklch(var(--primary))' : '1px solid transparent',
                            boxShadow: isSelected ? '0 0 0 1px oklch(var(--primary))' : 'none',
                        }}
                        onMouseDown={(e) => handleImageMouseDown(e, image)}
                    >
                        <img
                            src={image.src}
                            alt="Spreadsheet image"
                            className="w-full h-full object-contain pointer-events-none select-none"
                            draggable={false}
                        />

                        {isSelected && (
                            <>
                                {/* Delete button */}
                                <Button
                                    variant="destructive"
                                    size="icon"
                                    className="absolute -top-2 -right-2 h-6 w-6 rounded-full shadow-lg"
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        onImageDelete(image.id);
                                        setSelectedImageId(null);
                                    }}
                                >
                                    <X className="h-3 w-3" />
                                </Button>

                                {/* Resize handle */}
                                <div
                                    className="absolute bottom-0 right-0 w-4 h-4 bg-primary cursor-nwse-resize"
                                    style={{
                                        transform: 'translate(50%, 50%)',
                                    }}
                                    onMouseDown={(e) => handleResizeMouseDown(e, image)}
                                >
                                    <Maximize2 className="h-3 w-3 text-primary-foreground" />
                                </div>

                                {/* Corner resize handles */}
                                <div
                                    className="absolute top-0 left-0 w-2 h-2 bg-primary rounded-full cursor-nwse-resize"
                                    style={{ transform: 'translate(-50%, -50%)' }}
                                />
                                <div
                                    className="absolute top-0 right-0 w-2 h-2 bg-primary rounded-full cursor-nesw-resize"
                                    style={{ transform: 'translate(50%, -50%)' }}
                                />
                                <div
                                    className="absolute bottom-0 left-0 w-2 h-2 bg-primary rounded-full cursor-nesw-resize"
                                    style={{ transform: 'translate(-50%, 50%)' }}
                                />
                            </>
                        )}
                    </div>
                );
            })}
        </div>
    );
}
