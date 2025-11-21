import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription } from '@/components/ui/dialog';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Separator } from '@/components/ui/separator';

interface KeyboardShortcutsDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export function KeyboardShortcutsDialog({ open, onOpenChange }: KeyboardShortcutsDialogProps) {
  const isMac = /Mac|iPhone|iPad|iPod/.test(navigator.userAgent);
  const ctrlKey = isMac ? '⌘' : 'Ctrl';

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-3xl max-h-[80vh]">
        <DialogHeader>
          <DialogTitle>Keyboard Shortcuts</DialogTitle>
          <DialogDescription>
            Complete reference for keyboard shortcuts and KeyTips
          </DialogDescription>
        </DialogHeader>
        <Tabs defaultValue="shortcuts" className="w-full">
          <TabsList className="grid w-full grid-cols-2">
            <TabsTrigger value="shortcuts">Standard Shortcuts</TabsTrigger>
            <TabsTrigger value="keytips">KeyTips (Alt Mode)</TabsTrigger>
          </TabsList>
          
          <TabsContent value="shortcuts">
            <ScrollArea className="h-[500px] pr-4">
              <div className="space-y-6">
                <ShortcutSection
                  title="Navigation"
                  shortcuts={[
                    { keys: ['Arrow Keys'], description: 'Move between cells' },
                    { keys: [ctrlKey, 'Arrow'], description: 'Jump to edge of data region' },
                    { keys: [ctrlKey, 'Home'], description: 'Go to cell A1' },
                    { keys: ['Page Up/Down'], description: 'Scroll one page' },
                  ]}
                />
                <ShortcutSection
                  title="Editing"
                  shortcuts={[
                    { keys: ['Enter'], description: 'Confirm entry and move down' },
                    { keys: ['Shift', 'Enter'], description: 'Confirm entry and move up' },
                    { keys: ['Tab'], description: 'Confirm entry and move right' },
                    { keys: ['Shift', 'Tab'], description: 'Confirm entry and move left' },
                    { keys: ['F2'], description: 'Edit active cell' },
                    { keys: ['Esc'], description: 'Cancel editing' },
                    { keys: ['Delete'], description: 'Clear cell contents' },
                  ]}
                />
                <ShortcutSection
                  title="Selection"
                  shortcuts={[
                    { keys: ['Shift', 'Arrow'], description: 'Extend selection' },
                    { keys: [ctrlKey, 'A'], description: 'Select all cells' },
                    { keys: [ctrlKey, 'Space'], description: 'Select entire column' },
                    { keys: ['Shift', 'Space'], description: 'Select entire row' },
                  ]}
                />
                <ShortcutSection
                  title="Clipboard"
                  shortcuts={[
                    { keys: [ctrlKey, 'C'], description: 'Copy' },
                    { keys: [ctrlKey, 'X'], description: 'Cut' },
                    { keys: [ctrlKey, 'V'], description: 'Paste' },
                  ]}
                />
                <ShortcutSection
                  title="Formatting"
                  shortcuts={[
                    { keys: [ctrlKey, 'B'], description: 'Bold' },
                    { keys: [ctrlKey, 'I'], description: 'Italic' },
                    { keys: [ctrlKey, 'U'], description: 'Underline' },
                    { keys: [ctrlKey, '1'], description: 'Format cells dialog' },
                  ]}
                />
                <ShortcutSection
                  title="Other"
                  shortcuts={[
                    { keys: [ctrlKey, 'Z'], description: 'Undo' },
                    { keys: [ctrlKey, 'Y'], description: 'Redo' },
                    { keys: [ctrlKey, 'S'], description: 'Save' },
                    { keys: [ctrlKey, 'F'], description: 'Find' },
                    { keys: [ctrlKey, 'H'], description: 'Replace' },
                    { keys: [ctrlKey, 'PgUp/PgDn'], description: 'Switch sheets' },
                  ]}
                />
              </div>
            </ScrollArea>
          </TabsContent>

          <TabsContent value="keytips">
            <ScrollArea className="h-[500px] pr-4">
              <div className="space-y-6">
                <div className="space-y-2">
                  <h3 className="font-semibold">How to Use KeyTips</h3>
                  <p className="text-sm text-muted-foreground">
                    Press <kbd className="px-2 py-1 bg-muted rounded text-xs">Alt</kbd> 
                    {isMac && ' (or Ctrl+Option on Mac)'} to enter KeyTip mode. 
                    Small badges will appear showing available keys. Press the corresponding 
                    key to activate that tab or command.
                  </p>
                  <p className="text-sm text-muted-foreground">
                    Press <kbd className="px-2 py-1 bg-muted rounded text-xs">Esc</kbd> to 
                    cancel KeyTip mode at any time.
                  </p>
                </div>

                <Separator />

                <KeyTipSection
                  title="Ribbon Tabs"
                  keytips={[
                    { key: 'F', description: 'File tab - Export, import, delete' },
                    { key: 'H', description: 'Home tab - Formatting and editing' },
                    { key: 'N', description: 'Insert tab - Charts and objects' },
                    { key: 'M', description: 'Formulas tab - Function library' },
                    { key: 'A', description: 'Data tab - Sort and filter' },
                    { key: 'W', description: 'View tab - Display options' },
                  ]}
                />

                <KeyTipSection
                  title="Home Tab Commands (Alt → H → ...)"
                  keytips={[
                    { key: 'B', description: 'Bold formatting' },
                    { key: 'I', description: 'Italic formatting' },
                    { key: 'U', description: 'Underline formatting' },
                    { key: 'F', description: 'Font family selector' },
                    { key: 'S', description: 'Font size selector' },
                    { key: 'L', description: 'Align left' },
                    { key: 'C', description: 'Align center' },
                    { key: 'R', description: 'Align right' },
                    { key: 'M', description: 'Merge cells' },
                    { key: 'Z', description: 'Undo' },
                    { key: 'Y', description: 'Redo' },
                  ]}
                />

                <KeyTipSection
                  title="Insert Tab Commands (Alt → N → ...)"
                  keytips={[
                    { key: 'C', description: 'Insert chart' },
                  ]}
                />

                <KeyTipSection
                  title="Data Tab Commands (Alt → A → ...)"
                  keytips={[
                    { key: 'F', description: 'Filter data' },
                    { key: 'S', description: 'Sort data' },
                  ]}
                />

                <div className="mt-6 p-4 bg-muted rounded-lg">
                  <h4 className="font-semibold mb-2">Accessibility</h4>
                  <ul className="text-sm text-muted-foreground space-y-1">
                    <li>• KeyTips are announced by screen readers</li>
                    <li>• High contrast mode supported</li>
                    <li>• Can be disabled in settings if needed</li>
                    <li>• 2.5 second timeout for inactive sequences</li>
                  </ul>
                </div>
              </div>
            </ScrollArea>
          </TabsContent>
        </Tabs>
      </DialogContent>
    </Dialog>
  );
}

function ShortcutSection({ title, shortcuts }: { title: string; shortcuts: { keys: string[]; description: string }[] }) {
  return (
    <div className="space-y-2">
      <h3 className="font-semibold text-sm">{title}</h3>
      <div className="space-y-1">
        {shortcuts.map((shortcut, index) => (
          <div key={index} className="flex items-center justify-between text-sm">
            <span className="text-muted-foreground">{shortcut.description}</span>
            <div className="flex items-center gap-1">
              {shortcut.keys.map((key, i) => (
                <span key={i} className="flex items-center gap-1">
                  <kbd className="px-2 py-1 bg-muted rounded text-xs font-mono border border-border">
                    {key}
                  </kbd>
                  {i < shortcut.keys.length - 1 && <span className="text-muted-foreground">+</span>}
                </span>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function KeyTipSection({ title, keytips }: { title: string; keytips: { key: string; description: string }[] }) {
  return (
    <div className="space-y-2">
      <h3 className="font-semibold text-sm">{title}</h3>
      <div className="space-y-1">
        {keytips.map((keytip, index) => (
          <div key={index} className="flex items-center justify-between text-sm">
            <span className="text-muted-foreground">{keytip.description}</span>
            <kbd className="px-2 py-1 bg-primary text-primary-foreground rounded text-xs font-mono border border-primary">
              {keytip.key}
            </kbd>
          </div>
        ))}
      </div>
    </div>
  );
}
