import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription } from '@/components/ui/dialog';
import { Label } from '@/components/ui/label';
import { Switch } from '@/components/ui/switch';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import type { KeyTipSettings } from '../hooks/useKeyTips';

interface KeyTipSettingsDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  settings: KeyTipSettings;
  onSettingsChange: (settings: KeyTipSettings) => void;
}

export function KeyTipSettingsDialog({
  open,
  onOpenChange,
  settings,
  onSettingsChange,
}: KeyTipSettingsDialogProps) {
  const isMac = /Mac|iPhone|iPad|iPod/.test(navigator.userAgent);

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>KeyTips Settings</DialogTitle>
          <DialogDescription>
            Configure KeyTips behavior and activation keys
          </DialogDescription>
        </DialogHeader>
        <div className="space-y-6 py-4">
          <div className="flex items-center justify-between">
            <div className="space-y-0.5">
              <Label htmlFor="keytips-enabled">Enable KeyTips</Label>
              <p className="text-sm text-muted-foreground">
                Show keyboard shortcuts when Alt is pressed
              </p>
            </div>
            <Switch
              id="keytips-enabled"
              checked={settings.enabled}
              onCheckedChange={(enabled) =>
                onSettingsChange({ ...settings, enabled })
              }
            />
          </div>

          {isMac && (
            <div className="space-y-3">
              <Label>macOS Activation Key</Label>
              <RadioGroup
                value={settings.macModifier}
                onValueChange={(value) =>
                  onSettingsChange({
                    ...settings,
                    macModifier: value as KeyTipSettings['macModifier'],
                  })
                }
              >
                <div className="flex items-center space-x-2">
                  <RadioGroupItem value="ctrlOption" id="ctrl-option" />
                  <Label htmlFor="ctrl-option" className="font-normal">
                    Ctrl + Option (recommended)
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <RadioGroupItem value="alt" id="alt" />
                  <Label htmlFor="alt" className="font-normal">
                    Option/Alt only
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <RadioGroupItem value="cmd" id="cmd" />
                  <Label htmlFor="cmd" className="font-normal">
                    Command (⌘)
                  </Label>
                </div>
              </RadioGroup>
              <p className="text-xs text-muted-foreground">
                Choose which key combination activates KeyTips on macOS
              </p>
            </div>
          )}

          <div className="p-4 bg-muted rounded-lg space-y-2">
            <h4 className="font-semibold text-sm">About KeyTips</h4>
            <p className="text-xs text-muted-foreground">
              KeyTips provide Excel-style keyboard navigation for the ribbon interface. 
              Press the activation key to see available shortcuts, then press the 
              corresponding letter to activate that command.
            </p>
            <p className="text-xs text-muted-foreground">
              Press Esc to cancel KeyTip mode, or wait 2.5 seconds for automatic timeout.
            </p>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
