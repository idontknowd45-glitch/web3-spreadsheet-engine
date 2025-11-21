import { useState, useEffect } from 'react';
import { FileSpreadsheet, LogIn, LogOut, User, ChevronRight, ChevronLeft } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Separator } from '@/components/ui/separator';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription } from '@/components/ui/dialog';
import { Label } from '@/components/ui/label';
import { toast } from 'sonner';
import SpreadsheetGrid from './SpreadsheetGrid';
import { useListSpreadsheets, useCreateSpreadsheet, useGetCallerUserProfile, useSaveCallerUserProfile } from '../hooks/useQueries';
import { useInternetIdentity } from '../hooks/useInternetIdentity';
import { useQueryClient } from '@tanstack/react-query';

export default function SpreadsheetApp() {
    const [selectedSpreadsheetId, setSelectedSpreadsheetId] = useState<string | null>(null);
    const [newSpreadsheetName, setNewSpreadsheetName] = useState('');
    const [isCreating, setIsCreating] = useState(false);
    const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
    const [showProfileSetup, setShowProfileSetup] = useState(false);
    const [profileName, setProfileName] = useState('');

    const { login, clear, loginStatus, identity } = useInternetIdentity();
    const queryClient = useQueryClient();
    const { data: spreadsheets, isLoading } = useListSpreadsheets();
    const createSpreadsheet = useCreateSpreadsheet();
    const { data: userProfile, isLoading: profileLoading, isFetched } = useGetCallerUserProfile();
    const saveProfile = useSaveCallerUserProfile();

    const isAuthenticated = !!identity;
    const isLoggingIn = loginStatus === 'logging-in';

    useEffect(() => {
        if (spreadsheets && spreadsheets.length > 0 && !selectedSpreadsheetId) {
            setSelectedSpreadsheetId(spreadsheets[0].id);
        }
    }, [spreadsheets, selectedSpreadsheetId]);

    useEffect(() => {
        if (isAuthenticated && !profileLoading && isFetched && userProfile === null) {
            setShowProfileSetup(true);
        }
    }, [isAuthenticated, profileLoading, isFetched, userProfile]);

    const handleAuth = async () => {
        if (isAuthenticated) {
            await clear();
            queryClient.clear();
            setSelectedSpreadsheetId(null);
        } else {
            try {
                await login();
            } catch (error: any) {
                console.error('Login error:', error);
                if (error.message === 'User is already authenticated') {
                    await clear();
                    setTimeout(() => login(), 300);
                }
            }
        }
    };

    const handleSaveProfile = async () => {
        if (!profileName.trim()) {
            toast.error('Please enter your name');
            return;
        }

        try {
            await saveProfile.mutateAsync({ name: profileName });
            setShowProfileSetup(false);
            toast.success('Profile created successfully');
        } catch (error) {
            toast.error('Failed to create profile');
        }
    };

    const handleCreateSpreadsheet = async () => {
        if (!newSpreadsheetName.trim()) {
            toast.error('Please enter a spreadsheet name');
            return;
        }

        try {
            const id = await createSpreadsheet.mutateAsync(newSpreadsheetName);
            setSelectedSpreadsheetId(id);
            setNewSpreadsheetName('');
            setIsCreating(false);
            toast.success('Spreadsheet created successfully');
        } catch (error) {
            toast.error('Failed to create spreadsheet');
        }
    };

    if (!isAuthenticated) {
        return (
            <div className="flex h-screen flex-col items-center justify-center bg-gradient-to-br from-primary/5 via-background to-secondary/5">
                <div className="text-center space-y-6 max-w-md px-4">
                    <FileSpreadsheet className="h-20 w-20 text-primary mx-auto" />
                    <h1 className="text-4xl font-bold">Spreadsheet App</h1>
                    <p className="text-muted-foreground text-lg">
                        A powerful, production-ready spreadsheet application with advanced features
                    </p>
                    <Button
                        onClick={handleAuth}
                        disabled={isLoggingIn}
                        size="lg"
                        className="mt-8"
                    >
                        <LogIn className="mr-2 h-5 w-5" />
                        {isLoggingIn ? 'Logging in...' : 'Login with Internet Identity'}
                    </Button>
                </div>
            </div>
        );
    }

    return (
        <>
            <div className="flex h-screen flex-col bg-background">
                {/* Header */}
                <header className="flex items-center justify-between border-b bg-card px-4 py-2 shadow-sm">
                    <div className="flex items-center gap-3">
                        <FileSpreadsheet className="h-5 w-5 text-primary" />
                        <h1 className="text-lg font-semibold">Spreadsheet</h1>
                    </div>
                    <div className="flex items-center gap-3">
                        {userProfile && (
                            <div className="hidden sm:flex items-center gap-2 text-sm text-muted-foreground">
                                <User className="h-4 w-4" />
                                <span>{userProfile.name}</span>
                            </div>
                        )}
                        <Button
                            onClick={handleAuth}
                            variant="outline"
                            size="sm"
                        >
                            <LogOut className="mr-2 h-4 w-4" />
                            Logout
                        </Button>
                    </div>
                </header>

                <div className="flex flex-1 overflow-hidden">
                    {/* Collapsible Sidebar */}
                    <aside className={`border-r bg-card transition-all duration-300 ${sidebarCollapsed ? 'w-0' : 'w-64'} overflow-hidden`}>
                        <div className="flex h-full flex-col">
                            <div className="p-4">
                                <h2 className="mb-4 text-sm font-semibold text-muted-foreground">
                                    MY SPREADSHEETS
                                </h2>
                                <FileList
                                    spreadsheets={spreadsheets || []}
                                    selectedId={selectedSpreadsheetId}
                                    onSelect={setSelectedSpreadsheetId}
                                    isCreating={isCreating}
                                    setIsCreating={setIsCreating}
                                    newSpreadsheetName={newSpreadsheetName}
                                    setNewSpreadsheetName={setNewSpreadsheetName}
                                    onCreateSpreadsheet={handleCreateSpreadsheet}
                                    isLoading={isLoading}
                                />
                            </div>
                        </div>
                    </aside>

                    {/* Sidebar Toggle Button */}
                    <div className="relative">
                        <Button
                            variant="ghost"
                            size="icon"
                            className="absolute top-2 -left-3 z-10 h-6 w-6 rounded-full border bg-card shadow-sm"
                            onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
                            title={sidebarCollapsed ? 'Show sidebar' : 'Hide sidebar'}
                        >
                            {sidebarCollapsed ? (
                                <ChevronRight className="h-4 w-4" />
                            ) : (
                                <ChevronLeft className="h-4 w-4" />
                            )}
                        </Button>
                    </div>

                    {/* Main Content */}
                    <main className="flex-1 overflow-hidden">
                        {selectedSpreadsheetId ? (
                            <SpreadsheetGrid spreadsheetId={selectedSpreadsheetId} />
                        ) : (
                            <div className="flex h-full items-center justify-center">
                                <div className="text-center">
                                    <FileSpreadsheet className="mx-auto h-16 w-16 text-muted-foreground" />
                                    <h3 className="mt-4 text-lg font-semibold">No spreadsheet selected</h3>
                                    <p className="mt-2 text-sm text-muted-foreground">
                                        Create a new spreadsheet to get started
                                    </p>
                                </div>
                            </div>
                        )}
                    </main>
                </div>

                {/* Footer */}
                <footer className="border-t bg-card px-4 py-2 text-center text-xs text-muted-foreground">
                    © 2025. Built with <span className="text-primary">♥</span> using{' '}
                    <a
                        href="https://caffeine.ai"
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-primary hover:underline"
                    >
                        caffeine.ai
                    </a>
                </footer>
            </div>

            {/* Profile Setup Dialog */}
            <Dialog open={showProfileSetup} onOpenChange={setShowProfileSetup}>
                <DialogContent>
                    <DialogHeader>
                        <DialogTitle>Welcome!</DialogTitle>
                        <DialogDescription>
                            Please enter your name to complete your profile setup.
                        </DialogDescription>
                    </DialogHeader>
                    <div className="space-y-4 py-4">
                        <div className="space-y-2">
                            <Label htmlFor="name">Name</Label>
                            <Input
                                id="name"
                                value={profileName}
                                onChange={(e) => setProfileName(e.target.value)}
                                onKeyDown={(e) => {
                                    if (e.key === 'Enter') {
                                        handleSaveProfile();
                                    }
                                }}
                                placeholder="Enter your name"
                                autoFocus
                            />
                        </div>
                        <Button
                            onClick={handleSaveProfile}
                            disabled={saveProfile.isPending}
                            className="w-full"
                        >
                            {saveProfile.isPending ? 'Saving...' : 'Continue'}
                        </Button>
                    </div>
                </DialogContent>
            </Dialog>
        </>
    );
}

function FileList({
    spreadsheets,
    selectedId,
    onSelect,
    isCreating,
    setIsCreating,
    newSpreadsheetName,
    setNewSpreadsheetName,
    onCreateSpreadsheet,
    isLoading,
}: {
    spreadsheets: any[];
    selectedId: string | null;
    onSelect: (id: string) => void;
    isCreating: boolean;
    setIsCreating: (value: boolean) => void;
    newSpreadsheetName: string;
    setNewSpreadsheetName: (value: string) => void;
    onCreateSpreadsheet: () => void;
    isLoading: boolean;
}) {
    return (
        <div className="space-y-2">
            <Button
                variant="outline"
                size="sm"
                className="w-full justify-start"
                onClick={() => setIsCreating(true)}
            >
                <FileSpreadsheet className="mr-2 h-4 w-4" />
                New Spreadsheet
            </Button>

            {isCreating && (
                <div className="space-y-2 rounded-lg border bg-background p-3">
                    <Input
                        placeholder="Spreadsheet name"
                        value={newSpreadsheetName}
                        onChange={(e) => setNewSpreadsheetName(e.target.value)}
                        onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                                onCreateSpreadsheet();
                            } else if (e.key === 'Escape') {
                                setIsCreating(false);
                                setNewSpreadsheetName('');
                            }
                        }}
                        autoFocus
                    />
                    <div className="flex gap-2">
                        <Button size="sm" onClick={onCreateSpreadsheet} className="flex-1">
                            Create
                        </Button>
                        <Button
                            size="sm"
                            variant="outline"
                            onClick={() => {
                                setIsCreating(false);
                                setNewSpreadsheetName('');
                            }}
                            className="flex-1"
                        >
                            Cancel
                        </Button>
                    </div>
                </div>
            )}

            <Separator />

            <ScrollArea className="h-[calc(100vh-250px)]">
                {isLoading ? (
                    <div className="space-y-2">
                        {[1, 2, 3].map((i) => (
                            <div key={i} className="h-9 animate-pulse rounded-md bg-muted" />
                        ))}
                    </div>
                ) : spreadsheets.length === 0 ? (
                    <p className="py-4 text-center text-sm text-muted-foreground">
                        No spreadsheets yet
                    </p>
                ) : (
                    <div className="space-y-1">
                        {spreadsheets.map((spreadsheet) => (
                            <Button
                                key={spreadsheet.id}
                                variant={selectedId === spreadsheet.id ? 'secondary' : 'ghost'}
                                size="sm"
                                className="w-full justify-start"
                                onClick={() => onSelect(spreadsheet.id)}
                            >
                                <FileSpreadsheet className="mr-2 h-4 w-4" />
                                <span className="truncate">{spreadsheet.name}</span>
                            </Button>
                        ))}
                    </div>
                )}
            </ScrollArea>
        </div>
    );
}
