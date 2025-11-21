import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { ThemeProvider } from 'next-themes';
import { Toaster } from '@/components/ui/sonner';
import SpreadsheetApp from './components/SpreadsheetApp';

const queryClient = new QueryClient({
    defaultOptions: {
        queries: {
            refetchOnWindowFocus: false,
            retry: 1,
        },
    },
});

export default function App() {
    return (
        <ThemeProvider attribute="class" defaultTheme="light" enableSystem>
            <QueryClientProvider client={queryClient}>
                <SpreadsheetApp />
                <Toaster />
            </QueryClientProvider>
        </ThemeProvider>
    );
}
