import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { useActor } from './useActor';
import { useInternetIdentity } from './useInternetIdentity';
import type { SpreadsheetFile, Sheet, Cell, UserProfile, SpreadsheetPermission, CellFormat, ImageData } from '../backend';
import { Principal } from '@icp-sdk/core/principal';

// User Profile Hooks
export function useGetCallerUserProfile() {
    const { actor, isFetching: actorFetching } = useActor();

    const query = useQuery<UserProfile | null>({
        queryKey: ['currentUserProfile'],
        queryFn: async () => {
            if (!actor) throw new Error('Actor not available');
            return actor.getCallerUserProfile();
        },
        enabled: !!actor && !actorFetching,
        retry: false,
    });

    return {
        ...query,
        isLoading: actorFetching || query.isLoading,
        isFetched: !!actor && query.isFetched,
    };
}

export function useSaveCallerUserProfile() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async (profile: UserProfile) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.saveCallerUserProfile(profile);
        },
        onSuccess: () => {
            queryClient.invalidateQueries({ queryKey: ['currentUserProfile'] });
        },
    });
}

// Spreadsheet Hooks
export function useListSpreadsheets() {
    const { actor, isFetching } = useActor();

    return useQuery<SpreadsheetFile[]>({
        queryKey: ['spreadsheets'],
        queryFn: async () => {
            if (!actor) return [];
            return actor.listSpreadsheets();
        },
        enabled: !!actor && !isFetching,
    });
}

export function useGetSpreadsheet(id: string | null) {
    const { actor, isFetching } = useActor();

    return useQuery<SpreadsheetFile | null>({
        queryKey: ['spreadsheet', id],
        queryFn: async () => {
            if (!actor || !id) return null;
            return actor.getSpreadsheet(id);
        },
        enabled: !!actor && !isFetching && !!id,
    });
}

export function useCreateSpreadsheet() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async (name: string) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.createSpreadsheet(name);
        },
        onSuccess: () => {
            queryClient.invalidateQueries({ queryKey: ['spreadsheets'] });
        },
    });
}

export function useDeleteSpreadsheet() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async (spreadsheetId: string) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.deleteSpreadsheet(spreadsheetId);
        },
        onSuccess: () => {
            queryClient.invalidateQueries({ queryKey: ['spreadsheets'] });
        },
    });
}

// Sheet Hooks
export function useGetSheet(spreadsheetId: string | null, sheetName: string | null) {
    const { actor, isFetching } = useActor();

    return useQuery<Sheet | null>({
        queryKey: ['sheet', spreadsheetId, sheetName],
        queryFn: async () => {
            if (!actor || !spreadsheetId || !sheetName) return null;
            return actor.getSheet(spreadsheetId, sheetName);
        },
        enabled: !!actor && !isFetching && !!spreadsheetId && !!sheetName,
    });
}

export function useListSheets(spreadsheetId: string | null) {
    const { actor, isFetching } = useActor();

    return useQuery<string[]>({
        queryKey: ['sheets', spreadsheetId],
        queryFn: async () => {
            if (!actor || !spreadsheetId) return [];
            return actor.listSheets(spreadsheetId);
        },
        enabled: !!actor && !isFetching && !!spreadsheetId,
    });
}

export function useAddSheet() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
        }: {
            spreadsheetId: string;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.addSheet(spreadsheetId);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['spreadsheet', variables.spreadsheetId],
            });
            queryClient.invalidateQueries({
                queryKey: ['sheets', variables.spreadsheetId],
            });
        },
    });
}

export function useDeleteSheet() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
        }: {
            spreadsheetId: string;
            sheetName: string;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.deleteSheet(spreadsheetId, sheetName);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['spreadsheet', variables.spreadsheetId],
            });
            queryClient.invalidateQueries({
                queryKey: ['sheets', variables.spreadsheetId],
            });
        },
    });
}

export function useSwitchSheet() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
        }: {
            spreadsheetId: string;
            sheetName: string;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.switchSheet(spreadsheetId, sheetName);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['spreadsheet', variables.spreadsheetId],
            });
        },
    });
}

export function useGetActiveSheet(spreadsheetId: string | null) {
    const { actor, isFetching } = useActor();

    return useQuery<string | null>({
        queryKey: ['activeSheet', spreadsheetId],
        queryFn: async () => {
            if (!actor || !spreadsheetId) return null;
            return actor.getActiveSheet(spreadsheetId);
        },
        enabled: !!actor && !isFetching && !!spreadsheetId,
    });
}

// Cell Hooks
export function useSaveCell() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
            cellId,
            value,
            formula,
        }: {
            spreadsheetId: string;
            sheetName: string;
            cellId: string;
            value: string;
            formula: string | null;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.saveCell(spreadsheetId, sheetName, cellId, value, formula);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['sheet', variables.spreadsheetId, variables.sheetName],
            });
        },
    });
}

export function useGetRange() {
    const { actor } = useActor();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
            cellIds,
        }: {
            spreadsheetId: string;
            sheetName: string;
            cellIds: string[];
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.getRange(spreadsheetId, sheetName, cellIds);
        },
    });
}

// Formatting Hooks
export function useApplyFormatToSelection() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
            cellIds,
            formatObj,
        }: {
            spreadsheetId: string;
            sheetName: string;
            cellIds: string[];
            formatObj: CellFormat;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.applyFormatToSelection(spreadsheetId, sheetName, cellIds, formatObj);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['sheet', variables.spreadsheetId, variables.sheetName],
            });
        },
    });
}

export function useApplyFontColor() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
            cellIds,
            color,
        }: {
            spreadsheetId: string;
            sheetName: string;
            cellIds: string[];
            color: string;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.applyFontColor(spreadsheetId, sheetName, cellIds, color);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['sheet', variables.spreadsheetId, variables.sheetName],
            });
        },
    });
}

export function useApplyFillColor() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
            cellIds,
            color,
        }: {
            spreadsheetId: string;
            sheetName: string;
            cellIds: string[];
            color: string;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.applyFillColor(spreadsheetId, sheetName, cellIds, color);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['sheet', variables.spreadsheetId, variables.sheetName],
            });
        },
    });
}

// Image Hooks
export function useAddImage() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
            image,
        }: {
            spreadsheetId: string;
            sheetName: string;
            image: ImageData;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.addImage(spreadsheetId, sheetName, image);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['sheet', variables.spreadsheetId, variables.sheetName],
            });
        },
    });
}

export function useUpdateImage() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
            image,
        }: {
            spreadsheetId: string;
            sheetName: string;
            image: ImageData;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.updateImage(spreadsheetId, sheetName, image);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['sheet', variables.spreadsheetId, variables.sheetName],
            });
        },
    });
}

export function useDeleteImage() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            sheetName,
            imageId,
        }: {
            spreadsheetId: string;
            sheetName: string;
            imageId: string;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            return actor.deleteImage(spreadsheetId, sheetName, imageId);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['sheet', variables.spreadsheetId, variables.sheetName],
            });
        },
    });
}

// Sharing Hooks
export function useShareSpreadsheet() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            user,
            permission,
        }: {
            spreadsheetId: string;
            user: string;
            permission: SpreadsheetPermission;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            const principal = Principal.fromText(user);
            return actor.shareSpreadsheet(spreadsheetId, principal, permission);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['spreadsheet', variables.spreadsheetId],
            });
        },
    });
}

export function useRevokeSpreadsheetAccess() {
    const { actor } = useActor();
    const queryClient = useQueryClient();

    return useMutation({
        mutationFn: async ({
            spreadsheetId,
            user,
        }: {
            spreadsheetId: string;
            user: string;
        }) => {
            if (!actor) throw new Error('Actor not initialized');
            const principal = Principal.fromText(user);
            return actor.revokeSpreadsheetAccess(spreadsheetId, principal);
        },
        onSuccess: (_, variables) => {
            queryClient.invalidateQueries({
                queryKey: ['spreadsheet', variables.spreadsheetId],
            });
        },
    });
}
