// Ambient shim for the one Electron surface editEmailTemplate touches. The
// real `electron` package is an optional peer dependency — installing it just
// for these types would pull down its multi-hundred-MB binary, so this
// hand-written slice stands in instead. If editEmailTemplate's macOS
// BrowserWindow usage grows, extend this shim to match.
declare module 'electron' {
    interface WebContents {
        on(event: 'page-title-updated', listener: (event: { preventDefault(): void }, title: string) => void): void;
        // `any` here matches Electron's own (permissive) signature
        executeJavaScript(code: string): Promise<any>;
    }

    interface BrowserWindowOptions {
        width?: number;
        height?: number;
        title?: string;
        webPreferences?: {
            nodeIntegration?: boolean;
            contextIsolation?: boolean;
        };
    }

    class BrowserWindow {
        constructor(options?: BrowserWindowOptions);
        webContents: WebContents;
        on(event: 'closed', listener: () => void): void;
        loadURL(url: string): void;
        destroy(): void;
    }
}
