/**
 * History API polyfill for iframe/restricted contexts
 * 
 * Some environments (like Outlook add-in iframes or v0 preview) may not have
 * full History API support. This polyfill provides no-op implementations
 * to prevent errors from libraries like MSAL or Next.js router.
 */

// Run immediately when this module is imported
if (typeof window !== "undefined") {
    // Check if history exists but methods are not functions (can happen in iframes)
    if (typeof window.history !== "undefined") {
        if (typeof window.history.replaceState !== "function") {
            (window.history as History).replaceState = function (
                _data: unknown,
                _unused: string,
                _url?: string | URL | null
            ): void {
                // No-op polyfill for restricted contexts
            };
        }
        if (typeof window.history.pushState !== "function") {
            (window.history as History).pushState = function (
                _data: unknown,
                _unused: string,
                _url?: string | URL | null
            ): void {
                // No-op polyfill for restricted contexts
            };
        }
    }
}

export { };
