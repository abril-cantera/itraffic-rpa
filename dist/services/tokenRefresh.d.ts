/**
 * Token Refresh Service
 * Handles automatic token renewal using MSAL silent token acquisition
 */
interface TokenInfo {
    accessToken: string;
    expiresOn: string;
    account: {
        homeAccountId: string;
        environment: string;
        tenantId: string;
        username: string;
        localAccountId: string;
        name: string;
    };
}
declare class TokenRefreshService {
    private refreshIntervalId;
    private tokenInfo;
    private readonly REFRESH_INTERVAL;
    private readonly TOKEN_BUFFER;
    /**
     * Start automatic token refresh
     */
    startAutoRefresh(initialTokenInfo: TokenInfo): void;
    /**
     * Stop automatic token refresh
     */
    stopAutoRefresh(): void;
    /**
     * Manually refresh token
     */
    refreshToken(): Promise<void>;
    /**
     * Acquire token silently using a hidden dialog
     */
    private acquireTokenSilent;
    /**
     * Update token in backend
     */
    private updateTokenInBackend;
    /**
     * Notify user that re-authentication is needed
     */
    private notifyUserReauthNeeded;
}
export declare const tokenRefreshService: TokenRefreshService;
export {};
//# sourceMappingURL=tokenRefresh.d.ts.map