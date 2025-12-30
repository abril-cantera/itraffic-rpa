/**
 * Token Refresh Service
 * Handles automatic token renewal using MSAL silent token acquisition
 */

import { config } from '../config/config';

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

class TokenRefreshService {
  private refreshIntervalId: number | null = null;
  private tokenInfo: TokenInfo | null = null;
  private readonly REFRESH_INTERVAL = 50 * 60 * 1000; // 50 minutes in milliseconds
  private readonly TOKEN_BUFFER = 5 * 60 * 1000; // 5 minutes buffer before expiry

  /**
   * Start automatic token refresh
   */
  startAutoRefresh(initialTokenInfo: TokenInfo): void {
    console.log('🔄 [TOKEN-REFRESH] Starting automatic token refresh service');
    this.tokenInfo = initialTokenInfo;

    // Calculate when to refresh
    const expiresOn = new Date(initialTokenInfo.expiresOn);
    const now = new Date();
    const timeUntilExpiry = expiresOn.getTime() - now.getTime();
    const timeUntilRefresh = Math.max(timeUntilExpiry - this.TOKEN_BUFFER, 60000); // At least 1 minute

    console.log(`⏰ [TOKEN-REFRESH] Token expires at: ${expiresOn.toISOString()}`);
    console.log(`⏰ [TOKEN-REFRESH] Will refresh in: ${Math.round(timeUntilRefresh / 1000 / 60)} minutes`);

    // Set up periodic refresh
    this.refreshIntervalId = window.setInterval(() => {
      this.refreshToken();
    }, this.REFRESH_INTERVAL);

    // Also schedule first refresh before expiry
    setTimeout(() => {
      this.refreshToken();
    }, timeUntilRefresh);
  }

  /**
   * Stop automatic token refresh
   */
  stopAutoRefresh(): void {
    if (this.refreshIntervalId !== null) {
      console.log('🛑 [TOKEN-REFRESH] Stopping automatic token refresh');
      clearInterval(this.refreshIntervalId);
      this.refreshIntervalId = null;
    }
  }

  /**
   * Manually refresh token
   */
  async refreshToken(): Promise<void> {
    if (!this.tokenInfo) {
      console.error('❌ [TOKEN-REFRESH] No token info available');
      return;
    }

    try {
      console.log('🔄 [TOKEN-REFRESH] Refreshing token...');

      // Open a hidden dialog to perform silent token acquisition
      const newTokenInfo = await this.acquireTokenSilent();

      if (newTokenInfo) {
        console.log('✅ [TOKEN-REFRESH] Token refreshed successfully');
        console.log(`⏰ [TOKEN-REFRESH] New expiry: ${newTokenInfo.expiresOn}`);
        
        this.tokenInfo = newTokenInfo;

        // Update token in backend
        await this.updateTokenInBackend(newTokenInfo);
      }
    } catch (error: any) {
      console.error('❌ [TOKEN-REFRESH] Failed to refresh token:', error);
      
      // If silent refresh fails, user needs to re-authenticate
      if (error.message?.includes('interaction_required') || error.message?.includes('consent_required')) {
        console.warn('⚠️ [TOKEN-REFRESH] User interaction required - stopping auto-refresh');
        this.stopAutoRefresh();
        this.notifyUserReauthNeeded();
      }
    }
  }

  /**
   * Acquire token silently using a hidden dialog
   */
  private async acquireTokenSilent(): Promise<TokenInfo | null> {
    return new Promise((resolve, reject) => {
      // Create URL for silent token acquisition
      // Office dialogs REQUIRE HTTPS - use dialogBaseUrl (always production)
      const dialogUrl = `${config.dialogBaseUrl}/login.html?silent=true&loginHint=${encodeURIComponent(this.tokenInfo!.account.username)}`;

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 1, width: 1, displayInIframe: false }, // Tiny hidden dialog
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error('Failed to open silent auth dialog'));
            return;
          }

          const dialog = result.value;
          const timeout = setTimeout(() => {
            dialog.close();
            reject(new Error('Silent auth timeout'));
          }, 30000); // 30 second timeout

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
            clearTimeout(timeout);
            dialog.close();

            try {
              const message = JSON.parse(arg.message);

              if (message.status === 'success' && message.token) {
                resolve({
                  accessToken: message.token,
                  expiresOn: message.expiresOn,
                  account: {
                    homeAccountId: message.userId,
                    environment: 'login.microsoftonline.com',
                    tenantId: 'common',
                    username: message.email,
                    localAccountId: message.userId,
                    name: message.name
                  }
                });
              } else {
                reject(new Error(message.error || 'Silent auth failed'));
              }
            } catch (error) {
              reject(new Error('Invalid silent auth response'));
            }
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
            clearTimeout(timeout);
            if (arg.error === 12006) {
              reject(new Error('Dialog closed by user'));
            } else {
              reject(new Error('Silent auth dialog error'));
            }
          });
        }
      );
    });
  }

  /**
   * Update token in backend
   */
  private async updateTokenInBackend(tokenInfo: TokenInfo): Promise<void> {
    try {
      console.log('📡 [TOKEN-REFRESH] Updating token in backend...');

      // Use dialogBaseUrl for token refresh since user data is stored in production
      const response = await fetch(`${config.dialogBaseUrl}/auth/refresh-token`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          userId: tokenInfo.account.localAccountId,
          email: tokenInfo.account.username,
          accessToken: tokenInfo.accessToken,
          expiresOn: tokenInfo.expiresOn
        })
      });

      if (!response.ok) {
        throw new Error(`Backend update failed: ${response.statusText}`);
      }

      console.log('✅ [TOKEN-REFRESH] Token updated in backend');
    } catch (error: any) {
      console.error('❌ [TOKEN-REFRESH] Failed to update token in backend:', error);
      // Don't throw - this is not critical, token is still valid
    }
  }

  /**
   * Notify user that re-authentication is needed
   */
  private notifyUserReauthNeeded(): void {
    // Show a non-intrusive notification
    const notification = document.createElement('div');
    notification.className = 'token-refresh-notification';
    notification.innerHTML = `
      <div style="background: #fff3cd; border: 1px solid #ffc107; padding: 12px; margin: 10px; border-radius: 4px;">
        <strong>⚠️ Re-autenticación necesaria</strong>
        <p>Tu sesión expirará pronto. Por favor, cierra sesión y vuelve a iniciar sesión.</p>
      </div>
    `;
    
    const container = document.getElementById('content-main');
    if (container) {
      container.prepend(notification);
    }
  }
}

// Export singleton instance
export const tokenRefreshService = new TokenRefreshService();
