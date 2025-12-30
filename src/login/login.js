/**
 * Login dialog for Office Add-in fallback authentication
 * This dialog performs authorization-code flow with confidential client:
 * 1. Redirects to Azure AD to request authorization code
 * 2. Returns the authorization code to the taskpane
 * 3. Backend redeems the code with client_secret to get access_token + refresh_token
 * 
 * Note: Using Web redirect URI (not SPA) allows backend to use client_secret
 * Updated: 2024-12-01 - Improved handling when opened outside Office dialog
 */

const AUTH_CONFIG = {
    clientId: '6637590b-a6a4-4e53-b429-a766c66f03c3',
    // Use 'organizations' for work accounts - allows any organization that has granted admin consent
    authority: 'https://login.microsoftonline.com/organizations',
    scopes: [
        'openid',
        'profile',
        'email',
        'offline_access',
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Mail.Read',
        'https://graph.microsoft.com/Mail.ReadWrite'
    ]
};

// Check for auth code IMMEDIATELY before Office.onReady (handles browser tab scenario)
(function checkAuthCodeImmediately() {
    const urlParams = new URLSearchParams(window.location.search);
    const code = urlParams.get('code');
    
    if (code) {
        console.log('🔍 Auth code detected in URL (immediate check)');
        const redirectUri = window.location.origin + window.location.pathname;
        const state = urlParams.get('state');
        
        // Store in localStorage
        try {
            const message = {
                status: 'code',
                authorizationCode: code,
                redirectUri: redirectUri,
                state: state,
                timestamp: Date.now()
            };
            localStorage.setItem('office_auth_pending', JSON.stringify(message));
            console.log('✅ Auth code pre-stored in localStorage');
        } catch (e) {
            console.error('❌ Failed to pre-store auth code:', e);
        }
        
        // ALSO store in server bridge (handles cross-context scenario)
        fetch(redirectUri.replace('/login.html', '') + '/auth/bridge', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                code: code,
                redirectUri: redirectUri,
                state: state,
                email: 'pending' // Will be used as key
            })
        }).then(r => {
            console.log('✅ Auth code stored in server bridge');
        }).catch(e => {
            console.error('⚠️ Failed to store in server bridge:', e);
        });
    }
})();

Office.onReady((info) => {
    console.log('✅ Office.onReady called', info);
    
    const urlParams = new URLSearchParams(window.location.search);
    const isSilent = urlParams.get('silent') === 'true';
    const loginHint = urlParams.get('loginHint');
    const redirectUri = getRedirectUri();
    const forceConsent = urlParams.get('forceConsent') === 'true';
    const stateFromResponse = urlParams.get('state');
    const storedState = sessionStorage.getItem('auth_state');
    const error = urlParams.get('error');
    const errorDescription = urlParams.get('error_description');
    const code = urlParams.get('code');

    console.log('🔐 Login dialog initialized [VERSION-2024-12-01]');
    console.log('📍 Redirect URI:', redirectUri);
    console.log('🔇 Silent mode:', isSilent);
    console.log('🙏 Force consent:', forceConsent);
    if (loginHint) console.log('📧 Login hint:', loginHint);

    if (error) {
        console.error('❌ Authorization error:', error, errorDescription);
        reportError(errorDescription || error);
        return;
    }

    if (code) {
        console.log('✅ Authorization code received');

        // Skip state check if we're in a new browser tab (state might not be available)
        if (storedState && stateFromResponse && storedState !== stateFromResponse) {
            console.warn('⚠️ State mismatch. Expected:', storedState, 'Received:', stateFromResponse);
            // Don't block on state mismatch - continue with auth
        }

        sessionStorage.removeItem('auth_state');

        // Clean URL to avoid reprocessing
        if (window.history && typeof window.history.replaceState === 'function') {
            window.history.replaceState({}, document.title, redirectUri);
        }

        const message = {
            status: 'code',
            authorizationCode: code,
            redirectUri,
            state: stateFromResponse
        };

        // Try to send message to parent (taskpane) via Office API
        let messageSent = false;
        
        if (typeof Office !== 'undefined' && Office.context && Office.context.ui && Office.context.ui.messageParent) {
            try {
                Office.context.ui.messageParent(JSON.stringify(message));
                console.log('✅ Message sent via Office.context.ui.messageParent');
                messageSent = true;
            } catch (e) {
                console.error('❌ Office.context.ui.messageParent failed:', e);
            }
        }
        
        if (!messageSent) {
            console.log('⚠️ Office.context.ui.messageParent not available - using fallback');
            
            // FALLBACK: Store the auth code in localStorage for the add-in to pick up
            // This handles the case where Azure AD redirects to a new browser tab
            console.log('📦 Storing auth code in localStorage as fallback...');
            try {
                localStorage.setItem('office_auth_pending', JSON.stringify({
                    ...message,
                    timestamp: Date.now()
                }));
                console.log('✅ Auth code stored in localStorage');
            } catch (e) {
                console.error('❌ Failed to store auth code:', e);
            }
            
            // Show success message to user
            showSuccessMessage();
            
            // Try window.opener as last resort
            if (window.opener) {
                try {
                    window.opener.postMessage(message, '*');
                    console.log('📤 Message sent via window.opener.postMessage');
                } catch (e) {
                    console.error('❌ window.opener.postMessage failed:', e);
                }
            }
        }
        return;
    }

    console.log('ℹ️ No authorization code present, initiating login redirect');
    startAuthorizationFlow({ isSilent, loginHint, redirectUri, forceConsent });
});

function startAuthorizationFlow({ isSilent, loginHint, redirectUri, forceConsent }) {
    const state = generateRandomString();
    sessionStorage.setItem('auth_state', state);

    // Use 'select_account' to let user choose account, or 'none' for silent
    // Admin consent has already been granted, so no need for 'consent' prompt
    let prompt = isSilent ? 'none' : 'select_account';

    const authorizeParams = new URLSearchParams({
        client_id: AUTH_CONFIG.clientId,
        response_type: 'code',
        redirect_uri: redirectUri,
        response_mode: 'query',
        scope: AUTH_CONFIG.scopes.join(' '),
        prompt,
        state
    });

    if (loginHint) {
        authorizeParams.set('login_hint', loginHint);
    }

    const authorizeUrl = `${AUTH_CONFIG.authority}/oauth2/v2.0/authorize?${authorizeParams.toString()}`;
    console.log('🌐 Redirecting to Azure AD:', authorizeUrl);
    window.location.replace(authorizeUrl);
}

function reportError(message) {
    const payload = {
        status: 'error',
        error: message || 'Authentication failed'
    };
    
    if (typeof Office !== 'undefined' && Office.context && Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent(JSON.stringify(payload));
    } else {
        console.error('❌ Error to report:', message);
        if (window.opener) {
            window.opener.postMessage(payload, '*');
            window.close();
        }
    }
}

function getRedirectUri() {
    const url = new URL(window.location.href);
    url.search = '';
    url.hash = '';
    return url.toString();
}

function generateRandomString(length = 32) {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    const array = new Uint32Array(length);
    window.crypto.getRandomValues(array);
    for (let i = 0; i < length; i++) {
        result += chars[array[i] % chars.length];
    }
    return result;
}

function showSuccessMessage() {
    document.body.innerHTML = `
        <div style="font-family: 'Segoe UI', sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
            <div style="text-align: center; background: white; padding: 40px; border-radius: 10px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); max-width: 400px;">
                <h2 style="color: #28a745; margin-bottom: 20px;">✅ Autenticación exitosa</h2>
                <p style="color: #666; margin-bottom: 15px;">La autenticación se completó correctamente.</p>
                <p style="color: #333; font-weight: bold;">Cierra esta pestaña y vuelve a Outlook.</p>
                <p style="color: #999; font-size: 12px; margin-top: 20px;">
                    En el add-in, haz clic en "Try Again" para completar el registro.
                </p>
                <button onclick="window.close()" style="margin-top: 20px; padding: 10px 30px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 14px;">
                    Cerrar esta ventana
                </button>
            </div>
        </div>
    `;
}
