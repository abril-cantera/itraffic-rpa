/**
 * Configuration for Outlook Add-in
 */

interface Config {
  apiBaseUrl: string;
  dialogBaseUrl: string; // For Office dialogs (MUST be HTTPS)
  environment: 'development' | 'production';
  version: string;
}

// Determine environment based on hostname
const isDevelopment = window.location.hostname === 'localhost';

// Backend URLs
const LOCAL_BACKEND = 'http://localhost:7071';
const PROD_BACKEND = 'https://happy-flower-09b6bd81e.4.azurestaticapps.net';

export const config: Config = {
  // Local development uses local backend for API calls
  apiBaseUrl: isDevelopment ? LOCAL_BACKEND : PROD_BACKEND,
  // Office dialogs REQUIRE HTTPS - always use production for dialogs
  dialogBaseUrl: PROD_BACKEND,
  environment: isDevelopment ? 'development' : 'production',
  version: '1.0.0',
};

// Log configuration on load
console.log('📝 Add-in Configuration:', {
  environment: config.environment,
  apiBaseUrl: config.apiBaseUrl,
  dialogBaseUrl: config.dialogBaseUrl,
  version: config.version,
});
