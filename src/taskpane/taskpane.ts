/**
 * Email Classifier - Outlook Add-in
 * Main taskpane logic with SSO authentication
 */

import { config } from '../config/config';
import { rpaService } from '../services/rpaService';
import type { ReservationData, ExtractionApiResponse, ExtractionApiError, ExtractionState } from '../types/reservation';

// Application state
interface AppState {
  isRegistered: boolean;
  userId: string | null;
  email: string | null;
  isLoading: boolean;
  error: string | null;
  isAuthInProgress: boolean;
  authRetryCount: number;
  lastAuthAttempt: number | null;
}

const state: AppState = {
  isRegistered: false,
  userId: null,
  email: null,
  isLoading: false,
  error: null,
  isAuthInProgress: false,
  authRetryCount: 0,
  lastAuthAttempt: null,
};

// Global dialog reference to prevent multiple dialogs
let activeDialog: Office.Dialog | null = null;

/**
 * Close any active dialog before opening a new one
 */
function closeActiveDialog(): void {
  if (activeDialog) {
    try {
      console.log('Closing existing dialog...');
      activeDialog.close();
    } catch (e) {
      console.log('Dialog already closed or invalid');
    }
    activeDialog = null;
  }
}

// Extraction state
const extractionState: ExtractionState = {
  isExtracting: false,
  hasExtracted: false,
  data: null,
  error: null,
  isEditing: false,
};

// Master data state
interface MasterData {
  sellers: Array<{ code: string; name: string }>;
  clients: Array<{ code: string; name: string; displayName?: string; cuit?: string }>;
  contacts: Array<{ code: string; name: string; displayName?: string; email?: string; phone?: string }>;
  currencies: Array<{ code: string; name: string }>;
  statuses: Array<{ code: string; name: string }>;
  reservationTypes: Array<{ code: string; name: string }>;
  genders: Array<{ code: string; name: string }>;
  documentTypes: Array<{ code: string; name: string }>;
  countries: Array<{ code: string; name: string; fiscalCode?: string }>;
  loaded: boolean;
}

const masterData: MasterData = {
  sellers: [],
  clients: [],
  contacts: [],
  currencies: [],
  statuses: [],
  reservationTypes: [],
  genders: [],
  documentTypes: [],
  countries: [],
  loaded: false,
};

/**
 * Office.js initialization
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('📧 Email Classifier Add-in loaded');
    console.log('🔄 VERSION: 2024-11-06-15:15 - SSO DISABLED - DIALOG ONLY');
    console.log('🚀 Build timestamp:', new Date().toISOString());
    initializeApp();
  }
});

/**
 * Initialize the application
 */
async function initializeApp(): Promise<void> {
  // Keep loading visible while checking registration
  // Don't hide it yet - let checkRegistrationStatus do it

  // Setup event listeners
  document.getElementById('register-button')!.addEventListener('click', handleRegister);
  document.getElementById('retry-button')!.addEventListener('click', handleRegister);
  document.getElementById('logout-button')!.addEventListener('click', handleLogout);
  document.getElementById('help-link')!.addEventListener('click', showHelp);

  // Extraction event listeners
  document.getElementById('extract-button')!.addEventListener('click', handleExtractReservation);
  document.getElementById('reanalyze-button')!.addEventListener('click', handleReanalyze);
  document.getElementById('confirm-button')!.addEventListener('click', handleConfirmReservation);
  document.getElementById('retry-extraction-button')!.addEventListener('click', handleExtractReservation);

  // Header toggle
  document.getElementById('toggle-header-btn')!.addEventListener('click', toggleHeader);

  // Load master data for dropdowns (in parallel with registration check)
  loadMasterData();

  // Test RPA connection
  testRPAConnection();

  // Check if user is already registered (this will hide loading when done)
  await checkRegistrationStatus();
}

/**
 * Load master data from API (sellers, clients, currencies, statuses, reservationTypes)
 */
async function loadMasterData(): Promise<void> {
  console.log('📋 Loading master data...');

  // Show loaders
  showFieldLoader('seller', true);
  showFieldLoader('client', true);
  showFieldLoader('contact', true);
  showFieldLoader('currency', true);
  showFieldLoader('status', true);
  showFieldLoader('reservation-type', true);

  try {
    const response = await fetch(`${config.apiBaseUrl}/api/master-data`, {
      method: 'GET',
      headers: { 'Content-Type': 'application/json' }
    });

    if (!response.ok) {
      console.warn('⚠️ Could not load master data, using defaults');
      hideAllFieldLoaders();
      enableMasterDataFields();
      return;
    }

    const result = await response.json();

    if (result.success && result.data) {
      masterData.sellers = result.data.sellers || [];
      masterData.clients = result.data.clients || [];
      masterData.contacts = result.data.contacts || [];
      masterData.currencies = result.data.currencies || [];
      masterData.statuses = result.data.statuses || [];
      masterData.reservationTypes = result.data.reservationTypes || [];
      masterData.genders = result.data.genders || [];
      masterData.documentTypes = result.data.documentTypes || [];
      masterData.countries = result.data.countries || [];
      masterData.loaded = true;

      console.log('✅ Master data loaded:', {
        sellers: masterData.sellers.length,
        clients: masterData.clients.length,
        contacts: masterData.contacts.length,
        currencies: masterData.currencies.length,
        statuses: masterData.statuses.length,
        reservationTypes: masterData.reservationTypes.length,
        genders: masterData.genders.length,
        documentTypes: masterData.documentTypes.length,
        countries: masterData.countries.length
      });

      // Populate dropdowns
      populateSellerDropdown();
      populateClientDatalist();
      populateContactDropdown();
      populateCurrencyDropdown();
      populateStatusDropdown();
      populateReservationTypeDropdown();
    }
  } catch (error) {
    console.error('❌ Error loading master data:', error);
    hideAllFieldLoaders();
    enableMasterDataFields();
  }
}

/**
 * Show/hide field loader
 */
function showFieldLoader(field: 'seller' | 'client' | 'contact' | 'currency' | 'status' | 'reservation-type', show: boolean): void {
  const loader = document.getElementById(`${field}-loader`);
  if (loader) {
    if (show) {
      loader.classList.add('active');
    } else {
      loader.classList.remove('active');
    }
  }
}

/**
 * Hide all field loaders
 */
function hideAllFieldLoaders(): void {
  showFieldLoader('seller', false);
  showFieldLoader('client', false);
  showFieldLoader('contact', false);
  showFieldLoader('currency', false);
  showFieldLoader('status', false);
  showFieldLoader('reservation-type', false);
}

/**
 * Enable master data fields after loading
 */
function enableMasterDataFields(): void {
  const sellerInput = document.getElementById('input-seller') as HTMLInputElement;
  const clientInput = document.getElementById('input-client') as HTMLInputElement;
  const contactInput = document.getElementById('input-contact') as HTMLInputElement;
  const currencyInput = document.getElementById('input-currency') as HTMLInputElement;
  const statusInput = document.getElementById('input-status') as HTMLInputElement;
  const reservationTypeInput = document.getElementById('input-reservation-type') as HTMLInputElement;

  if (sellerInput) {
    sellerInput.disabled = false;
    sellerInput.placeholder = 'Buscar vendedor...';
  }
  if (clientInput) {
    clientInput.disabled = false;
    clientInput.placeholder = 'Buscar cliente...';
  }
  if (contactInput) {
    contactInput.disabled = false;
    contactInput.placeholder = 'Buscar contacto...';
  }
  if (currencyInput) {
    currencyInput.disabled = false;
    currencyInput.placeholder = 'Buscar moneda...';
  }
  if (statusInput) {
    statusInput.disabled = false;
    statusInput.placeholder = 'Buscar estado...';
  }
  if (reservationTypeInput) {
    reservationTypeInput.disabled = false;
    reservationTypeInput.placeholder = 'Buscar tipo...';
  }
}

/**
 * Populate seller dropdown with master data (searchable)
 */
function populateSellerDropdown(): void {
  const input = document.getElementById('input-seller') as HTMLInputElement;
  const dropdown = document.getElementById('seller-dropdown') as HTMLDivElement;
  const hiddenInput = document.getElementById('input-seller-value') as HTMLInputElement;

  if (!input || !dropdown) return;

  // Enable input
  input.disabled = false;
  input.placeholder = 'Buscar vendedor...';
  showFieldLoader('seller', false);

  // Setup searchable dropdown
  setupSearchableDropdown(input, dropdown, hiddenInput, masterData.sellers, 'name');

  console.log(`📋 Seller searchable dropdown ready with ${masterData.sellers.length} options`);
}

/**
 * Populate client dropdown with master data (searchable with lazy loading)
 * Shows full format like iTraffic: "CODE - NAME - Cuit:XXX"
 */
function populateClientDatalist(): void {
  const input = document.getElementById('input-client') as HTMLInputElement;
  const dropdown = document.getElementById('client-dropdown') as HTMLDivElement;
  const hiddenInput = document.getElementById('input-client-value') as HTMLInputElement;

  if (!input || !dropdown) return;

  // Enable input
  input.disabled = false;
  input.placeholder = 'Buscar cliente...';
  showFieldLoader('client', false);

  // Setup searchable dropdown with lazy loading - shows displayName if available
  setupSearchableDropdownForClients(input, dropdown, hiddenInput, masterData.clients);

  console.log(`📋 Client searchable dropdown ready with ${masterData.clients.length} options (lazy loading)`);
}

/**
 * Populate contact dropdown with master data (separate contacts list)
 * Shows format: "CODE - NAME"
 * Field starts empty - user must select a contact
 */
function populateContactDropdown(): void {
  const input = document.getElementById('input-contact') as HTMLInputElement;
  const dropdown = document.getElementById('contact-dropdown') as HTMLDivElement;
  const hiddenInput = document.getElementById('input-contact-value') as HTMLInputElement;

  if (!input || !dropdown) return;

  // Enable input
  input.disabled = false;
  input.placeholder = 'Buscar contacto...';
  showFieldLoader('contact', false);

  // Setup searchable dropdown with lazy loading - uses contacts list
  setupSearchableDropdownForClients(input, dropdown, hiddenInput, masterData.contacts);

  // Leave empty by default - user must select
  input.value = '';
  if (hiddenInput) hiddenInput.value = '';

  console.log(`📋 Contact searchable dropdown ready with ${masterData.contacts.length} options (lazy loading)`);
}

/**
 * Populate currency dropdown with master data (searchable with aliases)
 */
function populateCurrencyDropdown(): void {
  const input = document.getElementById('input-currency') as HTMLInputElement;
  const dropdown = document.getElementById('currency-dropdown') as HTMLDivElement;
  const hiddenInput = document.getElementById('input-currency-value') as HTMLInputElement;

  if (!input || !dropdown) return;

  // Enable input
  input.disabled = false;
  input.placeholder = 'Buscar moneda...';
  showFieldLoader('currency', false);

  // Currency aliases for common abbreviations
  const currencyAliases: Record<string, string[]> = {
    'DOLARES': ['USD', 'DOLAR', 'US$', 'U$S', 'DOLLAR'],
    'PESOS': ['ARS', 'PESO', '$', 'AR$', 'PESOS ARGENTINOS'],
    'EUROS': ['EUR', 'EURO', '€'],
    'REALES': ['BRL', 'REAL', 'R$', 'REAIS'],
  };

  // Setup searchable dropdown with alias support
  setupSearchableDropdownWithAliases(input, dropdown, hiddenInput, masterData.currencies, currencyAliases);

  // Set default value to DOLARES if exists
  const dolares = masterData.currencies.find(c => c.name.toUpperCase().includes('DOLAR'));
  if (dolares) {
    input.value = dolares.name;
    if (hiddenInput) hiddenInput.value = dolares.name;
  } else if (masterData.currencies.length > 0) {
    input.value = masterData.currencies[0].name;
    if (hiddenInput) hiddenInput.value = masterData.currencies[0].name;
  }

  console.log(`📋 Currency searchable dropdown ready with ${masterData.currencies.length} options`);
}

/**
 * Populate status dropdown with master data (searchable with lazy loading)
 * Shows only name (no code)
 */
function populateStatusDropdown(): void {
  const input = document.getElementById('input-status') as HTMLInputElement;
  const dropdown = document.getElementById('status-dropdown') as HTMLDivElement;
  const hiddenInput = document.getElementById('input-status-value') as HTMLInputElement;

  if (!input || !dropdown) return;

  // Enable input
  input.disabled = false;
  input.placeholder = 'Buscar estado...';
  showFieldLoader('status', false);

  // Setup searchable dropdown with lazy loading - name only display
  setupSearchableDropdownLazyNameOnly(input, dropdown, hiddenInput, masterData.statuses);

  // Set default value to CONFIRMACION if exists
  const confirmacion = masterData.statuses.find(s => s.code === 'FI');
  if (confirmacion) {
    input.value = confirmacion.name;
    if (hiddenInput) hiddenInput.value = confirmacion.code;
  }

  console.log(`📋 Status searchable dropdown ready with ${masterData.statuses.length} options`);
}

/**
 * Populate reservation type dropdown with master data (searchable with lazy loading)
 * Shows only name (no code)
 */
function populateReservationTypeDropdown(): void {
  const input = document.getElementById('input-reservation-type') as HTMLInputElement;
  const dropdown = document.getElementById('reservation-type-dropdown') as HTMLDivElement;
  const hiddenInput = document.getElementById('input-reservation-type-value') as HTMLInputElement;

  if (!input || !dropdown) return;

  // Enable input
  input.disabled = false;
  input.placeholder = 'Buscar tipo...';
  showFieldLoader('reservation-type', false);

  // Setup searchable dropdown with lazy loading - name only display
  setupSearchableDropdownLazyNameOnly(input, dropdown, hiddenInput, masterData.reservationTypes);

  console.log(`📋 Reservation type searchable dropdown ready with ${masterData.reservationTypes.length} options`);
}

/**
 * Setup a searchable dropdown component
 */
function setupSearchableDropdown(
  input: HTMLInputElement,
  dropdown: HTMLDivElement,
  hiddenInput: HTMLInputElement | null,
  items: Array<{ code: string; name: string }>,
  displayField: 'name' | 'code'
): void {
  let highlightedIndex = -1;
  let filteredItems = [...items];

  // Render dropdown items
  function renderDropdown(filter: string = ''): void {
    const searchTerm = filter.toLowerCase();
    filteredItems = items.filter(item =>
      item.name.toLowerCase().includes(searchTerm) ||
      item.code.toLowerCase().includes(searchTerm)
    );

    dropdown.innerHTML = '';

    if (filteredItems.length === 0) {
      dropdown.innerHTML = '<div class="dropdown-empty">No se encontraron resultados</div>';
      return;
    }

    // Limit to 50 items for performance
    const displayItems = filteredItems.slice(0, 50);

    displayItems.forEach((item, index) => {
      const div = document.createElement('div');
      div.className = 'dropdown-item';
      div.dataset.value = item[displayField];
      div.dataset.index = index.toString();
      div.textContent = item.name;

      div.addEventListener('click', () => {
        selectItem(item);
      });

      dropdown.appendChild(div);
    });

    if (filteredItems.length > 50) {
      const moreDiv = document.createElement('div');
      moreDiv.className = 'dropdown-empty';
      moreDiv.textContent = `... y ${filteredItems.length - 50} más. Escribe para filtrar.`;
      dropdown.appendChild(moreDiv);
    }
  }

  // Select an item
  function selectItem(item: { code: string; name: string }): void {
    input.value = item.name;
    if (hiddenInput) hiddenInput.value = item[displayField];
    dropdown.classList.remove('active');
    highlightedIndex = -1;
  }

  // Show dropdown
  function showDropdown(): void {
    renderDropdown(input.value);
    dropdown.classList.add('active');
  }

  // Hide dropdown
  function hideDropdown(): void {
    dropdown.classList.remove('active');
    highlightedIndex = -1;
  }

  // Update highlight
  function updateHighlight(): void {
    const items = dropdown.querySelectorAll('.dropdown-item');
    items.forEach((item, index) => {
      item.classList.toggle('highlighted', index === highlightedIndex);
    });

    // Scroll into view
    if (highlightedIndex >= 0 && items[highlightedIndex]) {
      items[highlightedIndex].scrollIntoView({ block: 'nearest' });
    }
  }

  // Event: Focus - show dropdown
  input.addEventListener('focus', () => {
    showDropdown();
  });

  // Event: Input - filter items
  input.addEventListener('input', () => {
    renderDropdown(input.value);
    dropdown.classList.add('active');
    highlightedIndex = -1;
  });

  // Event: Keydown - navigation
  input.addEventListener('keydown', (e) => {
    const items = dropdown.querySelectorAll('.dropdown-item');

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        if (!dropdown.classList.contains('active')) {
          showDropdown();
        }
        highlightedIndex = Math.min(highlightedIndex + 1, items.length - 1);
        updateHighlight();
        break;

      case 'ArrowUp':
        e.preventDefault();
        highlightedIndex = Math.max(highlightedIndex - 1, 0);
        updateHighlight();
        break;

      case 'Enter':
        e.preventDefault();
        if (highlightedIndex >= 0 && filteredItems[highlightedIndex]) {
          selectItem(filteredItems[highlightedIndex]);
        }
        break;

      case 'Escape':
        hideDropdown();
        break;

      case 'Tab':
        hideDropdown();
        break;
    }
  });

  // Event: Click outside - hide dropdown
  document.addEventListener('click', (e) => {
    if (!input.contains(e.target as Node) && !dropdown.contains(e.target as Node)) {
      hideDropdown();
    }
  });
}

/**
 * Setup a searchable dropdown with lazy loading (for large datasets like clients)
 * - Only renders visible items initially
 * - Loads more items as user scrolls
 * - Searches across ALL items when filtering
 */
function setupSearchableDropdownLazy(
  input: HTMLInputElement,
  dropdown: HTMLDivElement,
  hiddenInput: HTMLInputElement | null,
  items: Array<{ code: string; name: string }>,
  displayField: 'name' | 'code'
): void {
  let highlightedIndex = -1;
  let filteredItems = [...items];
  let renderedCount = 0;
  const BATCH_SIZE = 30; // Items to load per batch
  const INITIAL_LOAD = 30; // Initial items to show

  // Render dropdown items with lazy loading
  function renderDropdown(filter: string = '', loadMore: boolean = false): void {
    const searchTerm = filter.toLowerCase().trim();

    // Filter ALL items (search over complete dataset)
    if (searchTerm) {
      filteredItems = items.filter(item =>
        item.name.toLowerCase().includes(searchTerm) ||
        item.code.toLowerCase().includes(searchTerm)
      );
    } else {
      filteredItems = [...items];
    }

    // If not loading more, reset the dropdown
    if (!loadMore) {
      dropdown.innerHTML = '';
      renderedCount = 0;
    }

    if (filteredItems.length === 0) {
      dropdown.innerHTML = '<div class="dropdown-empty">No se encontraron resultados</div>';
      return;
    }

    // Calculate how many items to render
    const startIndex = renderedCount;
    const endIndex = Math.min(startIndex + (loadMore ? BATCH_SIZE : INITIAL_LOAD), filteredItems.length);

    // Render items
    for (let i = startIndex; i < endIndex; i++) {
      const item = filteredItems[i];
      const div = document.createElement('div');
      div.className = 'dropdown-item';
      div.dataset.value = item[displayField];
      div.dataset.index = i.toString();
      // Show code and name for clients
      div.textContent = `${item.code} - ${item.name}`;

      div.addEventListener('click', () => {
        selectItem(item);
      });

      dropdown.appendChild(div);
    }

    renderedCount = endIndex;

    // Show "load more" indicator if there are more items
    const existingMore = dropdown.querySelector('.dropdown-more');
    if (existingMore) existingMore.remove();

    if (renderedCount < filteredItems.length) {
      const moreDiv = document.createElement('div');
      moreDiv.className = 'dropdown-empty dropdown-more';
      moreDiv.textContent = `Mostrando ${renderedCount} de ${filteredItems.length}. Scroll para ver más...`;
      dropdown.appendChild(moreDiv);
    }
  }

  // Select an item
  function selectItem(item: { code: string; name: string }): void {
    input.value = `${item.code} - ${item.name}`;
    if (hiddenInput) hiddenInput.value = item[displayField];
    dropdown.classList.remove('active');
    highlightedIndex = -1;
  }

  // Show dropdown
  function showDropdown(): void {
    renderDropdown(input.value);
    dropdown.classList.add('active');
  }

  // Hide dropdown
  function hideDropdown(): void {
    dropdown.classList.remove('active');
    highlightedIndex = -1;
    // Sync hidden input with visible input value when closing without selection
    if (hiddenInput && input.value) {
      hiddenInput.value = input.value;
    }
  }

  // Update highlight
  function updateHighlight(): void {
    const visibleItems = dropdown.querySelectorAll('.dropdown-item');
    visibleItems.forEach((item, index) => {
      item.classList.toggle('highlighted', index === highlightedIndex);
    });

    // Scroll into view
    if (highlightedIndex >= 0 && visibleItems[highlightedIndex]) {
      visibleItems[highlightedIndex].scrollIntoView({ block: 'nearest' });
    }
  }

  // Event: Focus - show dropdown
  input.addEventListener('focus', () => {
    showDropdown();
  });

  // Event: Input - filter items (searches ALL items)
  input.addEventListener('input', () => {
    renderDropdown(input.value);
    dropdown.classList.add('active');
    highlightedIndex = -1;
  });

  // Event: Scroll - lazy load more items
  dropdown.addEventListener('scroll', () => {
    const scrollBottom = dropdown.scrollTop + dropdown.clientHeight;
    const scrollHeight = dropdown.scrollHeight;

    // Load more when user scrolls near the bottom (within 50px)
    if (scrollHeight - scrollBottom < 50 && renderedCount < filteredItems.length) {
      renderDropdown(input.value, true);
    }
  });

  // Event: Keydown - navigation
  input.addEventListener('keydown', (e) => {
    const visibleItems = dropdown.querySelectorAll('.dropdown-item');

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        if (!dropdown.classList.contains('active')) {
          showDropdown();
        }
        highlightedIndex = Math.min(highlightedIndex + 1, visibleItems.length - 1);
        updateHighlight();
        break;

      case 'ArrowUp':
        e.preventDefault();
        highlightedIndex = Math.max(highlightedIndex - 1, 0);
        updateHighlight();
        break;

      case 'Enter':
        e.preventDefault();
        if (highlightedIndex >= 0 && highlightedIndex < renderedCount) {
          selectItem(filteredItems[highlightedIndex]);
        }
        break;

      case 'Escape':
        hideDropdown();
        break;

      case 'Tab':
        hideDropdown();
        break;
    }
  });

  // Event: Click outside - hide dropdown
  document.addEventListener('click', (e) => {
    if (!input.contains(e.target as Node) && !dropdown.contains(e.target as Node)) {
      hideDropdown();
    }
  });
}

/**
 * Setup a searchable dropdown with lazy loading - NAME ONLY display (for clients)
 * Shows only the name without the code
 */
function setupSearchableDropdownLazyNameOnly(
  input: HTMLInputElement,
  dropdown: HTMLDivElement,
  hiddenInput: HTMLInputElement | null,
  items: Array<{ code: string; name: string }>
): void {
  let highlightedIndex = -1;
  let filteredItems = [...items];
  let renderedCount = 0;
  const BATCH_SIZE = 30;
  const INITIAL_LOAD = 30;

  function renderDropdown(filter: string = '', loadMore: boolean = false): void {
    const searchTerm = filter.toLowerCase().trim();

    if (searchTerm) {
      filteredItems = items.filter(item =>
        item.name.toLowerCase().includes(searchTerm) ||
        item.code.toLowerCase().includes(searchTerm)
      );
    } else {
      filteredItems = [...items];
    }

    if (!loadMore) {
      dropdown.innerHTML = '';
      renderedCount = 0;
    }

    if (filteredItems.length === 0) {
      dropdown.innerHTML = '<div class="dropdown-empty">No se encontraron resultados</div>';
      return;
    }

    const startIndex = renderedCount;
    const endIndex = Math.min(startIndex + (loadMore ? BATCH_SIZE : INITIAL_LOAD), filteredItems.length);

    for (let i = startIndex; i < endIndex; i++) {
      const item = filteredItems[i];
      const div = document.createElement('div');
      div.className = 'dropdown-item';
      div.dataset.value = item.name;
      div.dataset.index = i.toString();
      // Show only name (no code)
      div.textContent = item.name;

      div.addEventListener('click', () => selectItem(item));
      dropdown.appendChild(div);
    }

    renderedCount = endIndex;

    const existingMore = dropdown.querySelector('.dropdown-more');
    if (existingMore) existingMore.remove();

    if (renderedCount < filteredItems.length) {
      const moreDiv = document.createElement('div');
      moreDiv.className = 'dropdown-empty dropdown-more';
      moreDiv.textContent = `Mostrando ${renderedCount} de ${filteredItems.length}. Scroll para ver más...`;
      dropdown.appendChild(moreDiv);
    }
  }

  function selectItem(item: { code: string; name: string }): void {
    // Show only name when selected
    input.value = item.name;
    if (hiddenInput) hiddenInput.value = item.name;
    dropdown.classList.remove('active');
    highlightedIndex = -1;
  }

  function showDropdown(): void {
    renderDropdown(input.value);
    dropdown.classList.add('active');
  }

  function hideDropdown(): void {
    dropdown.classList.remove('active');
    highlightedIndex = -1;
    // Sync hidden input with visible input value when closing without selection
    if (hiddenInput && input.value) {
      hiddenInput.value = input.value;
    }
  }

  function updateHighlight(): void {
    const visibleItems = dropdown.querySelectorAll('.dropdown-item');
    visibleItems.forEach((item, index) => {
      item.classList.toggle('highlighted', index === highlightedIndex);
    });
    if (highlightedIndex >= 0 && visibleItems[highlightedIndex]) {
      visibleItems[highlightedIndex].scrollIntoView({ block: 'nearest' });
    }
  }

  input.addEventListener('focus', () => showDropdown());

  input.addEventListener('input', () => {
    renderDropdown(input.value);
    dropdown.classList.add('active');
    highlightedIndex = -1;
  });

  dropdown.addEventListener('scroll', () => {
    const scrollBottom = dropdown.scrollTop + dropdown.clientHeight;
    const scrollMax = dropdown.scrollHeight - 50;

    if (scrollBottom >= scrollMax && renderedCount < filteredItems.length) {
      renderDropdown(input.value, true);
    }
  });

  input.addEventListener('keydown', (e) => {
    const visibleItems = dropdown.querySelectorAll('.dropdown-item');

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        if (!dropdown.classList.contains('active')) showDropdown();
        highlightedIndex = Math.min(highlightedIndex + 1, visibleItems.length - 1);
        updateHighlight();
        break;
      case 'ArrowUp':
        e.preventDefault();
        highlightedIndex = Math.max(highlightedIndex - 1, 0);
        updateHighlight();
        break;
      case 'Enter':
        e.preventDefault();
        if (highlightedIndex >= 0 && filteredItems[highlightedIndex]) {
          selectItem(filteredItems[highlightedIndex]);
        }
        break;
      case 'Escape':
      case 'Tab':
        hideDropdown();
        break;
    }
  });

  document.addEventListener('click', (e) => {
    if (!input.contains(e.target as Node) && !dropdown.contains(e.target as Node)) {
      hideDropdown();
    }
  });
}

/**
 * Setup a searchable dropdown for clients with displayName support
 * Shows full format like iTraffic: "CODE - NAME - Cuit:XXX"
 */
function setupSearchableDropdownForClients(
  input: HTMLInputElement,
  dropdown: HTMLDivElement,
  hiddenInput: HTMLInputElement | null,
  items: Array<{ code: string; name: string; displayName?: string }>
): void {
  let highlightedIndex = -1;
  let filteredItems = [...items];
  let renderedCount = 0;
  const BATCH_SIZE = 30;
  const INITIAL_LOAD = 30;

  function renderDropdown(filter: string = '', loadMore: boolean = false): void {
    const searchTerm = filter.toLowerCase().trim();

    if (searchTerm) {
      filteredItems = items.filter(item =>
        item.name.toLowerCase().includes(searchTerm) ||
        item.code.toLowerCase().includes(searchTerm) ||
        (item.displayName && item.displayName.toLowerCase().includes(searchTerm))
      );
    } else {
      filteredItems = [...items];
    }

    if (!loadMore) {
      dropdown.innerHTML = '';
      renderedCount = 0;
    }

    if (filteredItems.length === 0) {
      dropdown.innerHTML = '<div class="dropdown-empty">No se encontraron resultados</div>';
      return;
    }

    const startIndex = renderedCount;
    const endIndex = Math.min(startIndex + (loadMore ? BATCH_SIZE : INITIAL_LOAD), filteredItems.length);

    for (let i = startIndex; i < endIndex; i++) {
      const item = filteredItems[i];
      const div = document.createElement('div');
      div.className = 'dropdown-item';
      div.dataset.value = item.displayName || `${item.code} - ${item.name}`;
      div.dataset.index = i.toString();
      // Show displayName if available, otherwise "CODE - NAME"
      div.textContent = item.displayName || `${item.code} - ${item.name}`;

      div.addEventListener('click', () => selectItem(item));
      dropdown.appendChild(div);
    }

    renderedCount = endIndex;

    const existingMore = dropdown.querySelector('.dropdown-more');
    if (existingMore) existingMore.remove();

    if (renderedCount < filteredItems.length) {
      const moreDiv = document.createElement('div');
      moreDiv.className = 'dropdown-empty dropdown-more';
      moreDiv.textContent = `Mostrando ${renderedCount} de ${filteredItems.length}. Scroll para ver más...`;
      dropdown.appendChild(moreDiv);
    }
  }

  function selectItem(item: { code: string; name: string; displayName?: string }): void {
    // Show displayName when selected (full format like iTraffic)
    const displayValue = item.displayName || `${item.code} - ${item.name}`;
    input.value = displayValue;
    if (hiddenInput) hiddenInput.value = displayValue;
    dropdown.classList.remove('active');
    highlightedIndex = -1;
  }

  function showDropdown(): void {
    renderDropdown(input.value);
    dropdown.classList.add('active');
  }

  function hideDropdown(): void {
    dropdown.classList.remove('active');
    highlightedIndex = -1;
    // Sync hidden input with visible input value when closing without selection
    if (hiddenInput && input.value) {
      hiddenInput.value = input.value;
    }
  }

  function updateHighlight(): void {
    const visibleItems = dropdown.querySelectorAll('.dropdown-item');
    visibleItems.forEach((item, index) => {
      item.classList.toggle('highlighted', index === highlightedIndex);
    });
    if (highlightedIndex >= 0 && visibleItems[highlightedIndex]) {
      visibleItems[highlightedIndex].scrollIntoView({ block: 'nearest' });
    }
  }

  input.addEventListener('focus', () => showDropdown());

  input.addEventListener('input', () => {
    renderDropdown(input.value);
    dropdown.classList.add('active');
    highlightedIndex = -1;
  });

  dropdown.addEventListener('scroll', () => {
    const scrollBottom = dropdown.scrollTop + dropdown.clientHeight;
    const scrollMax = dropdown.scrollHeight - 50;

    if (scrollBottom >= scrollMax && renderedCount < filteredItems.length) {
      renderDropdown(input.value, true);
    }
  });

  input.addEventListener('keydown', (e) => {
    const visibleItems = dropdown.querySelectorAll('.dropdown-item');

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        if (!dropdown.classList.contains('active')) showDropdown();
        highlightedIndex = Math.min(highlightedIndex + 1, visibleItems.length - 1);
        updateHighlight();
        break;
      case 'ArrowUp':
        e.preventDefault();
        highlightedIndex = Math.max(highlightedIndex - 1, 0);
        updateHighlight();
        break;
      case 'Enter':
        e.preventDefault();
        if (highlightedIndex >= 0 && filteredItems[highlightedIndex]) {
          selectItem(filteredItems[highlightedIndex]);
        }
        break;
      case 'Escape':
      case 'Tab':
        hideDropdown();
        break;
    }
  });

  document.addEventListener('click', (e) => {
    if (!input.contains(e.target as Node) && !dropdown.contains(e.target as Node)) {
      hideDropdown();
    }
  });
}

/**
 * Setup a searchable dropdown with alias support (for currencies)
 * Allows searching by common abbreviations like USD -> DOLARES
 */
function setupSearchableDropdownWithAliases(
  input: HTMLInputElement,
  dropdown: HTMLDivElement,
  hiddenInput: HTMLInputElement | null,
  items: Array<{ code: string; name: string }>,
  aliases: Record<string, string[]>
): void {
  let highlightedIndex = -1;
  let filteredItems = [...items];

  // Build reverse alias map: USD -> DOLARES
  const reverseAliases: Record<string, string> = {};
  for (const [mainName, aliasList] of Object.entries(aliases)) {
    for (const alias of aliasList) {
      reverseAliases[alias.toUpperCase()] = mainName.toUpperCase();
    }
  }

  function renderDropdown(filter: string = ''): void {
    const searchTerm = filter.toLowerCase().trim();
    const searchTermUpper = filter.toUpperCase().trim();

    // Check if search term matches an alias
    const resolvedTerm = reverseAliases[searchTermUpper] || searchTerm;

    if (searchTerm) {
      filteredItems = items.filter(item => {
        const nameMatch = item.name.toLowerCase().includes(searchTerm) ||
          item.name.toLowerCase().includes(resolvedTerm.toLowerCase());
        const codeMatch = item.code.toLowerCase().includes(searchTerm);

        // Also check if any alias of this item matches
        const itemAliases = aliases[item.name.toUpperCase()] || [];
        const aliasMatch = itemAliases.some(a => a.toLowerCase().includes(searchTerm));

        return nameMatch || codeMatch || aliasMatch;
      });
    } else {
      filteredItems = [...items];
    }

    dropdown.innerHTML = '';

    if (filteredItems.length === 0) {
      dropdown.innerHTML = '<div class="dropdown-empty">No se encontraron resultados</div>';
      return;
    }

    const displayItems = filteredItems.slice(0, 50);

    displayItems.forEach((item, index) => {
      const div = document.createElement('div');
      div.className = 'dropdown-item';
      div.dataset.value = item.name;
      div.dataset.index = index.toString();
      div.textContent = item.name;

      div.addEventListener('click', () => selectItem(item));
      dropdown.appendChild(div);
    });

    if (filteredItems.length > 50) {
      const moreDiv = document.createElement('div');
      moreDiv.className = 'dropdown-empty';
      moreDiv.textContent = `... y ${filteredItems.length - 50} más. Escribe para filtrar.`;
      dropdown.appendChild(moreDiv);
    }
  }

  function selectItem(item: { code: string; name: string }): void {
    input.value = item.name;
    if (hiddenInput) hiddenInput.value = item.name;
    dropdown.classList.remove('active');
    highlightedIndex = -1;
  }

  function showDropdown(): void {
    renderDropdown(input.value);
    dropdown.classList.add('active');
  }

  function hideDropdown(): void {
    dropdown.classList.remove('active');
    highlightedIndex = -1;
    // Sync hidden input with visible input value when closing without selection
    if (hiddenInput && input.value) {
      hiddenInput.value = input.value;
    }
  }

  function updateHighlight(): void {
    const visibleItems = dropdown.querySelectorAll('.dropdown-item');
    visibleItems.forEach((item, index) => {
      item.classList.toggle('highlighted', index === highlightedIndex);
    });
    if (highlightedIndex >= 0 && visibleItems[highlightedIndex]) {
      visibleItems[highlightedIndex].scrollIntoView({ block: 'nearest' });
    }
  }

  input.addEventListener('focus', () => showDropdown());

  input.addEventListener('input', () => {
    renderDropdown(input.value);
    dropdown.classList.add('active');
    highlightedIndex = -1;
  });

  input.addEventListener('keydown', (e) => {
    const visibleItems = dropdown.querySelectorAll('.dropdown-item');

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        if (!dropdown.classList.contains('active')) showDropdown();
        highlightedIndex = Math.min(highlightedIndex + 1, visibleItems.length - 1);
        updateHighlight();
        break;
      case 'ArrowUp':
        e.preventDefault();
        highlightedIndex = Math.max(highlightedIndex - 1, 0);
        updateHighlight();
        break;
      case 'Enter':
        e.preventDefault();
        if (highlightedIndex >= 0 && filteredItems[highlightedIndex]) {
          selectItem(filteredItems[highlightedIndex]);
        }
        break;
      case 'Escape':
      case 'Tab':
        hideDropdown();
        break;
    }
  });

  document.addEventListener('click', (e) => {
    if (!input.contains(e.target as Node) && !dropdown.contains(e.target as Node)) {
      hideDropdown();
    }
  });
}

/**
 * Toggle header visibility
 */
function toggleHeader(): void {
  const header = document.getElementById('collapsible-header')!;
  const btn = document.getElementById('toggle-header-btn')!;

  if (header.style.display === 'none') {
    header.style.display = 'block';
    btn.textContent = '🔼';
  } else {
    header.style.display = 'none';
    btn.textContent = '🔽';
  }
}

/**
 * Get current user email from Outlook
 */
function getCurrentUserEmail(): string | null {
  try {
    // Get email from Outlook mailbox
    const userProfile = Office.context.mailbox?.userProfile;
    if (userProfile && userProfile.emailAddress) {
      return userProfile.emailAddress.toLowerCase();
    }
    return null;
  } catch (error) {
    console.error('❌ Error getting current user email:', error);
    return null;
  }
}

/**
 * Check if user is already registered
 */
async function checkRegistrationStatus(): Promise<void> {
  try {
    // Get current Outlook user email
    const currentEmail = getCurrentUserEmail();
    console.log('📧 Current Outlook user:', currentEmail);

    // Check if we have a cached userId in localStorage
    const cachedUserId = localStorage.getItem('email-classifier-userId');
    const cachedEmail = localStorage.getItem('email-classifier-email');

    if (!cachedUserId || !cachedEmail) {
      console.log('⏭️ No cached user found - user must login manually');
      // Hide loading, show registration UI
      document.getElementById('loading')!.style.display = 'none';
      document.getElementById('app-content')!.style.display = 'block';
      return;
    }

    // Validate that cached email matches current Outlook email
    if (currentEmail && cachedEmail.toLowerCase() !== currentEmail) {
      console.log('⚠️ Cached email does not match current Outlook email');
      console.log('   Cached:', cachedEmail);
      console.log('   Current:', currentEmail);
      console.log('   Clearing cache and requiring re-login');

      localStorage.removeItem('email-classifier-userId');
      localStorage.removeItem('email-classifier-email');

      // Hide loading, show registration UI
      document.getElementById('loading')!.style.display = 'none';
      document.getElementById('app-content')!.style.display = 'block';
      return;
    }

    console.log('🔍 Found cached user:', cachedEmail, '- verifying with backend...');

    // Verify with production backend where user data is stored
    const response = await fetch(`${config.dialogBaseUrl}/auth/status?userId=${cachedUserId}`);

    if (!response.ok) {
      console.log('⚠️ User verification failed - clearing cache');
      localStorage.removeItem('email-classifier-userId');
      localStorage.removeItem('email-classifier-email');

      // Hide loading, show registration UI
      document.getElementById('loading')!.style.display = 'none';
      document.getElementById('app-content')!.style.display = 'block';
      return;
    }

    const data = await response.json();

    if (data.isRegistered && data.user) {
      console.log('✅ User verified - updating UI');

      // Update state
      state.isRegistered = true;
      state.userId = data.user.userId;
      state.email = data.user.email;

      // Hide loading, show app
      document.getElementById('loading')!.style.display = 'none';
      document.getElementById('app-content')!.style.display = 'block';

      // Update UI
      updateUIForRegisteredUser(data.user);
    } else {
      console.log('⚠️ User not found in backend - clearing cache');
      localStorage.removeItem('email-classifier-userId');
      localStorage.removeItem('email-classifier-email');

      // Hide loading, show registration UI
      document.getElementById('loading')!.style.display = 'none';
      document.getElementById('app-content')!.style.display = 'block';
    }
  } catch (error) {
    console.error('❌ Error checking registration status:', error);
    // Hide loading, show registration UI
    document.getElementById('loading')!.style.display = 'none';
    document.getElementById('app-content')!.style.display = 'block';
  }
}

/**
 * Handle registration button click - SSO Flow
 */
async function handleRegister(): Promise<void> {
  // Prevent multiple simultaneous authentication attempts
  if (state.isAuthInProgress) {
    console.log('⚠️ Authentication already in progress, ignoring request');
    showError('Authentication is already in progress. Please wait.');
    return;
  }

  // Implement exponential backoff for retries
  const now = Date.now();
  if (state.lastAuthAttempt && (now - state.lastAuthAttempt) < 2000) {
    console.log('⚠️ Too many authentication attempts, please wait');
    showError('Please wait a moment before trying again.');
    return;
  }

  hideError();
  setLoading(true);
  state.isAuthInProgress = true;
  state.lastAuthAttempt = now;

  try {
    console.log('🔐 Starting SSO authentication... [VERSION 2024-12-02-DIALOG-PREFERRED]');

    // Get current Outlook email for validation
    const currentOutlookEmail = getCurrentUserEmail();
    if (!currentOutlookEmail) {
      throw new Error('No se pudo obtener el email de Outlook. Por favor, asegúrate de que Outlook esté completamente cargado.');
    }

    console.log('📧 Current Outlook email:', currentOutlookEmail);

    // TEMPORARY: Skip SSO and always use dialog flow
    // SSO + OBO flow has issues with token exchange, dialog flow works reliably
    // TODO: Fix OBO flow and re-enable SSO for better UX
    console.log('⚠️ Skipping SSO, using dialog authentication (more reliable)...');
    await handleRegisterViaDialog({ forceConsent: true });
    return;

    // Get SSO token from Office (DISABLED - code below kept for future re-enablement)
    /*
    console.log('🔑 Requesting SSO token from Office.auth.getAccessToken()...');
    let ssoToken: string;
    
    try {
      ssoToken = await getOfficeSSToken();
      console.log('✅ SSO token obtained successfully');
      console.log('   Token length:', ssoToken.length);
    } catch (ssoError: any) {
      console.error('❌ SSO failed:', ssoError);
      
      // If SSO fails, fall back to dialog flow
      // 13001: No user signed in
      // 13012: API not compatible / sideload
      // 13000: General SSO not available
      // 13003: SSO not supported in this version
      if (ssoError.code === 13001 || ssoError.code === 13012 || ssoError.code === 13000 || ssoError.code === 13003) {
        console.log('⚠️ SSO not available (code: ' + ssoError.code + '), falling back to dialog auth...');
        // Use forceConsent=true to ensure refresh tokens for personal Microsoft accounts
        await handleRegisterViaDialog({ forceConsent: true });
        return;
      }
      
      throw ssoError;
    }
    
    // Exchange SSO token for Graph token via backend OBO flow
    // Must use production backend which has the client_secret for OBO
    console.log('🔄 Exchanging SSO token for Graph tokens via OBO flow...');
    const exchangeResponse = await fetch(`${config.dialogBaseUrl}/auth/exchange-token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ ssoToken })
    });

    if (!exchangeResponse.ok) {
      const errorData = await exchangeResponse.json().catch(() => ({ error: 'Unknown error' }));
      throw new Error(`Token exchange failed: ${errorData.error || exchangeResponse.statusText}`);
    }

    const exchangeData = await exchangeResponse.json();
    console.log('✅ Token exchange successful:', exchangeData);
    
    // Validate email matches
    if (exchangeData.email.toLowerCase() !== currentOutlookEmail) {
      console.error('❌ Email mismatch after exchange!');
      console.error('   Backend email:', exchangeData.email);
      console.error('   Outlook email:', currentOutlookEmail);
      throw new Error(
        `Error de coherencia: El email autenticado (${exchangeData.email}) no coincide con tu cuenta de Outlook (${currentOutlookEmail})`
      );
    }
    
    // Save to localStorage AND roamingSettings for persistence across contexts
    localStorage.setItem('email-classifier-userId', exchangeData.userId);
    localStorage.setItem('email-classifier-email', exchangeData.email);
    
    // Also save to roamingSettings to share with ribbon commands
    const settings = Office.context.roamingSettings;
    settings.set('email-classifier-userId', exchangeData.userId);
    settings.set('email-classifier-email', exchangeData.email);
    await settings.saveAsync();
    
    console.log('💾 User credentials saved to localStorage and roamingSettings');
    
    // Update state
    state.isRegistered = true;
    state.userId = exchangeData.userId;
    state.email = exchangeData.email;

    // Update UI
    updateUIForRegisteredUser({
      email: exchangeData.email,
      id: exchangeData.userId
    });
    
    // Show success message
    const message = exchangeData.hasRefreshToken
      ? '🎉 Autenticación completa! Tus emails se clasificarán automáticamente sin vencimiento.'
      : '🎉 Autenticación completa! Tus emails se clasificarán automáticamente.';
    
    showSuccessMessage(message);
    */

  } catch (error: any) {
    console.error('❌ Registration error:', error);
    state.authRetryCount++;
    handleRegistrationError(error);
  } finally {
    setLoading(false);
    state.isAuthInProgress = false;
  }
}

/**
 * Get SSO token from Office.auth.getAccessToken()
 */
function getOfficeSSToken(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true
    })
      .then(token => resolve(token))
      .catch(error => reject(error));
  });
}

/**
 * Fallback: Handle registration via dialog (for sideload scenarios)
 */
interface DialogOptions {
  forceConsent?: boolean;
}

async function handleRegisterViaDialog(options: DialogOptions = {}): Promise<void> {
  try {
    console.log('🪟 Using dialog-based authentication fallback');
    if (options.forceConsent) {
      console.log('🙏 Forcing consent screen for offline_access permission');
    }

    // Additional check to prevent dialog loops
    if (state.authRetryCount > 3) {
      throw new Error('Too many authentication attempts. Please refresh the page and try again.');
    }

    const dialogResult = await getAccessTokenViaDialog(options);
    if (!dialogResult) {
      throw new Error('Failed to obtain authorization via dialog');
    }

    console.log('✅ Dialog returned payload:', {
      status: dialogResult.status,
      hasCode: !!dialogResult.authorizationCode,
      hasToken: !!dialogResult.token
    });

    // Reset retry count on any successful dialog response
    state.authRetryCount = 0;

    const currentOutlookEmail = getCurrentUserEmail();

    if (dialogResult.authorizationCode) {
      await completeRegistrationWithAuthorizationCode(dialogResult, currentOutlookEmail, options);
      return;
    }

    if (dialogResult.token) {
      console.warn('⚠️ Received legacy token payload from dialog. Proceeding with deprecated flow.');
      await completeLegacyRegistration(dialogResult, currentOutlookEmail);
      return;
    }

    throw new Error('Invalid authentication response from dialog');
  } catch (error: any) {
    console.error('❌ Dialog registration error:', error);
    throw error;
  }
}

async function completeRegistrationWithAuthorizationCode(
  dialogResult: any,
  currentOutlookEmail: string | null,
  options: DialogOptions = {}
): Promise<void> {
  console.log('📡 Redeeming authorization code via backend (confidential client)...');
  console.log('ℹ️ Backend will use client_secret to get access_token + refresh_token');

  // Use dialogBaseUrl for token exchange since the auth code was issued there
  const redeemResponse = await fetch(`${config.dialogBaseUrl}/auth/redeem-code`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      code: dialogResult.authorizationCode,
      redirectUri: dialogResult.redirectUri
    })
  });

  if (!redeemResponse.ok) {
    const errorData = await redeemResponse.json().catch((data: any) => ({ data: data }));
    console.error('❌ Authorization code redemption failed:', errorData);

    if (redeemResponse.status === 403 && errorData.requiresInteractiveAuth) {
      console.error('❌ Consent required but cannot reopen dialog programmatically (Error 12011)');
      throw new Error(
        'Se requiere consentimiento adicional. Por favor, haz clic en "Retry" para autorizar nuevamente. ' +
        'Si el problema persiste, contacta a tu administrador de TI.'
      );
    }

    throw new Error(`No se pudo completar la autenticación: ${JSON.stringify(redeemResponse.json())}, ${JSON.stringify(dialogResult)}`);
  }

  const redeemData = await redeemResponse.json();
  console.log('✅ Backend redeemed authorization code successfully:', redeemData);

  if (!redeemData.user?.id || !redeemData.user?.email) {
    throw new Error('La respuesta del servidor no incluye la información del usuario.');
  }

  const backendEmail = redeemData.user.email.toLowerCase();
  if (currentOutlookEmail && backendEmail !== currentOutlookEmail) {
    console.error('❌ Email mismatch detected after redemption');
    throw new Error(
      `El usuario autenticado (${backendEmail}) no coincide con tu cuenta de Outlook (${currentOutlookEmail}).`
    );
  }

  await persistUserSession(redeemData.user.id, redeemData.user.email);
  state.isRegistered = true;
  state.userId = redeemData.user.id;
  state.email = redeemData.user.email;

  updateUIForRegisteredUser({
    email: redeemData.user.email,
    id: redeemData.user.id,
    stats: redeemData.user.stats
  });

  // Show appropriate message based on refresh token availability
  let successMessage = '🎉 Autenticación completa!';
  if (redeemData.hasRefreshToken) {
    successMessage += ' Tus correos se clasificarán automáticamente con acceso permanente.';
  } else {
    console.warn('⚠️ No refresh token received - user will need to re-authenticate periodically');
    successMessage += ' Tus correos se clasificarán automáticamente. Nota: Deberás volver a autenticarte después de 1 hora.';
  }
  showSuccessMessage(successMessage);
}

async function completeLegacyRegistration(messageData: any, currentOutlookEmail: string | null): Promise<void> {
  if (!messageData.email || !messageData.userId) {
    throw new Error('Invalid legacy authentication response');
  }

  if (currentOutlookEmail && messageData.email.toLowerCase() !== currentOutlookEmail) {
    throw new Error(
      `Por favor, inicia sesión con la misma cuenta de Outlook (${currentOutlookEmail}). ` +
      `Intentaste autenticarte con: ${messageData.email}`
    );
  }

  // Use dialogBaseUrl for registration since it stores user data in production CosmosDB
  const registerResponse = await fetch(`${config.dialogBaseUrl}/auth/register`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      accessToken: messageData.token,
      refreshToken: messageData.refreshToken,
      userId: messageData.userId,
      email: messageData.email,
      name: messageData.name,
      expiresOn: messageData.expiresOn
    })
  });

  if (!registerResponse.ok) {
    const errorData = await registerResponse.json().catch(() => ({ error: 'Unknown error' }));
    throw new Error(`Registration failed: ${ JSON.stringify(errorData)}`);
  }

  await persistUserSession(messageData.userId, messageData.email);
  state.isRegistered = true;
  state.userId = messageData.userId;
  state.email = messageData.email;

  updateUIForRegisteredUser({
    email: messageData.email,
    id: messageData.userId
  });

  showSuccessMessage('🎉 Autenticación completa (modo legacy). Considera actualizar la versión del complemento.');
}

async function persistUserSession(userId: string, email: string): Promise<void> {
  // Save to localStorage
  localStorage.setItem('email-classifier-userId', userId);
  localStorage.setItem('email-classifier-email', email);

  // Save to roamingSettings to share with ribbon commands
  const settings = Office.context.roamingSettings;
  settings.set('email-classifier-userId', userId);
  settings.set('email-classifier-email', email);

  // Save settings asynchronously
  return new Promise((resolve) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log('💾 User credentials saved to localStorage and roamingSettings');
      } else {
        console.error('⚠️ Failed to save roamingSettings:', result.error);
      }
      resolve();
    });
  });
}

/**
 * Handle logout button click
 */
async function handleLogout(): Promise<void> {
  try {
    console.log('🚪 Logging out user...');

    // Clear localStorage
    localStorage.removeItem('email-classifier-userId');
    localStorage.removeItem('email-classifier-email');

    // Clear roamingSettings
    const settings = Office.context.roamingSettings;
    settings.remove('email-classifier-userId');
    settings.remove('email-classifier-email');
    settings.saveAsync();

    console.log('✅ Cleared localStorage and roamingSettings');

    // Reset state
    state.isRegistered = false;
    state.userId = null;
    state.email = null;

    // Reset UI
    const statusBox = document.getElementById('status-indicator')!;
    statusBox.className = 'status-box unregistered';
    statusBox.innerHTML = `
      <div class="status-icon">⏸️</div>
      <div class="status-text">
        <h3>Not Registered</h3>
        <p>Click below to activate automatic email classification</p>
      </div>
    `;

    // Show registration button, hide logout button
    document.getElementById('registration-section')!.style.display = 'block';
    document.getElementById('logout-section')!.style.display = 'none';
    document.getElementById('stats-section')!.style.display = 'none';

    console.log('✅ Logged out from add-in');

    // Logout from Microsoft account (clear MSAL cache)
    console.log('🔓 Clearing Microsoft session...');
    const clientId = '6637590b-a6a4-4e53-b429-a766c66f03c3';
    const logoutUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/logout?post_logout_redirect_uri=${encodeURIComponent(window.location.origin)}`;

    // Open logout in a hidden dialog to clear Microsoft session
    Office.context.ui.displayDialogAsync(
      logoutUrl,
      { height: 20, width: 20, promptBeforeOpen: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          // Close dialog after 2 seconds
          setTimeout(() => {
            dialog.close();
            console.log('✅ Microsoft session cleared');
            showSuccessMessage('👋 Signed out completely. Next login will require re-authentication.');
          }, 2000);
        } else {
          console.log('⚠️ Could not open logout dialog, but local logout succeeded');
          showSuccessMessage('👋 Signed out successfully. Sign in again to re-activate auto-classification.');
        }
      }
    );

  } catch (error) {
    console.error('❌ Logout error:', error);
    showError('Failed to sign out. Please try again.');
  }
}

/**
 * Get access token using Office.js SSO with fallback to dialog auth
 */
async function getAccessToken(allowPrompt: boolean): Promise<string | null> {
  try {
    const options: Office.AuthOptions = {
      allowSignInPrompt: allowPrompt,
      allowConsentPrompt: allowPrompt,
      forMSGraphAccess: false, // We're accessing our own API, not Graph
    };

    const token = await Office.auth.getAccessToken(options);
    return token;

  } catch (error: any) {
    console.error('SSO error:', error);

    // Handle specific error codes
    if (error.code === 13001) {
      // User is not signed in
      if (allowPrompt) {
        throw new Error('Please sign in to Outlook to use this add-in');
      }
      return null;
    } else if (error.code === 13002) {
      // User aborted consent
      throw new Error('You must grant permission to use this add-in');
    } else if (error.code === 13003) {
      // SSO not supported
      throw new Error('Single sign-on is not supported in your Outlook version');
    } else if (error.code === 13006) {
      // Admin consent required
      throw new Error('This add-in requires administrator approval. Please contact your IT department.');
    } else if (error.code === 13007) {
      // Invalid grant or consent revoked
      console.log('Invalid grant or consent revoked, falling back to dialog auth');
      return await getAccessTokenViaDialog();
    } else if (error.code === 13012) {
      // API not compatible or sideloading in Outlook
      // This is EXPECTED when sideloading in Outlook - use fallback auth
      console.log('SSO unavailable (error 13012), falling back to dialog auth');
      return await getAccessTokenViaDialog();
    }

    // For any other error, try fallback auth
    console.log(`SSO failed with error ${error.code}, attempting fallback auth`);
    return await getAccessTokenViaDialog();
  }
}

/**
 * Fallback authentication using Office dialog
 */
async function getAccessTokenViaDialog(options: DialogOptions = {}): Promise<any> {
  // Close any existing dialog first to prevent "already has an active dialog" error
  closeActiveDialog();

  return new Promise((resolve, reject) => {
    const params = new URLSearchParams();
    if (options.forceConsent) {
      params.set('forceConsent', 'true');
    }
    const queryString = params.toString();
    // Office dialogs REQUIRE HTTPS - use dialogBaseUrl (always production)
    const dialogUrl = queryString
      ? `${config.dialogBaseUrl}/login.html?${queryString}`
      : `${config.dialogBaseUrl}/login.html`;

    console.log('Opening authentication dialog:', dialogUrl);

    const dialogOptions: Office.DialogOptions = {
      height: 60,
      width: 30,
      promptBeforeOpen: false,
      displayInIframe: false // Azure AD blocks iframing; always use pop-out window
    };

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      dialogOptions,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error('Failed to open dialog:', result.error);

          // Handle "already has an active dialog" error (12007)
          if (result.error && result.error.code === 12007) {
            console.log('Dialog already active, clearing reference and retrying...');
            activeDialog = null;
            // Retry once after clearing
            setTimeout(() => {
              reject(new Error('Ya hay un dialogo de autenticacion abierto. Cierra la ventana anterior y vuelve a intentar.'));
            }, 500);
            return;
          }

          if (result.error && result.error.code === 12011) {
            const guidance = 'No pudimos abrir la ventana del consentimiento porque el explorador considera diferente la zona de seguridad. Abre Outlook en Edge/Chrome actualizado o agrega https://app-itraffic-rpa.whiteflower-4df565a8.eastus2.azurecontainerapps.io a los sitios de confianza.';
            reject(new Error(guidance));
            return;
          }

          reject(new Error(`Failed to open authentication dialog: ${result.error.message || result.error.code}`));
          return;
        }

        // Store reference to active dialog
        const dialog = result.value;
        activeDialog = dialog;
        let dialogClosed = false;

        // Function to check localStorage for auth code (fallback mechanism)
        const checkLocalStorageForAuth = (): any | null => {
          try {
            const stored = localStorage.getItem('office_auth_pending');
            if (stored) {
              const data = JSON.parse(stored);
              // Only use if less than 5 minutes old
              if (data.timestamp && (Date.now() - data.timestamp) < 300000) {
                console.log('Found auth code in localStorage (fallback)');
                localStorage.removeItem('office_auth_pending');
                return data;
              } else {
                // Clean up old entry
                localStorage.removeItem('office_auth_pending');
              }
            }
          } catch (e) {
            console.error('Error checking localStorage:', e);
          }
          return null;
        };

        // Function to check server bridge for auth code (handles cross-context scenario)
        // Bridge must use same URL as dialog (production) since that's where auth code is stored
        const checkServerBridgeForAuth = async (): Promise<any | null> => {
          try {
            const response = await fetch(`${config.dialogBaseUrl}/auth/bridge?email=pending`);
            if (response.ok) {
              const data = await response.json();
              if (data.success && data.code) {
                console.log('📦 Found auth code in server bridge');
                return {
                  status: 'code',
                  authorizationCode: data.code,
                  redirectUri: data.redirectUri,
                  state: data.state
                };
              }
            }
          } catch (e) {
            // Silently fail - bridge might not exist or have no data
          }
          return null;
        };

        // Poll localStorage and server bridge every 2 seconds as fallback
        const localStorageInterval = setInterval(async () => {
          if (dialogClosed) {
            clearInterval(localStorageInterval);
            return;
          }

          // First check localStorage
          let authData = checkLocalStorageForAuth();

          // If not in localStorage, check server bridge
          if (!authData) {
            authData = await checkServerBridgeForAuth();
          }

          if (authData && authData.status === 'code' && authData.authorizationCode) {
            console.log('Retrieved auth code from fallback mechanism');
            clearInterval(localStorageInterval);
            clearTimeout(timeout);
            dialogClosed = true;
            activeDialog = null;
            try { dialog.close(); } catch (e) { /* dialog might already be closed */ }
            resolve(authData);
          }
        }, 2000);

        // Set a timeout to prevent infinite waiting
        const timeout = setTimeout(async () => {
          if (!dialogClosed) {
            // Before failing, check localStorage one more time
            let authData = checkLocalStorageForAuth();
            if (!authData) {
              authData = await checkServerBridgeForAuth();
            }

            if (authData && authData.status === 'code' && authData.authorizationCode) {
              console.log('Retrieved auth code from fallback on timeout');
              clearInterval(localStorageInterval);
              dialogClosed = true;
              activeDialog = null;
              try { dialog.close(); } catch (e) { /* dialog might already be closed */ }
              resolve(authData);
              return;
            }

            console.error('Dialog timeout - no response received');
            clearInterval(localStorageInterval);
            dialogClosed = true;
            activeDialog = null;
            dialog.close();
            reject(new Error('Authentication timeout. Please try again.'));
          }
        }, 120000); // 2 minute timeout

        // Listen for messages from the dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
          if (dialogClosed) return;

          clearTimeout(timeout);
          clearInterval(localStorageInterval);
          dialogClosed = true;
          activeDialog = null;
          dialog.close();

          try {
            const message = JSON.parse(arg.message);

            console.log('Received message from dialog:', {
              status: message.status,
              hasCode: !!message.authorizationCode,
              hasToken: !!message.token
            });

            if (message.status === 'code' && message.authorizationCode) {
              resolve(message);
              return;
            }

            if (message.status === 'success' && message.token) {
              resolve(message);
              return;
            }

            reject(new Error(message.error || 'Authentication failed'));
          } catch (error) {
            console.error('❌ Error parsing dialog message:', error);
            reject(new Error('Invalid response from authentication dialog'));
          }
        });

        // Handle dialog closed by user
        dialog.addEventHandler(Office.EventType.DialogEventReceived, async (arg: any) => {
          if (dialogClosed) return;

          clearTimeout(timeout);
          clearInterval(localStorageInterval);

          console.log('Dialog event received:', arg);

          // Before failing, check localStorage and server bridge for auth code
          let authData = checkLocalStorageForAuth();
          if (!authData) {
            authData = await checkServerBridgeForAuth();
          }

          if (authData && authData.status === 'code' && authData.authorizationCode) {
            console.log('Retrieved auth code from fallback after dialog event');
            dialogClosed = true;
            activeDialog = null;
            resolve(authData);
            return;
          }

          dialogClosed = true;
          activeDialog = null;

          if (arg.error === 12006) {
            // User closed dialog - wait a bit and check again
            setTimeout(async () => {
              let delayedAuthData = checkLocalStorageForAuth();
              if (!delayedAuthData) {
                delayedAuthData = await checkServerBridgeForAuth();
              }

              if (delayedAuthData && delayedAuthData.status === 'code' && delayedAuthData.authorizationCode) {
                console.log('✅ Retrieved auth code from fallback after dialog close delay');
                resolve(delayedAuthData);
              } else {
                reject(new Error('Authentication cancelled by user'));
              }
            }, 1500);
          } else if (arg.error === 12002) {
            // Dialog navigation error
            reject(new Error('Dialog navigation failed. Please check your internet connection.'));
          } else {
            reject(new Error(`Authentication dialog error: ${arg.error || 'Unknown error'}`));
          }
        });
      }
    );
  });
}

/**
 * Update UI for registered user
 */
function updateUIForRegisteredUser(data: any): void {
  console.log('🎨 updateUIForRegisteredUser called with data:', data);

  // Update status indicator
  const statusBox = document.getElementById('status-indicator')!;
  statusBox.className = 'status-box registered';
  statusBox.innerHTML = `
    <div class="status-icon">✅</div>
    <div class="status-text">
      <h3>Active</h3>
      <p>Auto-classification enabled for ${data.email || 'your account'}</p>
    </div>
  `;

  // Hide registration button, show logout button
  document.getElementById('registration-section')!.style.display = 'none';
  document.getElementById('logout-section')!.style.display = 'block';

  // Show stats section
  document.getElementById('stats-section')!.style.display = 'block';

  // Show extraction section
  const extractionSection = document.getElementById('extraction-section')!;
  extractionSection.style.display = 'block';
  console.log('✅ Extraction section displayed:', extractionSection.style.display);

  // Update stats if available
  if (data.stats) {
    document.getElementById('total-classified')!.textContent = data.stats.totalClassified || '0';
    document.getElementById('categories-count')!.textContent = data.stats.categoriesCount || '0';
  }
}

/**
 * Update field status (highlight if missing)
 * Handles both Inputs and Selects. For Selects, adds the option if missing.
 * For searchable dropdowns, finds the best match from master data.
 */
function updateFieldStatus(elementId: string, value: string | null): void {
  const element = document.getElementById(elementId) as HTMLInputElement | HTMLSelectElement;
  if (!element) return;

  const valToSet = value || '';

  // Helper function to normalize text (remove accents)
  const normalizeText = (text: string): string => {
    return text.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim();
  };

  // Special handling for searchable dropdown fields - find best match from master data
  if (element.tagName === 'INPUT' && valToSet) {
    const searchableFields: Record<string, Array<{ code: string; name: string }>> = {
      'input-seller': masterData.sellers,
      'input-client': masterData.clients,
      'input-currency': masterData.currencies,
      'input-status': masterData.statuses,
      'input-reservation-type': masterData.reservationTypes,
    };

    const dataSource = searchableFields[elementId];
    if (dataSource && dataSource.length > 0) {
      const searchTerm = normalizeText(valToSet);
      let match: { code: string; name: string } | undefined;

      // 1. Try exact match by code
      match = dataSource.find(item =>
        normalizeText(item.code) === searchTerm
      );

      // 2. Try exact match by name (normalized)
      if (!match) {
        match = dataSource.find(item =>
          normalizeText(item.name) === searchTerm
        );
      }

      // 3. For status/reservationType, check if name contains [CODE] and match
      // Format is "NAME [CODE]" so search for items where the code in brackets matches
      if (!match && (elementId === 'input-status' || elementId === 'input-reservation-type')) {
        match = dataSource.find(item => {
          // Extract code from name if it's in format "NAME [CODE]"
          const codeMatch = item.name.match(/\[([^\]]+)\]$/);
          if (codeMatch) {
            return normalizeText(codeMatch[1]) === searchTerm;
          }
          return false;
        });

        // Also try matching status name without code (e.g., "Pendiente de confirmación" -> "PENDIENTE DE CONFIRMACION [PC]")
        if (!match) {
          // First try to find exact match of name part
          match = dataSource.find(item => {
            const nameWithoutCode = item.name.replace(/\s*\[[^\]]+\]$/, '');
            return normalizeText(nameWithoutCode) === searchTerm;
          });
        }

        // Then try partial match but prefer longer matches (more specific)
        if (!match) {
          const candidates = dataSource.filter(item => {
            const nameWithoutCode = item.name.replace(/\s*\[[^\]]+\]$/, '');
            const normalizedName = normalizeText(nameWithoutCode);
            return normalizedName.includes(searchTerm) || searchTerm.includes(normalizedName);
          });

          // Sort by name length descending to prefer more specific matches
          // e.g., "PENDIENTE DE CONFIRMACION" over "CONFIRMACION"
          if (candidates.length > 0) {
            candidates.sort((a, b) => {
              const aName = a.name.replace(/\s*\[[^\]]+\]$/, '');
              const bName = b.name.replace(/\s*\[[^\]]+\]$/, '');
              // Prefer exact contains match
              const aContains = searchTerm.includes(normalizeText(aName));
              const bContains = searchTerm.includes(normalizeText(bName));
              if (aContains && !bContains) return -1;
              if (!aContains && bContains) return 1;
              // Then prefer longer names (more specific)
              return bName.length - aName.length;
            });
            match = candidates[0];
          }
        }
      }

      // 4. Try partial match - search term contains or is contained in name/code (normalized)
      // For status fields, this is handled above with better specificity
      if (!match && elementId !== 'input-status' && elementId !== 'input-reservation-type') {
        match = dataSource.find(item =>
          normalizeText(item.name).includes(searchTerm) ||
          searchTerm.includes(normalizeText(item.name)) ||
          normalizeText(item.code).includes(searchTerm)
        );
      }

      // 5. For sellers, try matching first part of name (e.g., "TEST" matches "TEST TEST")
      if (!match && elementId === 'input-seller') {
        match = dataSource.find(item => {
          const nameParts = normalizeText(item.name).split(' ');
          return nameParts.some(part => part === searchTerm || searchTerm.includes(part));
        });
      }

      // 6. For clients, try matching just the client name/code without the full format
      if (!match && elementId === 'input-client') {
        match = dataSource.find(item => {
          // Client format: "CODE - NAME - Cuit:XXX" or just "NAME"
          const parts = item.name.split(' - ');
          return parts.some(part =>
            normalizeText(part).includes(searchTerm) ||
            searchTerm.includes(normalizeText(part))
          );
        });
      }

      // 7. For currency, use alias mapping (USD -> DOLARES, EUR -> EUROS, etc.)
      if (!match && elementId === 'input-currency') {
        const currencyAliases: Record<string, string[]> = {
          'DOLARES': ['USD', 'DOLAR', 'US$', 'U$S', 'DOLLAR', 'DOLLARS', 'DÓLARES'],
          'PESOS': ['ARS', 'PESO', 'AR$', 'PESOS ARGENTINOS'],
          'EUROS': ['EUR', 'EURO', '€', 'EUROS'],
          'REALES': ['BRL', 'REAL', 'R$', 'REAIS'],
        };

        // Find which currency this alias belongs to
        for (const [currencyName, aliases] of Object.entries(currencyAliases)) {
          if (aliases.some(alias => normalizeText(alias) === searchTerm)) {
            // Find the currency in master data
            match = dataSource.find(item => normalizeText(item.name).includes(currencyName));
            if (match) {
              console.log(`💱 Currency alias found: "${valToSet}" -> "${match.name}"`);
              break;
            }
          }
        }
      }

      if (match) {
        // Set the display value
        element.value = match.name;

        // Also set hidden value if exists
        const hiddenInput = document.getElementById(`${elementId}-value`) as HTMLInputElement;
        if (hiddenInput) {
          // Always store the name (which includes the code for status/reservationType)
          // The RPA needs the full name to do hasText match in iTraffic dropdowns
          hiddenInput.value = match.name;
        }

        element.classList.remove('missing-field');
        console.log(`✅ Field ${elementId}: matched "${valToSet}" -> "${match.name}"`);
        return;
      } else {
        console.warn(`⚠️ Field ${elementId}: no match found for "${valToSet}" in ${dataSource.length} items`);
      }
    }
  }

  // Special handling for Select elements: Add option if it doesn't exist
  if (element.tagName === 'SELECT' && valToSet) {
    const select = element as HTMLSelectElement;
    let exists = false;
    // Check if value exists (case insensitive check might be better but strict for now)
    for (let i = 0; i < select.options.length; i++) {
      if (select.options[i].value === valToSet) {
        exists = true;
        break;
      }
    }

    if (!exists) {
      const option = document.createElement('option');
      option.value = valToSet;
      option.text = `${valToSet} (Detectado)`;
      select.add(option);
    }
  }

  element.value = valToSet;

  if (!valToSet) {
    element.classList.add('missing-field');
    // Only set placeholder for inputs, not selects
    if (element.tagName === 'INPUT') {
      (element as HTMLInputElement).placeholder = 'Requerido - Completar';
    }
    // Ensure it's enabled
    element.disabled = false;
  } else {
    element.classList.remove('missing-field');
  }
}

/**
 * Handle registration errors with user-friendly messages
 */
function handleRegistrationError(error: any): void {
  let userMessage = 'Registration failed. ';

  if (error.message.includes('offline_access')) {
    userMessage += 'The app needs permission to work in the background. Please contact your administrator.';
  } else if (error.message.includes('consent')) {
    userMessage += 'You must grant permissions to use this add-in.';
  } else if (error.message.includes('administrator')) {
    userMessage += 'This add-in requires administrator approval.';
  } else if (error.message.includes('network') || error.message.includes('fetch')) {
    userMessage += 'Network error. Please check your internet connection.';
  } else {
    userMessage += error.message || 'Please try again or contact support.';
  }

  showError(userMessage);
}

/**
 * Show error message
 */
function showError(message: string): void {
  state.error = message;
  const errorSection = document.getElementById('error-section')!;
  const errorMessage = document.getElementById('error-message')!;

  errorMessage.textContent = message;
  errorSection.style.display = 'flex';
}

/**
 * Hide error message
 */
function hideError(): void {
  state.error = null;
  document.getElementById('error-section')!.style.display = 'none';
}

/**
 * Show success message
 */
function showSuccessMessage(message: string): void {
  // You could implement a toast/notification here
  console.log('✅', message);

  // For now, temporarily update status box
  const statusBox = document.getElementById('status-indicator')!;
  const originalContent = statusBox.innerHTML;

  statusBox.innerHTML = `
    <div class="status-icon">✅</div>
    <div class="status-text">
      <h3>Success!</h3>
      <p>${message}</p>
    </div>
  `;

  statusBox.style.background = '#dff6dd';
  statusBox.style.borderColor = '#107c10';

  setTimeout(() => {
    statusBox.style.background = '';
    statusBox.style.borderColor = '';
  }, 3000);
}

/**
 * Set loading state
 */
function setLoading(loading: boolean): void {
  state.isLoading = loading;
  const button = document.getElementById('register-button') as HTMLButtonElement;

  if (button) {
    button.disabled = loading;
    button.innerHTML = loading
      ? '<span class="button-icon">⏳</span><span>Registering...</span>'
      : '<span class="button-icon">✨</span><span>Activate Auto-Classification</span>';
  }
}

/**
 * Show help dialog
 */
function showHelp(event: Event): void {
  event.preventDefault();

  Office.context.ui.displayDialogAsync(
    `${window.location.origin}/help.html`,
    { height: 60, width: 40 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('Failed to open help dialog:', result.error);
      }
    }
  );
}

// ============================================================================
// RESERVATION EXTRACTION FEATURE
// ============================================================================

/**
 * Handle extraction button click
 */
async function handleExtractReservation(): Promise<void> {
  console.log('🔍 handleExtractReservation called');
  console.log('   State:', { userId: state.userId, isRegistered: state.isRegistered, email: state.email });
  console.log('   Extraction visible:', document.getElementById('extraction-section')!.style.display);

  if (extractionState.isExtracting) {
    console.log('⚠️ Extraction already in progress');
    return;
  }

  // Ensure userId is available - check localStorage if not in state
  if (!state.userId) {
    const cachedUserId = localStorage.getItem('email-classifier-userId');
    const cachedEmail = localStorage.getItem('email-classifier-email');

    if (cachedUserId && cachedEmail) {
      console.log('♻️ Restored userId from localStorage');
      state.userId = cachedUserId;
      state.email = cachedEmail;
      state.isRegistered = true;
    } else {
      console.error('❌ No userId in state or localStorage!');
      showError('Please register first to use extraction feature');
      return;
    }
  }

  // Hide any previous errors
  document.getElementById('extraction-error')!.style.display = 'none';

  // Show loading
  document.getElementById('extraction-loading')!.style.display = 'block';
  document.getElementById('extraction-results')!.style.display = 'none';

  extractionState.isExtracting = true;
  extractionState.error = null;

  try {
    console.log('📧 Getting email content...');

    // Get email content from current item
    const emailContent = await getEmailConversationContent();

    if (!emailContent || emailContent.length < 50) {
      throw new Error('El contenido del email es demasiado corto o no se pudo obtener');
    }

    console.log(`✅ Email content obtained: ${emailContent.length} characters`);
    console.log('🚀 Sending to extraction API...');

    // Call backend extraction API
    const response = await fetch(`${config.apiBaseUrl}/api/extract-reservation`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        userId: state.userId,
        emailContent: emailContent,
        conversationId: Office.context.mailbox.item?.conversationId || null
      })
    });

    if (!response.ok) {
      const errorData: ExtractionApiError = await response.json().catch(() => ({
        success: false,
        error: 'Request failed',
        details: `HTTP ${response.status}: ${response.statusText}`
      }));

      throw new Error(errorData.details || errorData.error || 'Extraction failed');
    }

    const result: ExtractionApiResponse = await response.json();

    // Update state
    extractionState.hasExtracted = true;
    extractionState.data = result.data;
    extractionState.error = null;

    // Display results
    displayExtractionResults(result.data);

    // Hide loading
    document.getElementById('extraction-loading')!.style.display = 'none';

  } catch (error: any) {
    console.error('❌ Extraction error:', error);

    extractionState.error = error.message || 'Unknown error occurred';

    // Hide loading
    document.getElementById('extraction-loading')!.style.display = 'none';

    // Show error
    const errorSection = document.getElementById('extraction-error')!;
    const errorMessage = document.getElementById('extraction-error-message')!;

    errorMessage.textContent = error.message || 'No se pudo extraer la información. Por favor, intenta nuevamente.';
    errorSection.style.display = 'flex';

  } finally {
    extractionState.isExtracting = false;
  }
}

/**
 * Get email conversation content (including forwarded emails and replies)
 */
async function getEmailConversationContent(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;

    if (!item) {
      reject(new Error('No email item selected'));
      return;
    }

    console.log('📧 Getting email body (includes forwarded emails in body)...');

    // Get body as plain text (includes forwarded emails)
    // Office.CoercionType.Text preserves forwarded content and email chains
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const bodyContent = result.value || '';

        // Also get subject
        const subject = item.subject || '';

        // Get sender info
        const from = item.from?.displayName || item.from?.emailAddress || '';

        // Combine all information
        // The body content already includes forwarded emails, so we just need to add context
        const fullContent = `Asunto: ${subject}
De: ${from}

${bodyContent}`;

        console.log(`✅ Email content obtained: ${fullContent.length} characters`);
        console.log(`   Subject: ${subject.substring(0, 50)}...`);
        console.log(`   Body preview: ${bodyContent.substring(0, 100)}...`);

        resolve(fullContent);
      } else {
        console.error('❌ Failed to get email body:', result.error);
        reject(new Error('No se pudo obtener el contenido del email'));
      }
    });
  });
}

/**
 * Display extraction results in the UI
 */
function displayExtractionResults(data: ReservationData): void {
  console.log('📊 Displaying extraction results...');

  // Auto-collapse header to show results
  const header = document.getElementById('collapsible-header');
  const btn = document.getElementById('toggle-header-btn');
  if (header && btn) {
    header.style.display = 'none';
    btn.textContent = '🔽';
  }

  // Show results section
  document.getElementById('extraction-results')!.style.display = 'block';

  // Update confidence badge
  const confidenceBadge = document.getElementById('confidence-badge')!;
  const confidenceText = document.getElementById('confidence-text')!;
  const confidencePercent = Math.round((data.confidence || 0) * 100);

  confidenceText.textContent = `Confianza: ${confidencePercent}%`;

  // Set badge color based on confidence
  confidenceBadge.className = 'confidence-badge';
  if (confidencePercent >= 80) {
    // High confidence - green (default)
  } else if (confidencePercent >= 60) {
    confidenceBadge.classList.add('medium');
  } else {
    confidenceBadge.classList.add('low');
  }

  // Display passengers
  displayPassengers(data.passengers || []);

  // Calculate passenger counts
  const passengers = data.passengers || [];
  const adults = passengers.filter(p => p.paxType === 'ADU').length;
  const children = passengers.filter(p => p.paxType === 'CHD').length;
  const infants = passengers.filter(p => p.paxType === 'INF').length;

  // Set default dates
  const today = new Date().toISOString().split('T')[0];
  const travelDate = data.travelDate || data.checkIn || today;

  // Populate iTraffic fields (solo campos visibles)
  // Campos comentados en HTML - no intentar acceder
  // (document.getElementById('input-codigo') as HTMLInputElement).value = data.codigo || '';
  // (document.getElementById('input-estado-deuda') as HTMLInputElement).value = data.estadoDeuda || 'Aún no posee estado deudor';

  updateFieldStatus('input-reservation-type', data.reservationType || 'COMA'); // Default to Mayorista?
  updateFieldStatus('input-status', data.status || 'PC'); // Default to Pendiente

  // Campos de fecha comentados en HTML
  // (document.getElementById('input-reservation-date') as HTMLInputElement).value = data.reservationDate || today;
  (document.getElementById('input-travel-date') as HTMLInputElement).value = travelDate;
  // (document.getElementById('input-tour-end-date') as HTMLInputElement).value = data.tourEndDate || data.checkOut || '';
  // (document.getElementById('input-due-date') as HTMLInputElement).value = data.dueDate || today;

  updateFieldStatus('input-seller', data.seller);
  // (document.getElementById('input-passenger-name') as HTMLInputElement).value = data.passengerName || (passengers.length > 0 ? `${passengers[0].firstName} ${passengers[0].lastName}` : '');

  updateFieldStatus('input-client', data.client);

  // Campos comentados en HTML - no intentar acceder
  // const contactHiddenInput = document.getElementById('input-contact-value') as HTMLInputElement;
  // if (contactHiddenInput) contactHiddenInput.value = '';

  // (document.getElementById('input-commission') as HTMLInputElement).value = data.commission?.toString() || '';
  // updateFieldStatus('input-currency', data.currency || 'DOLARES');
  // (document.getElementById('input-exchange-rate') as HTMLInputElement).value = data.exchangeRate?.toString() || '1430.00';

  // (document.getElementById('input-adults') as HTMLInputElement).value = adults.toString();
  // (document.getElementById('input-children') as HTMLInputElement).value = children.toString();
  // (document.getElementById('input-infants') as HTMLInputElement).value = infants.toString();

  // (document.getElementById('input-trip-name') as HTMLInputElement).value = data.tripName || '';
  // (document.getElementById('input-product-code') as HTMLInputElement).value = data.productCode || data.reservationCode || '';

  // Secciones comentadas en HTML
  // if (data.flights && data.flights.length > 0) {
  //   displayFlights(data.flights);
  //   document.getElementById('flights-group')!.style.display = 'block';
  // } else {
  //   document.getElementById('flights-group')!.style.display = 'none';
  // }

  // if (data.services && data.services.length > 0) {
  //   displayServices(data.services);
  //   document.getElementById('services-group')!.style.display = 'block';
  // } else {
  //   document.getElementById('services-group')!.style.display = 'none';
  // }

  console.log('✅ Results displayed');
}

/**
 * Calculate age from date of birth
 */
function calculateAge(dateOfBirth: string): number | null {
  try {
    const birthDate = new Date(dateOfBirth);
    const today = new Date();
    let age = today.getFullYear() - birthDate.getFullYear();
    const monthDiff = today.getMonth() - birthDate.getMonth();
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
      age--;
    }
    return age;
  } catch {
    return null;
  }
}

/**
 * Display passengers list with editable inputs
 */
function displayPassengers(passengers: any[]): void {
  const passengersList = document.getElementById('passengers-list')!;
  const passengersCount = document.getElementById('passengers-count')!;

  passengersCount.textContent = passengers.length.toString();
  passengersList.innerHTML = '';

  if (passengers.length === 0) {
    // Add empty row for manual entry if needed
    passengers.push({});
  }

  // Build gender options
  const genderOptions = masterData.genders.length > 0
    ? masterData.genders.map(g => `<option value="${g.code}">${g.name}</option>`).join('')
    : `<option value="M">MASCULINO [M]</option><option value="F">FEMENINO [F]</option>`;

  // Build document type options
  const docTypeOptions = masterData.documentTypes.length > 0
    ? masterData.documentTypes.map(dt => `<option value="${dt.code}">${dt.name}</option>`).join('')
    : `<option value="PAS">PASAPORTE [PAS]</option><option value="DNI">DOCUMENTO NACIONAL DE IDENTIDAD [DNI]</option>`;

  // Build nationality options (countries)
  const nationalityOptions = masterData.countries.length > 0
    ? masterData.countries.map(c =>
      `<option value="${c.name}">${c.name}</option>`
    ).join('')
    : `<option value="ARGENTINA">ARGENTINA</option>
       <option value="BRASIL">BRASIL</option>
       <option value="ESTADOS UNIDOS">ESTADOS UNIDOS</option>`;

  passengers.forEach((passenger, index) => {
    const card = document.createElement('div');
    card.className = 'passenger-card';
    card.setAttribute('data-index', index.toString());

    // Set selected for current values
    const currentSex = passenger.sex || '';
    const currentDocType = passenger.documentType || '';
    const currentNationality = passenger.nationality || '';

    card.innerHTML = `
      <div class="passenger-card-header">
        <h4>Pasajero ${index + 1}</h4>
        <button class="icon-button remove-passenger" data-index="${index}" title="Eliminar">🗑️</button>
      </div>
      <div class="form-grid">
        <div class="form-group">
          <label>Nombre</label>
          <input type="text" class="passenger-input" data-field="firstName" value="${passenger.firstName || ''}" placeholder="Nombre">
        </div>
        <div class="form-group">
          <label>Apellido</label>
          <input type="text" class="passenger-input" data-field="lastName" value="${passenger.lastName || ''}" placeholder="Apellido">
        </div>
        <div class="form-group">
          <label>Tipo Pax</label>
          <select class="passenger-input" data-field="paxType">
            <option value="ADU" ${passenger.paxType === 'ADU' ? 'selected' : ''}>Adulto</option>
            <option value="CHD" ${passenger.paxType === 'CHD' ? 'selected' : ''}>Niño</option>
            <option value="INF" ${passenger.paxType === 'INF' ? 'selected' : ''}>Infante</option>
          </select>
        </div>
        <div class="form-group">
          <label>Sexo</label>
          <select class="passenger-input passenger-sex" data-field="sex" data-current="${currentSex}">
            <option value="">Seleccionar...</option>
            ${genderOptions}
          </select>
        </div>
        <div class="form-group">
          <label>Fecha Nacimiento</label>
          <input type="date" class="passenger-input" data-field="birthDate" value="${passenger.birthDate || ''}">
        </div>
        <div class="form-group">
          <label>Tipo Doc.</label>
          <select class="passenger-input passenger-doctype" data-field="documentType" data-current="${currentDocType}">
            <option value="">Seleccionar...</option>
            ${docTypeOptions}
          </select>
        </div>
        <div class="form-group">
          <label>Nro. Doc.</label>
          <input type="text" class="passenger-input" data-field="documentNumber" value="${passenger.documentNumber || ''}" placeholder="Número">
        </div>
        <div class="form-group">
          <label>CUIT-CUIL</label>
          <input type="text" class="passenger-input" data-field="cuilCuit" value="${passenger.cuilCuit || ''}" placeholder="11111111">
        </div>
        <div class="form-group full-width">
          <label>Dirección</label>
          <input type="text" class="passenger-input" data-field="direccion" value="${passenger.direccion || ''}" placeholder="Dirección completa">
        </div>
        <div class="form-group">
          <label>Nacionalidad</label>
          <select class="passenger-input passenger-nationality" data-field="nationality" data-current="${currentNationality}">
            <option value="">Seleccionar...</option>
            ${nationalityOptions}
          </select>
        </div>
        <!-- <div class="form-group">
          <label>Teléfono</label>
          <input type="text" class="passenger-input" data-field="phoneNumber" value="${passenger.phoneNumber || ''}" placeholder="+54...">
        </div> -->
      </div>
    `;

    passengersList.appendChild(card);
  });

  // Set selected values for selects after adding to DOM
  document.querySelectorAll('.passenger-sex').forEach((select: any) => {
    const current = select.getAttribute('data-current');
    if (current) select.value = current;
  });
  document.querySelectorAll('.passenger-doctype').forEach((select: any) => {
    const current = select.getAttribute('data-current');
    if (current) select.value = current;
  });
  document.querySelectorAll('.passenger-nationality').forEach((select: any) => {
    const current = select.getAttribute('data-current');
    if (current) select.value = current;
  });

  // Add "Add Passenger" button if not exists
  if (!document.getElementById('add-passenger-btn')) {
    const addBtn = document.createElement('button');
    addBtn.id = 'add-passenger-btn';
    addBtn.className = 'secondary-button small-button';
    addBtn.innerHTML = '➕ Agregar Pasajero';
    addBtn.onclick = () => {
      const currentData = getPassengersFromInputs();
      currentData.push({ paxType: 'ADU' });
      displayPassengers(currentData);
    };
    passengersList.parentNode?.appendChild(addBtn);
  }

  // Add event listeners for remove buttons
  document.querySelectorAll('.remove-passenger').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt((e.target as HTMLElement).getAttribute('data-index') || '0');
      const currentData = getPassengersFromInputs();
      currentData.splice(idx, 1);
      displayPassengers(currentData);
    });
  });
}

/**
 * Test RPA connection on startup
 */
async function testRPAConnection(): Promise<void> {
  console.log('🔌 Testing RPA connection...');
  
  try {
    const result = await rpaService.testConnection();
    
    if (result.success) {
      console.log('✅ RPA connection test successful:', result);
    } else {
      console.warn('⚠️ RPA connection test failed:', result.message);
    }
  } catch (error: any) {
    console.error('❌ RPA connection test error:', error);
  }
}

/**
 * Helper to get passengers data from inputs
 */
function getPassengersFromInputs(): any[] {
  const passengers: any[] = [];
  document.querySelectorAll('.passenger-card').forEach(card => {
    const p: any = {};
    card.querySelectorAll('.passenger-input').forEach((input: any) => {
      p[input.getAttribute('data-field')] = input.value;
    });
    passengers.push(p);
  });
  return passengers;
}


/**
 * Display flights list with editable inputs
 */
function displayFlights(flights: any[]): void {
  const flightsList = document.getElementById('flights-list')!;
  const flightsCount = document.getElementById('flights-count')!;

  flightsCount.textContent = flights.length.toString();
  flightsList.innerHTML = '';

  if (flights.length === 0) {
    flights.push({});
  }

  flights.forEach((flight, index) => {
    const card = document.createElement('div');
    card.className = 'flight-card';
    card.setAttribute('data-index', index.toString());

    card.innerHTML = `
      <div class="flight-card-header">
        <h4>Vuelo ${index + 1}</h4>
        <button class="icon-button remove-flight" data-index="${index}" title="Eliminar">🗑️</button>
      </div>
      <div class="form-grid">
        <div class="form-group">
          <label>Origen</label>
          <input type="text" class="flight-input" data-field="origin" value="${flight.origin || ''}" placeholder="EZE">
        </div>
        <div class="form-group">
          <label>Destino</label>
          <input type="text" class="flight-input" data-field="destination" value="${flight.destination || ''}" placeholder="MIA">
        </div>
        <div class="form-group">
          <label>Aerolínea</label>
          <input type="text" class="flight-input" data-field="airline" value="${flight.airline || ''}" placeholder="AA">
        </div>
        <div class="form-group">
          <label>Nro. Vuelo</label>
          <input type="text" class="flight-input" data-field="flightNumber" value="${flight.flightNumber || ''}" placeholder="900">
        </div>
        <div class="form-group">
          <label>Fecha Salida</label>
          <input type="date" class="flight-input" data-field="departureDate" value="${flight.departureDate || ''}">
        </div>
        <div class="form-group">
          <label>Hora Salida</label>
          <input type="time" class="flight-input" data-field="departureTime" value="${flight.departureTime || ''}">
        </div>
        <div class="form-group">
          <label>Fecha Llegada</label>
          <input type="date" class="flight-input" data-field="arrivalDate" value="${flight.arrivalDate || ''}">
        </div>
        <div class="form-group">
          <label>Hora Llegada</label>
          <input type="time" class="flight-input" data-field="arrivalTime" value="${flight.arrivalTime || ''}">
        </div>
      </div>
    `;

    flightsList.appendChild(card);
  });

  // Add "Add Flight" button if not exists
  if (!document.getElementById('add-flight-btn')) {
    const addBtn = document.createElement('button');
    addBtn.id = 'add-flight-btn';
    addBtn.className = 'secondary-button small-button';
    addBtn.innerHTML = '➕ Agregar Vuelo';
    addBtn.onclick = () => {
      const currentData = getFlightsFromInputs();
      currentData.push({});
      displayFlights(currentData);
    };
    flightsList.parentNode?.appendChild(addBtn);
  }

  // Add event listeners for remove buttons
  document.querySelectorAll('.remove-flight').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt((e.target as HTMLElement).getAttribute('data-index') || '0');
      const currentData = getFlightsFromInputs();
      currentData.splice(idx, 1);
      displayFlights(currentData);
    });
  });
}

/**
 * Helper to get flights data from inputs
 */
function getFlightsFromInputs(): any[] {
  const flights: any[] = [];
  document.querySelectorAll('.flight-card').forEach(card => {
    const f: any = {};
    card.querySelectorAll('.flight-input').forEach((input: any) => {
      f[input.getAttribute('data-field')] = input.value;
    });
    flights.push(f);
  });
  return flights;
}


/**
 * Display services list with editable inputs
 */
function displayServices(services: any[]): void {
  const servicesList = document.getElementById('services-list')!;
  const servicesCount = document.getElementById('services-count')!;

  servicesCount.textContent = services.length.toString();
  servicesList.innerHTML = '';

  if (services.length === 0) {
    services.push({});
  }

  services.forEach((service, index) => {
    const card = document.createElement('div');
    card.className = 'service-card';
    card.setAttribute('data-index', index.toString());

    card.innerHTML = `
      <div class="service-card-header">
        <h4>Servicio ${index + 1}</h4>
        <button class="icon-button remove-service" data-index="${index}" title="Eliminar">🗑️</button>
      </div>
      <div class="form-grid">
        <div class="form-group">
          <label>Tipo</label>
          <select class="service-input" data-field="type">
            <option value="transfer" ${service.type === 'transfer' ? 'selected' : ''}>Traslado</option>
            <option value="excursion" ${service.type === 'excursion' ? 'selected' : ''}>Excursión</option>
            <option value="meal" ${service.type === 'meal' ? 'selected' : ''}>Comida</option>
            <option value="other" ${service.type === 'other' || !service.type ? 'selected' : ''}>Otro</option>
          </select>
        </div>
        <div class="form-group full-width">
          <label>Descripción</label>
          <input type="text" class="service-input" data-field="description" value="${service.description || ''}" placeholder="Descripción del servicio">
        </div>
        <div class="form-group">
          <label>Fecha</label>
          <input type="date" class="service-input" data-field="date" value="${service.date || ''}">
        </div>
        <div class="form-group">
          <label>Ubicación</label>
          <input type="text" class="service-input" data-field="location" value="${service.location || ''}" placeholder="Lugar">
        </div>
      </div>
    `;

    servicesList.appendChild(card);
  });

  // Add "Add Service" button if not exists
  if (!document.getElementById('add-service-btn')) {
    const addBtn = document.createElement('button');
    addBtn.id = 'add-service-btn';
    addBtn.className = 'secondary-button small-button';
    addBtn.innerHTML = '➕ Agregar Servicio';
    addBtn.onclick = () => {
      const currentData = getServicesFromInputs();
      currentData.push({});
      displayServices(currentData);
    };
    servicesList.parentNode?.appendChild(addBtn);
  }

  // Add event listeners for remove buttons
  document.querySelectorAll('.remove-service').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt((e.target as HTMLElement).getAttribute('data-index') || '0');
      const currentData = getServicesFromInputs();
      currentData.splice(idx, 1);
      displayServices(currentData);
    });
  });
}

/**
 * Helper to get services data from inputs
 */
function getServicesFromInputs(): any[] {
  const services: any[] = [];
  document.querySelectorAll('.service-card').forEach(card => {
    const s: any = {};
    card.querySelectorAll('.service-input').forEach((input: any) => {
      s[input.getAttribute('data-field')] = input.value;
    });
    services.push(s);
  });
  return services;
}


/**
 * Handle re-analyze button click
 */
async function handleReanalyze(): Promise<void> {
  console.log('🔄 Re-analyzing email...');

  // Simply call extract again
  await handleExtractReservation();
}

/**
 * Handle confirm button click
 */
async function handleConfirmReservation(): Promise<void> {
  console.log('✅ Confirming reservation data...');

  if (!extractionState.data) {
    showError('No hay datos para confirmar');
    return;
  }

  // Ensure userId is available - check localStorage if not in state
  if (!state.userId) {
    const cachedUserId = localStorage.getItem('email-classifier-userId');
    const cachedEmail = localStorage.getItem('email-classifier-email');

    if (cachedUserId && cachedEmail) {
      console.log('♻️ Restored userId from localStorage for confirmation');
      state.userId = cachedUserId;
      state.email = cachedEmail;
      state.isRegistered = true;
    } else {
      showError('Usuario no autenticado');
      return;
    }
  }

  // Collect current values from form (in case user edited them)
  // For searchable dropdowns, use hidden input value (code) when available
  const getFieldValue = (inputId: string, hiddenId?: string): string | null => {
    if (hiddenId) {
      const hiddenInput = document.getElementById(hiddenId) as HTMLInputElement;
      if (hiddenInput && hiddenInput.value) {
        return hiddenInput.value;
      }
    }
    const input = document.getElementById(inputId) as HTMLInputElement;
    return input?.value || null;
  };

  // Helper to ensure field values match iTraffic format
  // This re-validates and converts codes to full names if needed
  const normalizeFieldValue = (
    value: string | null,
    dataSource: Array<{ code: string; name: string }>,
    fieldName: string
  ): string | null => {
    if (!value || !dataSource || dataSource.length === 0) return value;

    const searchTerm = value.toUpperCase().trim();

    // If already in correct format (contains brackets for status/type), return as-is
    if (searchTerm.includes('[') && searchTerm.includes(']')) {
      return value;
    }

    // Try to find match
    let match = dataSource.find(item => item.code.toUpperCase() === searchTerm);

    if (!match) {
      match = dataSource.find(item => item.name.toUpperCase() === searchTerm);
    }

    if (!match) {
      // Check if code is inside name brackets
      match = dataSource.find(item => {
        const codeMatch = item.name.match(/\[([^\]]+)\]$/);
        return codeMatch && codeMatch[1].toUpperCase() === searchTerm;
      });
    }

    if (!match) {
      // Partial match
      match = dataSource.find(item =>
        item.name.toUpperCase().includes(searchTerm) ||
        item.code.toUpperCase().includes(searchTerm)
      );
    }

    if (match) {
      console.log(`✅ Normalized ${fieldName}: "${value}" -> "${match.name}"`);
      return match.name;
    }

    console.warn(`⚠️ Could not normalize ${fieldName}: "${value}"`);
    return value;
  };

  // Solo incluir campos que realmente usamos en el RPA
  const confirmedData = {
    // Campos usados en RPA
    reservationType: normalizeFieldValue(
      getFieldValue('input-reservation-type', 'input-reservation-type-value'),
      masterData.reservationTypes,
      'reservationType'
    ),
    status: normalizeFieldValue(
      getFieldValue('input-status', 'input-status-value'),
      masterData.statuses,
      'status'
    ),
    client: normalizeFieldValue(
      getFieldValue('input-client', 'input-client-value'),
      masterData.clients,
      'client'
    ),
    travelDate: (document.getElementById('input-travel-date') as HTMLInputElement).value || null,
    seller: normalizeFieldValue(
      getFieldValue('input-seller', 'input-seller-value'),
      masterData.sellers,
      'seller'
    ),
    passengers: getPassengersFromInputs()
  };

  console.log('📋 Confirmed data:', confirmedData);

  // Update confirm button to show loading
  const confirmButton = document.getElementById('confirm-button') as HTMLButtonElement;
  const originalHTML = confirmButton.innerHTML;

  confirmButton.innerHTML = '<span class="button-icon">🤖</span><span>Procesando RPA...</span>';
  confirmButton.disabled = true;

  try {
    // Llamar directamente al servicio RPA para crear la reserva
    console.log('🚀 Calling RPA service to create reservation...');
    const rpaResult = await rpaService.createReservation(confirmedData as ReservationData);

    console.log('✅ RPA Result:', rpaResult);

    // Success
    confirmButton.innerHTML = '<span class="button-icon">✅</span><span>¡Procesado!</span>';
    confirmButton.style.background = '#107c10';
    showSuccessMessage(`✅ Reserva creada exitosamente en iTraffic: ${rpaResult.message}`);

    // Reset UI after success
    setTimeout(() => {
      // Hide results
      document.getElementById('extraction-results')!.style.display = 'none';
      // Show header again
      const header = document.getElementById('collapsible-header');
      const btn = document.getElementById('toggle-header-btn');
      if (header && btn) {
        header.style.display = 'block';
        btn.textContent = '🔼';
      }
      // Clear data
      extractionState.data = null;
      extractionState.hasExtracted = false;
    }, 2000);

    // Reset button after 5 seconds
    setTimeout(() => {
      confirmButton.innerHTML = originalHTML;
      confirmButton.disabled = false;
      confirmButton.style.background = '';
    }, 5000);

  } catch (error: any) {
    console.error('❌ Error creating reservation via RPA:', error);

    // Reset button
    confirmButton.innerHTML = originalHTML;
    confirmButton.disabled = false;

    // Mostrar error más descriptivo
    let errorMessage = 'Error al crear la reserva en iTraffic';
    if (error.message) {
      errorMessage += ': ' + error.message;
    }
    
    showError(errorMessage);
  }
}

/**
 * Export for testing (if needed)
 */
(window as any).EmailClassifier = {
  getAccessToken,
  handleRegister,
  checkRegistrationStatus,
};
