/**
 * RPA Service
 * Handles communication with the iTraffic RPA Backend
 */

import { config } from '../config/config';
import type { ReservationData } from '../types/reservation';

// ============================================================================
// Types
// ============================================================================

interface RPAReservationRequest {
  reservationData: {
    passengers: Array<{
      lastName: string;
      firstName: string;
      paxType: string;           // "ADU", "CHD", "INF"
      birthDate: string;         // Formato: MM/DD/AAAA
      nationality: string;       // Ej: "ARGENTINA"
      sex: string;               // "M" o "F"
      documentNumber: string;
      documentType: string;      // Ej: "DNI", "PASSPORT"
      cuilCuit: string;
      direccion: string;
    }>;
    reservationType: string;     // Ej: "AGENCIAS [COAG]"
    status: string;              // Ej: "PENDIENTE DE CONFIRMACION [PC]"
    client: string;              // Ej: "DESPEGAR - TEST - 1"
    travelDate: string;          // Formato: MM/DD/AAAA
    seller: string;              // Ej: "TEST TEST"
  };
}

interface RPAReservationResponse {
  success: boolean;
  message: string;
  data?: any;
  timestamp: string;
}

interface RPAStatusResponse {
  success: boolean;
  status: string;
  message: string;
  timestamp: string;
}

interface RPAHealthResponse {
  status: string;
  timestamp: string;
  uptime: number;
}

// ============================================================================
// RPA Service Class
// ============================================================================

class RPAService {
  private readonly baseUrl: string;
  private readonly timeout: number = 60000; // 60 segundos para operaciones RPA

  constructor() {
    this.baseUrl = config.apiBaseUrl;
    console.log('🤖 RPA Service initialized with base URL:', this.baseUrl);
  }

  /**
   * Health check del servicio RPA
   */
  async healthCheck(): Promise<RPAHealthResponse> {
    console.log('🏥 Checking RPA health...');
    
    try {
      const response = await this.fetchWithTimeout(`${this.baseUrl}/health`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json'
        }
      }, 10000); // 10 segundos para health check

      if (!response.ok) {
        throw new Error(`Health check failed: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      console.log('✅ RPA health check successful:', data);
      return data;
    } catch (error: any) {
      console.error('❌ RPA health check failed:', error);
      throw new Error(`Health check failed: ${error.message}`);
    }
  }

  /**
   * Obtener estado del servicio RPA
   */
  async getStatus(): Promise<RPAStatusResponse> {
    console.log('📊 Getting RPA status...');
    
    try {
      const response = await this.fetchWithTimeout(`${this.baseUrl}/api/status`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json'
        }
      }, 10000);

      if (!response.ok) {
        throw new Error(`Status check failed: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      console.log('✅ RPA status retrieved:', data);
      return data;
    } catch (error: any) {
      console.error('❌ RPA status check failed:', error);
      throw new Error(`Status check failed: ${error.message}`);
    }
  }

  /**
   * Crear reserva en iTraffic mediante RPA
   */
  async createReservation(reservationData: ReservationData): Promise<RPAReservationResponse> {
    console.log('🤖 Creating reservation via RPA...');
    console.log('📋 Reservation data:', reservationData);

    try {
      // Transformar datos al formato esperado por el RPA
      const rpaRequest: RPAReservationRequest = {
        reservationData: {
          passengers: reservationData.passengers.map(p => ({
            lastName: p.lastName || '',
            firstName: p.firstName || '',
            paxType: p.paxType || 'ADU',
            birthDate: this.formatDateForRPA(p.birthDate),
            nationality: p.nationality || '',
            sex: p.sex || 'M',
            documentNumber: p.documentNumber || '',
            documentType: p.documentType || 'DNI',
            cuilCuit: p.cuilCuit || '',
            direccion: p.direccion || ''
          })),
          reservationType: reservationData.reservationType || '',
          status: reservationData.status || '',
          client: reservationData.client || '',
          travelDate: this.formatDateForRPA(reservationData.travelDate),
          seller: reservationData.seller || ''
        }
      };

      console.log('📤 Sending to RPA:', rpaRequest);

      const response = await this.fetchWithTimeout(`${this.baseUrl}/api/reservations`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(rpaRequest)
      }, this.timeout);

      const rawText = await response.text();
      console.log('📥 RPA raw response:', rawText);

      let result: RPAReservationResponse;
      try {
        result = rawText ? JSON.parse(rawText) : { success: false, message: 'Empty response', timestamp: new Date().toISOString() };
      } catch (parseError) {
        console.error('❌ Failed to parse RPA response:', parseError);
        throw new Error(`Invalid response format: ${rawText.substring(0, 100)}`);
      }

      if (!response.ok) {
        console.error('❌ RPA request failed:', response.status, result);
        throw new Error(result.message || `RPA request failed: ${response.status}`);
      }

      if (!result.success) {
        console.error('❌ RPA operation failed:', result);
        throw new Error(result.message || 'RPA operation failed');
      }

      console.log('✅ Reservation created successfully:', result);
      return result;
    } catch (error: any) {
      console.error('❌ Create reservation failed:', error);
      
      // Proporcionar mensajes de error más descriptivos
      if (error.name === 'AbortError') {
        throw new Error('La operación tardó demasiado tiempo (timeout de 60s). El RPA puede estar procesando la reserva.');
      }
      
      throw new Error(error.message || 'Failed to create reservation');
    }
  }

  /**
   * Test de conexión con el backend RPA
   */
  async testConnection(): Promise<{ success: boolean; message: string; details?: any }> {
    console.log('🔌 Testing RPA connection...');
    
    try {
      const response = await this.fetchWithTimeout(`${this.baseUrl}/api/test-connection`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          timestamp: new Date().toISOString(),
          source: 'outlook-addin'
        })
      }, 10000);

      if (!response.ok) {
        return {
          success: false,
          message: `Connection test failed: ${response.status} ${response.statusText}`
        };
      }

      const data = await response.json();
      console.log('✅ Connection test successful:', data);
      
      return {
        success: true,
        message: 'Connection successful',
        details: data
      };
    } catch (error: any) {
      console.error('❌ Connection test failed:', error);
      return {
        success: false,
        message: `Connection failed: ${error.message}`
      };
    }
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  /**
   * Fetch con timeout
   */
  private async fetchWithTimeout(url: string, options: RequestInit, timeout: number): Promise<Response> {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeout);

    try {
      const response = await fetch(url, {
        ...options,
        signal: controller.signal
      });
      clearTimeout(timeoutId);
      return response;
    } catch (error: any) {
      clearTimeout(timeoutId);
      if (error.name === 'AbortError') {
        throw new Error(`Request timeout after ${timeout}ms`);
      }
      throw error;
    }
  }

  /**
   * Formatear fecha para el RPA (MM/DD/AAAA)
   * Acepta varios formatos de entrada y convierte a MM/DD/AAAA
   */
  private formatDateForRPA(dateString: string | null): string {
    if (!dateString) return '';

    try {
      // Intentar parsear diferentes formatos
      let date: Date | null = null;

      // Formato ISO (YYYY-MM-DD o YYYY-MM-DDTHH:mm:ss)
      if (dateString.includes('-')) {
        date = new Date(dateString);
      }
      // Formato DD/MM/AAAA o MM/DD/AAAA
      else if (dateString.includes('/')) {
        const parts = dateString.split('/');
        if (parts.length === 3) {
          // Asumimos DD/MM/AAAA si el primer número es > 12
          if (parseInt(parts[0]) > 12) {
            // DD/MM/AAAA -> MM/DD/AAAA
            date = new Date(`${parts[1]}/${parts[0]}/${parts[2]}`);
          } else {
            // Ya está en MM/DD/AAAA o ambiguo
            date = new Date(dateString);
          }
        }
      }

      if (!date || isNaN(date.getTime())) {
        console.warn('⚠️ Invalid date format:', dateString);
        return dateString; // Retornar original si no se puede parsear
      }

      // Formatear a MM/DD/AAAA
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      const year = date.getFullYear();

      const formatted = `${month}/${day}/${year}`;
      console.log(`📅 Date formatted: ${dateString} -> ${formatted}`);
      return formatted;
    } catch (error) {
      console.error('❌ Error formatting date:', error);
      return dateString; // Retornar original en caso de error
    }
  }
}

// ============================================================================
// Export Singleton
// ============================================================================

export const rpaService = new RPAService();

