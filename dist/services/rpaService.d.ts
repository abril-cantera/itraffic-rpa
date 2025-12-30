/**
 * RPA Service
 * Handles communication with the iTraffic RPA Backend
 */
import type { ReservationData } from '../types/reservation';
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
declare class RPAService {
    private readonly baseUrl;
    private readonly timeout;
    constructor();
    /**
     * Health check del servicio RPA
     */
    healthCheck(): Promise<RPAHealthResponse>;
    /**
     * Obtener estado del servicio RPA
     */
    getStatus(): Promise<RPAStatusResponse>;
    /**
     * Crear reserva en iTraffic mediante RPA
     */
    createReservation(reservationData: ReservationData): Promise<RPAReservationResponse>;
    /**
     * Test de conexión con el backend RPA
     */
    testConnection(): Promise<{
        success: boolean;
        message: string;
        details?: any;
    }>;
    /**
     * Fetch con timeout
     */
    private fetchWithTimeout;
    /**
     * Formatear fecha para el RPA (MM/DD/AAAA)
     * Acepta varios formatos de entrada y convierte a MM/DD/AAAA
     */
    private formatDateForRPA;
}
export declare const rpaService: RPAService;
export {};
//# sourceMappingURL=rpaService.d.ts.map