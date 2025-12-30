/**
* Reservation Extraction Types
*
* Type definitions for the email reservation extraction feature
*/
export type PassengerType = 'ADU' | 'CHD' | 'INF';
export type ServiceType = 'transfer' | 'excursion' | 'meal' | 'other';
export interface Passenger {
    firstName: string | null;
    lastName: string | null;
    paxType: PassengerType | null;
    documentType: string | null;
    documentNumber: string | null;
    nationality: string | null;
    birthDate: string | null;
    sex?: string | null;
    cuilCuit?: string | null;
    direccion?: string | null;
}
export interface Flight {
    flightNumber: string | null;
    airline: string | null;
    origin: string | null;
    destination: string | null;
    departureDate: string | null;
    departureTime: string | null;
    arrivalDate: string | null;
    arrivalTime: string | null;
}
export interface Service {
    type: ServiceType | null;
    description: string | null;
    date: string | null;
    location: string | null;
}
export interface ReservationData {
    checkIn: string | null;
    confidence: number;
    passengers: Passenger[];
    reservationType: string | null;
    status: string | null;
    client: string | null;
    travelDate: string | null;
    seller: string | null;
}
export interface ExtractionMetadata {
    extractedAt: string;
    processingTimeMs: number;
    passengersFound: number;
    qualityScore: number;
}
export interface ExtractionApiResponse {
    success: true;
    data: ReservationData;
    metadata: ExtractionMetadata;
}
export interface ExtractionApiError {
    success: false;
    error: string;
    details?: string;
}
export interface ExtractionState {
    isExtracting: boolean;
    hasExtracted: boolean;
    data: ReservationData | null;
    error: string | null;
    isEditing: boolean;
}
//# sourceMappingURL=reservation.d.ts.map