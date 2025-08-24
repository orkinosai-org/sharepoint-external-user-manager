export interface IMeetingRoom {
  id: string;
  name: string;
  location: string;
  capacity: number;
  description: string;
  amenities: string[];
  isAvailable: boolean;
  floor: string;
  building: string;
  imageUrl?: string;
  equipment: IEquipment[];
  accessibility: IAccessibility;
  hourlyRate?: number;
  bookingPolicy: IBookingPolicy;
}

export interface IEquipment {
  id: string;
  name: string;
  type: EquipmentType;
  isAvailable: boolean;
  requiresSetup: boolean;
  setupTime: number; // minutes
  description?: string;
}

export interface IAccessibility {
  wheelchairAccessible: boolean;
  hearingLoopAvailable: boolean;
  visualAidsSupport: boolean;
  accessibleParking: boolean;
}

export interface IBookingPolicy {
  maxBookingDuration: number; // hours
  advanceBookingLimit: number; // days
  cancellationDeadline: number; // hours
  requiresApproval: boolean;
  allowRecurring: boolean;
  restrictedHours?: ITimeRange[];
}

export interface ITimeRange {
  start: string; // HH:mm format
  end: string; // HH:mm format
}

export interface IBooking {
  id: string;
  roomId: string;
  roomName: string;
  title: string;
  description?: string;
  startTime: Date;
  endTime: Date;
  organizer: string;
  attendees: string[];
  status: BookingStatus;
  isRecurring: boolean;
  recurrencePattern?: IRecurrencePattern;
  teamsLink?: string;
  equipmentBooked: string[];
  catering?: ICateringOrder;
  approvedBy?: string;
  approvedDate?: Date;
  createdDate: Date;
  lastModified: Date;
}

export interface IRecurrencePattern {
  type: RecurrenceType;
  interval: number;
  endDate?: Date;
  occurrences?: number;
  daysOfWeek?: DayOfWeek[];
  dayOfMonth?: number;
}

export interface ICateringOrder {
  id: string;
  items: ICateringItem[];
  totalCost: number;
  deliveryTime: Date;
  specialInstructions?: string;
  status: CateringStatus;
}

export interface ICateringItem {
  id: string;
  name: string;
  quantity: number;
  unitPrice: number;
  category: CateringCategory;
}

export interface IResource {
  id: string;
  name: string;
  type: ResourceType;
  location: string;
  isAvailable: boolean;
  capacity?: number;
  hourlyRate?: number;
  description?: string;
  maintenanceSchedule?: IMaintenanceSchedule[];
}

export interface IMaintenanceSchedule {
  id: string;
  startDate: Date;
  endDate: Date;
  description: string;
  type: MaintenanceType;
}

export interface IAvailabilitySlot {
  start: Date;
  end: Date;
  isAvailable: boolean;
  isRestricted: boolean;
  restriction?: string;
}

export interface IBookingConflict {
  resourceId: string;
  resourceName: string;
  conflictingBooking: IBooking;
  suggestedAlternatives: IAvailabilitySlot[];
}

export interface ITeamsIntegration {
  meetingId?: string;
  joinUrl?: string;
  dialInNumbers?: string[];
  conferenceId?: string;
  isOnlineMeeting: boolean;
  autoCreateTeamsMeeting: boolean;
}

export type BookingStatus = 'Requested' | 'Approved' | 'Rejected' | 'Confirmed' | 'Completed' | 'Cancelled' | 'No Show';
export type RecurrenceType = 'Daily' | 'Weekly' | 'Monthly' | 'Yearly';
export type DayOfWeek = 'Monday' | 'Tuesday' | 'Wednesday' | 'Thursday' | 'Friday' | 'Saturday' | 'Sunday';
export type EquipmentType = 'AV' | 'Computing' | 'Furniture' | 'Catering' | 'Other';
export type ResourceType = 'Room' | 'Equipment' | 'Vehicle' | 'Facility' | 'Other';
export type CateringStatus = 'Ordered' | 'Confirmed' | 'Prepared' | 'Delivered' | 'Cancelled';
export type CateringCategory = 'Beverages' | 'Snacks' | 'Lunch' | 'Breakfast' | 'Dinner' | 'Other';
export type MaintenanceType = 'Routine' | 'Repair' | 'Upgrade' | 'Inspection' | 'Cleaning';