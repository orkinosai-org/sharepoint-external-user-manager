var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var MockBookingDataService = /** @class */ (function () {
    function MockBookingDataService() {
    }
    MockBookingDataService.getMeetingRooms = function () {
        return [
            {
                id: '1',
                name: 'Executive Boardroom',
                location: 'Building A, Floor 10',
                capacity: 20,
                description: 'Premium boardroom with city views and high-end AV equipment',
                amenities: ['Video Conferencing', 'Whiteboard', 'Coffee Station', 'Projector'],
                isAvailable: true,
                floor: '10',
                building: 'Building A',
                equipment: [
                    {
                        id: 'eq1',
                        name: '4K Projector',
                        type: 'AV',
                        isAvailable: true,
                        requiresSetup: true,
                        setupTime: 15
                    },
                    {
                        id: 'eq2',
                        name: 'Conference Phone',
                        type: 'AV',
                        isAvailable: true,
                        requiresSetup: false,
                        setupTime: 0
                    }
                ],
                accessibility: {
                    wheelchairAccessible: true,
                    hearingLoopAvailable: true,
                    visualAidsSupport: true,
                    accessibleParking: true
                },
                hourlyRate: 150,
                bookingPolicy: {
                    maxBookingDuration: 8,
                    advanceBookingLimit: 30,
                    cancellationDeadline: 24,
                    requiresApproval: true,
                    allowRecurring: true
                }
            },
            {
                id: '2',
                name: 'Innovation Lab',
                location: 'Building B, Floor 3',
                capacity: 12,
                description: 'Modern collaboration space with interactive displays and flexible seating',
                amenities: ['Interactive Display', 'Moveable Furniture', 'High-speed WiFi', 'Phone Booth'],
                isAvailable: true,
                floor: '3',
                building: 'Building B',
                equipment: [
                    {
                        id: 'eq3',
                        name: 'Interactive Whiteboard',
                        type: 'AV',
                        isAvailable: true,
                        requiresSetup: false,
                        setupTime: 0
                    }
                ],
                accessibility: {
                    wheelchairAccessible: true,
                    hearingLoopAvailable: false,
                    visualAidsSupport: true,
                    accessibleParking: true
                },
                hourlyRate: 75,
                bookingPolicy: {
                    maxBookingDuration: 4,
                    advanceBookingLimit: 14,
                    cancellationDeadline: 4,
                    requiresApproval: false,
                    allowRecurring: true
                }
            },
            {
                id: '3',
                name: 'Team Huddle Room',
                location: 'Building A, Floor 5',
                capacity: 6,
                description: 'Cozy space perfect for small team meetings and brainstorming',
                amenities: ['TV Display', 'Whiteboard', 'Conference Phone'],
                isAvailable: true,
                floor: '5',
                building: 'Building A',
                equipment: [
                    {
                        id: 'eq4',
                        name: '55" TV Display',
                        type: 'AV',
                        isAvailable: true,
                        requiresSetup: false,
                        setupTime: 0
                    }
                ],
                accessibility: {
                    wheelchairAccessible: true,
                    hearingLoopAvailable: false,
                    visualAidsSupport: false,
                    accessibleParking: false
                },
                hourlyRate: 25,
                bookingPolicy: {
                    maxBookingDuration: 2,
                    advanceBookingLimit: 7,
                    cancellationDeadline: 2,
                    requiresApproval: false,
                    allowRecurring: true
                }
            },
            {
                id: '4',
                name: 'Training Center',
                location: 'Building C, Floor 1',
                capacity: 50,
                description: 'Large training facility with theater-style seating and full AV setup',
                amenities: ['Theater Seating', 'Sound System', 'Stage Area', 'Recording Equipment'],
                isAvailable: true,
                floor: '1',
                building: 'Building C',
                equipment: [
                    {
                        id: 'eq5',
                        name: 'Professional Sound System',
                        type: 'AV',
                        isAvailable: true,
                        requiresSetup: true,
                        setupTime: 30
                    },
                    {
                        id: 'eq6',
                        name: 'Recording Equipment',
                        type: 'AV',
                        isAvailable: true,
                        requiresSetup: true,
                        setupTime: 45
                    }
                ],
                accessibility: {
                    wheelchairAccessible: true,
                    hearingLoopAvailable: true,
                    visualAidsSupport: true,
                    accessibleParking: true
                },
                hourlyRate: 200,
                bookingPolicy: {
                    maxBookingDuration: 8,
                    advanceBookingLimit: 60,
                    cancellationDeadline: 48,
                    requiresApproval: true,
                    allowRecurring: false
                }
            }
        ];
    };
    MockBookingDataService.getBookings = function () {
        var today = new Date();
        var tomorrow = new Date(today);
        tomorrow.setDate(tomorrow.getDate() + 1);
        return [
            {
                id: '1',
                roomId: '1',
                roomName: 'Executive Boardroom',
                title: 'Board Meeting',
                description: 'Monthly board meeting with all executives',
                startTime: new Date(today.getFullYear(), today.getMonth(), today.getDate(), 9, 0),
                endTime: new Date(today.getFullYear(), today.getMonth(), today.getDate(), 11, 0),
                organizer: 'ceo@company.com',
                attendees: ['cfo@company.com', 'cto@company.com', 'coo@company.com'],
                status: 'Confirmed',
                isRecurring: true,
                teamsLink: 'https://teams.microsoft.com/l/meetup-join/...',
                equipmentBooked: ['eq1', 'eq2'],
                createdDate: new Date('2024-01-01'),
                lastModified: new Date('2024-01-01')
            },
            {
                id: '2',
                roomId: '2',
                roomName: 'Innovation Lab',
                title: 'Product Design Workshop',
                description: 'Collaborative design session for new product features',
                startTime: new Date(today.getFullYear(), today.getMonth(), today.getDate(), 14, 0),
                endTime: new Date(today.getFullYear(), today.getMonth(), today.getDate(), 16, 0),
                organizer: 'designer@company.com',
                attendees: ['pm@company.com', 'dev1@company.com', 'dev2@company.com'],
                status: 'Confirmed',
                isRecurring: false,
                equipmentBooked: ['eq3'],
                createdDate: new Date('2024-01-10'),
                lastModified: new Date('2024-01-10')
            },
            {
                id: '3',
                roomId: '3',
                roomName: 'Team Huddle Room',
                title: 'Daily Standup',
                description: 'Daily team synchronization meeting',
                startTime: new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate(), 9, 30),
                endTime: new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate(), 10, 0),
                organizer: 'scrum@company.com',
                attendees: ['dev1@company.com', 'dev2@company.com', 'qa@company.com'],
                status: 'Confirmed',
                isRecurring: true,
                recurrencePattern: {
                    type: 'Daily',
                    interval: 1,
                    daysOfWeek: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
                },
                equipmentBooked: ['eq4'],
                createdDate: new Date('2024-01-05'),
                lastModified: new Date('2024-01-05')
            }
        ];
    };
    MockBookingDataService.getAvailabilitySlots = function (roomId, date) {
        return __awaiter(this, void 0, void 0, function () {
            var slots, startHour, endHour, _loop_1, this_1, hour;
            return __generator(this, function (_a) {
                slots = [];
                startHour = 8;
                endHour = 18;
                _loop_1 = function (hour) {
                    var slotStart = new Date(date.getFullYear(), date.getMonth(), date.getDate(), hour, 0);
                    var slotEnd = new Date(date.getFullYear(), date.getMonth(), date.getDate(), hour + 1, 0);
                    // Check if slot conflicts with existing bookings
                    var bookings = this_1.getBookings();
                    var hasConflict = bookings.some(function (booking) {
                        if (booking.roomId !== roomId)
                            return false;
                        var bookingStart = new Date(booking.startTime);
                        var bookingEnd = new Date(booking.endTime);
                        return slotStart < bookingEnd && slotEnd > bookingStart;
                    });
                    slots.push({
                        start: slotStart,
                        end: slotEnd,
                        isAvailable: !hasConflict,
                        isRestricted: hour < 9 || hour > 17 // Business hours restriction
                    });
                };
                this_1 = this;
                for (hour = startHour; hour < endHour; hour++) {
                    _loop_1(hour);
                }
                return [2 /*return*/, slots];
            });
        });
    };
    MockBookingDataService.createBooking = function (booking) {
        return __awaiter(this, void 0, void 0, function () {
            var newBooking;
            return __generator(this, function (_a) {
                newBooking = {
                    id: Math.random().toString(36).substr(2, 9),
                    roomId: booking.roomId || '',
                    roomName: booking.roomName || '',
                    title: booking.title || '',
                    description: booking.description,
                    startTime: booking.startTime || new Date(),
                    endTime: booking.endTime || new Date(),
                    organizer: booking.organizer || 'user@company.com',
                    attendees: booking.attendees || [],
                    status: booking.status || 'Confirmed',
                    isRecurring: booking.isRecurring || false,
                    equipmentBooked: booking.equipmentBooked || [],
                    createdDate: new Date(),
                    lastModified: new Date(),
                    teamsLink: booking.teamsLink
                };
                console.log('Creating booking:', newBooking);
                return [2 /*return*/, newBooking];
            });
        });
    };
    MockBookingDataService.updateBooking = function (id, updates) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                console.log("Updating booking ".concat(id), updates);
                return [2 /*return*/];
            });
        });
    };
    MockBookingDataService.cancelBooking = function (id, reason) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                console.log("Cancelling booking ".concat(id), reason);
                return [2 /*return*/];
            });
        });
    };
    MockBookingDataService.getMyBookings = function (userId) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would filter by user
                return [2 /*return*/, this.getBookings().filter(function (booking) { return booking.organizer === userId; })];
            });
        });
    };
    MockBookingDataService.getResourceUtilization = function (resourceId, startDate, endDate) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would calculate utilization metrics
                return [2 /*return*/, {
                        totalHours: 40,
                        bookedHours: 28,
                        utilizationRate: 0.7,
                        peakHours: ['9:00-10:00', '14:00-15:00'],
                        averageBookingDuration: 1.5
                    }];
            });
        });
    };
    return MockBookingDataService;
}());
export { MockBookingDataService };
//# sourceMappingURL=MockBookingDataService.js.map