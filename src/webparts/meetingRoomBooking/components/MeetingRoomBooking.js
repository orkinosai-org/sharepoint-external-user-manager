var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
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
import * as React from 'react';
import { useState, useEffect } from 'react';
import { Stack, CommandBar, DetailsList, SelectionMode, Spinner, SpinnerSize, DatePicker, Dropdown, MessageBar, MessageBarType, Panel, PanelType, TextField, PrimaryButton, DefaultButton, Text, Icon, Toggle } from '@fluentui/react';
import { MockBookingDataService } from '../services/MockBookingDataService';
import { TeamsIntegrationService } from '../services/TeamsIntegrationService';
import styles from './MeetingRoomBooking.module.scss';
var MeetingRoomBooking = function (props) {
    var _a = useState([]), rooms = _a[0], setRooms = _a[1];
    var _b = useState([]), bookings = _b[0], setBookings = _b[1];
    var _c = useState(new Date()), selectedDate = _c[0], setSelectedDate = _c[1];
    var _d = useState(''), selectedRoom = _d[0], setSelectedRoom = _d[1];
    var _e = useState(true), loading = _e[0], setLoading = _e[1];
    var _f = useState(false), showBookingPanel = _f[0], setShowBookingPanel = _f[1];
    var _g = useState({}), newBooking = _g[0], setNewBooking = _g[1];
    var _h = useState([]), availabilitySlots = _h[0], setAvailabilitySlots = _h[1];
    useEffect(function () {
        loadRoomsAndBookings();
    }, []);
    useEffect(function () {
        if (selectedDate && selectedRoom) {
            loadAvailability();
        }
    }, [selectedDate, selectedRoom]);
    var loadRoomsAndBookings = function () { return __awaiter(void 0, void 0, void 0, function () {
        return __generator(this, function (_a) {
            setLoading(true);
            try {
                setTimeout(function () {
                    var mockRooms = MockBookingDataService.getMeetingRooms();
                    var mockBookings = MockBookingDataService.getBookings();
                    setRooms(mockRooms);
                    setBookings(mockBookings);
                    setLoading(false);
                }, 1000);
            }
            catch (error) {
                console.error('Error loading rooms and bookings:', error);
                setLoading(false);
            }
            return [2 /*return*/];
        });
    }); };
    var loadAvailability = function () { return __awaiter(void 0, void 0, void 0, function () {
        var slots, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, MockBookingDataService.getAvailabilitySlots(selectedRoom, selectedDate)];
                case 1:
                    slots = _a.sent();
                    setAvailabilitySlots(slots);
                    return [3 /*break*/, 3];
                case 2:
                    error_1 = _a.sent();
                    console.error('Error loading availability:', error_1);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleCreateBooking = function () { return __awaiter(void 0, void 0, void 0, function () {
        var teamsLink, teamsIntegration, teamsError_1, booking, error_2;
        var _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _b.trys.push([0, 6, , 7]);
                    teamsLink = void 0;
                    if (!(props.enableTeamsIntegration && newBooking.title)) return [3 /*break*/, 4];
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, TeamsIntegrationService.createTeamsMeeting({
                            subject: newBooking.title,
                            startTime: newBooking.startTime,
                            endTime: newBooking.endTime,
                            attendees: newBooking.attendees || []
                        })];
                case 2:
                    teamsIntegration = _b.sent();
                    teamsLink = teamsIntegration.joinUrl;
                    return [3 /*break*/, 4];
                case 3:
                    teamsError_1 = _b.sent();
                    console.error('Teams integration error:', teamsError_1);
                    return [3 /*break*/, 4];
                case 4:
                    booking = __assign(__assign({}, newBooking), { roomId: selectedRoom, roomName: ((_a = rooms.find(function (r) { return r.id === selectedRoom; })) === null || _a === void 0 ? void 0 : _a.name) || '', status: props.requireApproval ? 'Requested' : 'Confirmed', createdDate: new Date(), lastModified: new Date(), organizer: props.context.pageContext.user.email, teamsLink: teamsLink });
                    return [4 /*yield*/, MockBookingDataService.createBooking(booking)];
                case 5:
                    _b.sent();
                    setShowBookingPanel(false);
                    setNewBooking({});
                    loadRoomsAndBookings();
                    return [3 /*break*/, 7];
                case 6:
                    error_2 = _b.sent();
                    console.error('Error creating booking:', error_2);
                    return [3 /*break*/, 7];
                case 7: return [2 /*return*/];
            }
        });
    }); };
    var roomOptions = rooms.map(function (room) { return ({
        key: room.id,
        text: "".concat(room.name, " (").concat(room.location, ") - Capacity: ").concat(room.capacity)
    }); });
    var todaysBookings = bookings.filter(function (booking) {
        var bookingDate = new Date(booking.startTime);
        return bookingDate.toDateString() === selectedDate.toDateString();
    });
    var selectedRoomBookings = todaysBookings.filter(function (booking) { return booking.roomId === selectedRoom; });
    var columns = [
        {
            key: 'title',
            name: 'Meeting Title',
            fieldName: 'title',
            minWidth: 150,
            maxWidth: 250,
            isResizable: true
        },
        {
            key: 'time',
            name: 'Time',
            fieldName: 'startTime',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            onRender: function (item) {
                return "".concat(new Date(item.startTime).toLocaleTimeString(), " - ").concat(new Date(item.endTime).toLocaleTimeString());
            }
        },
        {
            key: 'organizer',
            name: 'Organizer',
            fieldName: 'organizer',
            minWidth: 120,
            maxWidth: 180,
            isResizable: true
        },
        {
            key: 'status',
            name: 'Status',
            fieldName: 'status',
            minWidth: 80,
            maxWidth: 100,
            isResizable: true
        },
        {
            key: 'teamsLink',
            name: 'Teams',
            fieldName: 'teamsLink',
            minWidth: 60,
            maxWidth: 80,
            isResizable: true,
            onRender: function (item) {
                return item.teamsLink ? React.createElement(Icon, { iconName: "TeamsLogo", style: { color: '#464EB8' } }) : null;
            }
        }
    ];
    var commandBarItems = [
        {
            key: 'newBooking',
            text: 'New Booking',
            iconProps: { iconName: 'Add' },
            onClick: function () { return setShowBookingPanel(true); }
        },
        {
            key: 'viewCalendar',
            text: 'Calendar View',
            iconProps: { iconName: 'Calendar' },
            onClick: function () { return alert('Calendar view will be implemented'); }
        },
        {
            key: 'myBookings',
            text: 'My Bookings',
            iconProps: { iconName: 'Contact' },
            onClick: function () { return alert('My bookings view will be implemented'); }
        },
        {
            key: 'reports',
            text: 'Reports',
            iconProps: { iconName: 'BarChart4' },
            onClick: function () { return alert('Reports functionality will be implemented'); }
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: loadRoomsAndBookings
        }
    ];
    var renderAvailabilitySlots = function () { return (React.createElement(Stack, { tokens: { childrenGap: 10 } },
        React.createElement(Text, { variant: "mediumPlus" }, "Available Time Slots"),
        availabilitySlots.length === 0 ? (React.createElement(Text, { variant: "medium" }, "No availability data for selected room and date.")) : (availabilitySlots.map(function (slot, index) { return (React.createElement("div", { key: index, className: slot.isAvailable ? styles.availableSlot : styles.unavailableSlot, onClick: function () {
                if (slot.isAvailable) {
                    setNewBooking(__assign(__assign({}, newBooking), { startTime: slot.start, endTime: slot.end }));
                    setShowBookingPanel(true);
                }
            } },
            React.createElement(Text, { variant: "medium" },
                slot.start.toLocaleTimeString(),
                " - ",
                slot.end.toLocaleTimeString(),
                slot.isAvailable ? ' (Available)' : ' (Booked)'))); })))); };
    var renderBookingPanel = function () { return (React.createElement(Panel, { isOpen: showBookingPanel, type: PanelType.medium, onDismiss: function () { return setShowBookingPanel(false); }, headerText: "Create New Booking" },
        React.createElement(Stack, { tokens: { childrenGap: 15 } },
            React.createElement(TextField, { label: "Meeting Title", value: newBooking.title || '', onChange: function (_, value) { return setNewBooking(__assign(__assign({}, newBooking), { title: value })); }, required: true }),
            React.createElement(TextField, { label: "Description", multiline: true, rows: 3, value: newBooking.description || '', onChange: function (_, value) { return setNewBooking(__assign(__assign({}, newBooking), { description: value })); } }),
            React.createElement(Dropdown, { label: "Meeting Room", options: roomOptions, selectedKey: selectedRoom, onChange: function (_, option) { return setSelectedRoom(option === null || option === void 0 ? void 0 : option.key); }, required: true }),
            React.createElement(DatePicker, { label: "Date", value: selectedDate, onSelectDate: function (date) { return date && setSelectedDate(date); } }),
            React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 } },
                React.createElement(TextField, { label: "Start Time", type: "time", value: newBooking.startTime ?
                        new Date(newBooking.startTime).toTimeString().slice(0, 5) : '', onChange: function (_, value) {
                        if (value) {
                            var _a = value.split(':'), hours = _a[0], minutes = _a[1];
                            var startTime = new Date(selectedDate);
                            startTime.setHours(parseInt(hours), parseInt(minutes));
                            setNewBooking(__assign(__assign({}, newBooking), { startTime: startTime }));
                        }
                    } }),
                React.createElement(TextField, { label: "End Time", type: "time", value: newBooking.endTime ?
                        new Date(newBooking.endTime).toTimeString().slice(0, 5) : '', onChange: function (_, value) {
                        if (value) {
                            var _a = value.split(':'), hours = _a[0], minutes = _a[1];
                            var endTime = new Date(selectedDate);
                            endTime.setHours(parseInt(hours), parseInt(minutes));
                            setNewBooking(__assign(__assign({}, newBooking), { endTime: endTime }));
                        }
                    } })),
            props.enableTeamsIntegration && (React.createElement(Toggle, { label: "Create Teams Meeting", checked: true, onText: "Teams meeting will be created", offText: "No Teams meeting" })),
            React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 } },
                React.createElement(PrimaryButton, { text: "Book Room", onClick: handleCreateBooking }),
                React.createElement(DefaultButton, { text: "Cancel", onClick: function () { return setShowBookingPanel(false); } }))))); };
    return (React.createElement("div", { className: styles.meetingRoomBooking },
        React.createElement(Stack, { tokens: { childrenGap: 20 } },
            React.createElement("h2", null, "Meeting Room & Resource Booking"),
            props.requireApproval && (React.createElement(MessageBar, { messageBarType: MessageBarType.info }, "Bookings require approval. You will receive a notification once your booking is reviewed.")),
            React.createElement(CommandBar, { items: commandBarItems }),
            React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 20 } },
                React.createElement(Stack, { tokens: { childrenGap: 15 }, styles: { root: { width: '300px' } } },
                    React.createElement(DatePicker, { label: "Select Date", value: selectedDate, onSelectDate: function (date) { return date && setSelectedDate(date); } }),
                    React.createElement(Dropdown, { label: "Select Room", placeholder: "Choose a meeting room", options: roomOptions, selectedKey: selectedRoom, onChange: function (_, option) { return setSelectedRoom(option === null || option === void 0 ? void 0 : option.key); } }),
                    selectedRoom && renderAvailabilitySlots()),
                React.createElement(Stack, { tokens: { childrenGap: 15 }, styles: { root: { flex: 1 } } },
                    React.createElement(Text, { variant: "xLarge" },
                        "Bookings for ",
                        selectedDate.toLocaleDateString()),
                    loading ? (React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading bookings..." })) : (React.createElement(DetailsList, { items: selectedRoomBookings, columns: columns, selectionMode: SelectionMode.none, setKey: "set", layoutMode: 0, isHeaderVisible: true })))),
            renderBookingPanel())));
};
export default MeetingRoomBooking;
//# sourceMappingURL=MeetingRoomBooking.js.map