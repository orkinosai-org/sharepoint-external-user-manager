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
import { Stack, CommandBar, DetailsList, SelectionMode, Spinner, SpinnerSize } from '@fluentui/react';
import { MockTimesheetDataService } from '../services/MockTimesheetDataService';
import styles from './TimesheetManagement.module.scss';
var TimesheetManagement = function (props) {
    var _a = useState([]), timesheets = _a[0], setTimesheets = _a[1];
    var _b = useState(true), loading = _b[0], setLoading = _b[1];
    useEffect(function () {
        loadTimesheets();
    }, []);
    var loadTimesheets = function () { return __awaiter(void 0, void 0, void 0, function () {
        return __generator(this, function (_a) {
            setLoading(true);
            try {
                // Simulate API call
                setTimeout(function () {
                    var mockTimesheets = MockTimesheetDataService.getTimesheets();
                    setTimesheets(mockTimesheets);
                    setLoading(false);
                }, 1000);
            }
            catch (error) {
                console.error('Error loading timesheets:', error);
                setLoading(false);
            }
            return [2 /*return*/];
        });
    }); };
    var columns = [
        {
            key: 'employee',
            name: 'Employee',
            fieldName: 'employee',
            minWidth: 150,
            maxWidth: 200,
            isResizable: true
        },
        {
            key: 'weekEnding',
            name: 'Week Ending',
            fieldName: 'weekEnding',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            onRender: function (item) { return new Date(item.weekEnding).toLocaleDateString(); }
        },
        {
            key: 'totalHours',
            name: 'Total Hours',
            fieldName: 'totalHours',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true
        },
        {
            key: 'status',
            name: 'Status',
            fieldName: 'status',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true
        }
    ];
    var commandBarItems = [
        {
            key: 'newTimesheet',
            text: 'New Timesheet',
            iconProps: { iconName: 'Add' },
            onClick: function () { return alert('New Timesheet functionality will be implemented'); }
        },
        {
            key: 'approve',
            text: 'Approve',
            iconProps: { iconName: 'CheckMark' },
            onClick: function () { return alert('Approve functionality will be implemented'); }
        },
        {
            key: 'reject',
            text: 'Reject',
            iconProps: { iconName: 'Cancel' },
            onClick: function () { return alert('Reject functionality will be implemented'); }
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: loadTimesheets
        }
    ];
    return (React.createElement("div", { className: styles.timesheetManagement },
        React.createElement(Stack, { tokens: { childrenGap: 20 } },
            React.createElement("h2", null, "Timesheet Management"),
            React.createElement(CommandBar, { items: commandBarItems }),
            loading ? (React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading timesheets..." })) : (React.createElement(DetailsList, { items: timesheets, columns: columns, selectionMode: SelectionMode.multiple, setKey: "set", layoutMode: 0, isHeaderVisible: true })))));
};
export default TimesheetManagement;
//# sourceMappingURL=TimesheetManagement.js.map