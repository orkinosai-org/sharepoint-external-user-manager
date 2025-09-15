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
var MockTimesheetDataService = /** @class */ (function () {
    function MockTimesheetDataService() {
    }
    MockTimesheetDataService.getTimesheets = function () {
        return [
            {
                id: '1',
                employee: 'John Smith',
                weekEnding: new Date('2024-01-12'),
                totalHours: 40,
                status: 'Approved',
                submittedDate: new Date('2024-01-13'),
                approvedBy: 'Jane Manager',
                approvedDate: new Date('2024-01-14'),
                entries: [
                    {
                        id: '1-1',
                        date: new Date('2024-01-08'),
                        project: 'Project Alpha',
                        task: 'Development',
                        hours: 8,
                        description: 'Frontend development'
                    },
                    {
                        id: '1-2',
                        date: new Date('2024-01-09'),
                        project: 'Project Alpha',
                        task: 'Testing',
                        hours: 8,
                        description: 'Unit testing'
                    }
                ]
            },
            {
                id: '2',
                employee: 'Sarah Johnson',
                weekEnding: new Date('2024-01-12'),
                totalHours: 35,
                status: 'Submitted',
                submittedDate: new Date('2024-01-13'),
                entries: [
                    {
                        id: '2-1',
                        date: new Date('2024-01-08'),
                        project: 'Project Beta',
                        task: 'Analysis',
                        hours: 7,
                        description: 'Requirements analysis'
                    }
                ]
            },
            {
                id: '3',
                employee: 'Mike Wilson',
                weekEnding: new Date('2024-01-12'),
                totalHours: 42,
                status: 'Draft',
                entries: [
                    {
                        id: '3-1',
                        date: new Date('2024-01-08'),
                        project: 'Project Gamma',
                        task: 'Design',
                        hours: 8,
                        description: 'UI/UX design'
                    }
                ]
            }
        ];
    };
    MockTimesheetDataService.createTimesheet = function (timesheet) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                return [2 /*return*/, {
                        id: Math.random().toString(36).substr(2, 9),
                        employee: timesheet.employee || '',
                        weekEnding: timesheet.weekEnding || new Date(),
                        totalHours: timesheet.totalHours || 0,
                        status: 'Draft',
                        entries: timesheet.entries || []
                    }];
            });
        });
    };
    MockTimesheetDataService.updateTimesheet = function (id, updates) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                console.log("Updating timesheet ".concat(id), updates);
                return [2 /*return*/];
            });
        });
    };
    MockTimesheetDataService.deleteTimesheet = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                console.log("Deleting timesheet ".concat(id));
                return [2 /*return*/];
            });
        });
    };
    return MockTimesheetDataService;
}());
export { MockTimesheetDataService };
//# sourceMappingURL=MockTimesheetDataService.js.map