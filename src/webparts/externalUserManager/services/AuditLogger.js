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
var AuditLogger = /** @class */ (function () {
    function AuditLogger(context) {
        this.context = context;
        this.sessionId = this.generateSessionId();
    }
    /**
     * Log informational events
     */
    AuditLogger.prototype.logInfo = function (action, details, metadata) {
        this.log('info', action, details, metadata);
    };
    /**
     * Log warning events
     */
    AuditLogger.prototype.logWarning = function (action, details, metadata) {
        this.log('warning', action, details, metadata);
    };
    /**
     * Log error events
     */
    AuditLogger.prototype.logError = function (action, details, error, metadata) {
        var errorDetails = error ? "".concat(details, " - Error: ").concat(error.message || error) : details;
        var errorMetadata = __assign(__assign({}, metadata), { error: error ? {
                message: error.message,
                stack: error.stack,
                name: error.name
            } : undefined });
        this.log('error', action, errorDetails, errorMetadata);
    };
    /**
     * Core logging method
     */
    AuditLogger.prototype.log = function (level, action, details, metadata) {
        var logEntry = {
            timestamp: new Date(),
            userId: this.context.pageContext.user.loginName,
            userDisplayName: this.context.pageContext.user.displayName,
            action: action,
            details: details,
            level: level,
            metadata: metadata,
            sessionId: this.sessionId
        };
        // Log to browser console for development
        this.logToConsole(logEntry);
        // In production, you would also log to:
        // - SharePoint list for audit trail
        // - Application Insights
        // - Azure Log Analytics
        // - Custom logging service
        this.logToSharePointList(logEntry);
    };
    /**
     * Log to browser console for development/debugging
     */
    AuditLogger.prototype.logToConsole = function (entry) {
        var message = "[".concat(entry.level.toUpperCase(), "] ").concat(entry.action, ": ").concat(entry.details);
        switch (entry.level) {
            case 'error':
                console.error(message, entry);
                break;
            case 'warning':
                console.warn(message, entry);
                break;
            default:
                console.log(message, entry);
                break;
        }
    };
    /**
     * Log to SharePoint list for persistent audit trail
     * Note: This requires a pre-configured audit log list
     */
    AuditLogger.prototype.logToSharePointList = function (entry) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                try {
                    // In a real implementation, you would:
                    // 1. Check if audit log list exists, create if needed
                    // 2. Add item to the list with proper error handling
                    // 3. Handle batch operations for performance
                    // For now, we'll simulate this
                    if (this.context.pageContext.web.absoluteUrl) {
                        // This would be the actual SharePoint list logging implementation
                        // await sp.web.lists.getByTitle("AuditLog").items.add({
                        //   Title: entry.action,
                        //   UserId: entry.userId,
                        //   UserDisplayName: entry.userDisplayName,
                        //   Details: entry.details,
                        //   Level: entry.level,
                        //   Metadata: JSON.stringify(entry.metadata),
                        //   SessionId: entry.sessionId,
                        //   Timestamp: entry.timestamp.toISOString()
                        // });
                    }
                }
                catch (error) {
                    // Don't let audit logging failures break the main functionality
                    console.error('Failed to log to SharePoint audit list:', error);
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Generate a unique session ID for tracking related operations
     */
    AuditLogger.prototype.generateSessionId = function () {
        return "".concat(Date.now(), "-").concat(Math.random().toString(36).substr(2, 9));
    };
    /**
     * Get all audit logs for a specific action or time period
     * This would be used by admin reporting features
     */
    AuditLogger.prototype.getAuditLogs = function (filter) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                try {
                    // In a real implementation, this would query the SharePoint audit log list
                    // with appropriate filters and return the results
                    // For now, return empty array as this is primarily for future functionality
                    return [2 /*return*/, []];
                }
                catch (error) {
                    console.error('Failed to retrieve audit logs:', error);
                    return [2 /*return*/, []];
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Export audit logs to CSV for compliance reporting
     */
    AuditLogger.prototype.exportAuditLogs = function (filter) {
        return __awaiter(this, void 0, void 0, function () {
            var logs, headers, csvRows_1, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getAuditLogs(filter)];
                    case 1:
                        logs = _a.sent();
                        headers = ['Timestamp', 'User', 'Action', 'Details', 'Level', 'SessionId'];
                        csvRows_1 = [headers.join(',')];
                        logs.forEach(function (log) {
                            var row = [
                                log.timestamp.toISOString(),
                                "\"".concat(log.userDisplayName, "\""),
                                "\"".concat(log.action, "\""),
                                "\"".concat(log.details.replace(/"/g, '""'), "\""),
                                log.level,
                                log.sessionId || ''
                            ];
                            csvRows_1.push(row.join(','));
                        });
                        return [2 /*return*/, csvRows_1.join('\n')];
                    case 2:
                        error_1 = _a.sent();
                        console.error('Failed to export audit logs:', error_1);
                        return [2 /*return*/, ''];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    return AuditLogger;
}());
export { AuditLogger };
//# sourceMappingURL=AuditLogger.js.map