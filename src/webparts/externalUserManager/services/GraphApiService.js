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
import { AuditLogger } from './AuditLogger';
/**
 * Microsoft Graph API Service for SharePoint library management
 *
 * Used as fallback when PnPjs doesn't support specific operations:
 * - Tenant-level operations
 * - Cross-site operations requiring elevated permissions
 * - Advanced user management scenarios
 * - External sharing configuration
 */
var GraphApiService = /** @class */ (function () {
    function GraphApiService(context) {
        this.context = context;
        this.auditLogger = new AuditLogger(context);
    }
    /**
     * Get Microsoft Graph client
     */
    GraphApiService.prototype.getGraphClient = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    /**
     * Enable external sharing for a site using Graph API
     * This requires tenant admin permissions
     */
    GraphApiService.prototype.enableExternalSharingForSite = function (siteId) {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, siteUpdateRequest, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        this.auditLogger.logInfo('enableExternalSharingForSite', "Enabling external sharing for site: ".concat(siteId));
                        return [4 /*yield*/, this.getGraphClient()];
                    case 1:
                        graphClient = _a.sent();
                        siteUpdateRequest = {
                            sharingCapabilities: 'ExternalUserAndGuestSharing',
                            allowDownload: true,
                            allowSharing: true
                        };
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId))
                                .patch(siteUpdateRequest)];
                    case 2:
                        _a.sent();
                        this.auditLogger.logInfo('enableExternalSharingForSite', "Successfully enabled external sharing for site: ".concat(siteId));
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        this.auditLogger.logError('enableExternalSharingForSite', "Failed to enable external sharing for site: ".concat(siteId), error_1);
                        throw new Error("Failed to enable external sharing: ".concat(error_1.message));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get external users for a specific site
     * Uses Graph API to get more detailed user information
     */
    GraphApiService.prototype.getExternalUsersForSite = function (siteId) {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, response, externalUsers, error_2;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        this.auditLogger.logInfo('getExternalUsersForSite', "Fetching external users for site: ".concat(siteId));
                        return [4 /*yield*/, this.getGraphClient()];
                    case 1:
                        graphClient = _a.sent();
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId, "/permissions"))
                                .get()];
                    case 2:
                        response = _a.sent();
                        externalUsers = response.value.filter(function (permission) {
                            return permission.grantedToIdentities &&
                                permission.grantedToIdentities.some(function (identity) {
                                    return identity.user && identity.user.email &&
                                        (identity.user.email.includes('#ext#') || !identity.user.email.endsWith(_this.context.pageContext.aadInfo.tenantId));
                                });
                        });
                        this.auditLogger.logInfo('getExternalUsersForSite', "Found ".concat(externalUsers.length, " external users"));
                        return [2 /*return*/, externalUsers];
                    case 3:
                        error_2 = _a.sent();
                        this.auditLogger.logError('getExternalUsersForSite', "Failed to get external users for site: ".concat(siteId), error_2);
                        throw new Error("Failed to get external users: ".concat(error_2.message));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Create sharing link for a library
     * Uses Graph API for advanced sharing configuration
     */
    GraphApiService.prototype.createSharingLink = function (siteId, listId, options) {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, sharingRequest, response, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        this.auditLogger.logInfo('createSharingLink', "Creating sharing link for list: ".concat(listId));
                        return [4 /*yield*/, this.getGraphClient()];
                    case 1:
                        graphClient = _a.sent();
                        sharingRequest = __assign(__assign({ type: options.type, scope: options.scope }, (options.expirationDateTime && { expirationDateTime: options.expirationDateTime.toISOString() })), (options.password && { password: options.password }));
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId, "/lists/").concat(listId, "/createLink"))
                                .post(sharingRequest)];
                    case 2:
                        response = _a.sent();
                        this.auditLogger.logInfo('createSharingLink', "Successfully created sharing link for list: ".concat(listId));
                        return [2 /*return*/, response.link.webUrl];
                    case 3:
                        error_3 = _a.sent();
                        this.auditLogger.logError('createSharingLink', "Failed to create sharing link for list: ".concat(listId), error_3);
                        throw new Error("Failed to create sharing link: ".concat(error_3.message));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Invite external users to a library
     * Uses Graph API for advanced invitation features
     */
    GraphApiService.prototype.inviteExternalUsers = function (siteId, listId, invitations) {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, _i, invitations_1, invitation, inviteRequest, error_4, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 8, , 9]);
                        this.auditLogger.logInfo('inviteExternalUsers', "Inviting ".concat(invitations.length, " external users to list: ").concat(listId));
                        return [4 /*yield*/, this.getGraphClient()];
                    case 1:
                        graphClient = _a.sent();
                        _i = 0, invitations_1 = invitations;
                        _a.label = 2;
                    case 2:
                        if (!(_i < invitations_1.length)) return [3 /*break*/, 7];
                        invitation = invitations_1[_i];
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 5, , 6]);
                        inviteRequest = {
                            recipients: [__assign({ email: invitation.email }, (invitation.displayName && { displayName: invitation.displayName }))],
                            message: invitation.message || 'You have been invited to access a SharePoint library.',
                            requireSignIn: true,
                            sendInvitation: true,
                            roles: [invitation.role]
                        };
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId, "/lists/").concat(listId, "/invite"))
                                .post(inviteRequest)];
                    case 4:
                        _a.sent();
                        this.auditLogger.logInfo('inviteExternalUsers', "Successfully invited user: ".concat(invitation.email));
                        return [3 /*break*/, 6];
                    case 5:
                        error_4 = _a.sent();
                        this.auditLogger.logError('inviteExternalUsers', "Failed to invite user: ".concat(invitation.email), error_4);
                        return [3 /*break*/, 6];
                    case 6:
                        _i++;
                        return [3 /*break*/, 2];
                    case 7: return [3 /*break*/, 9];
                    case 8:
                        error_5 = _a.sent();
                        this.auditLogger.logError('inviteExternalUsers', "Failed to invite external users to list: ".concat(listId), error_5);
                        throw new Error("Failed to invite external users: ".concat(error_5.message));
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get site information including sharing settings
     */
    GraphApiService.prototype.getSiteInfo = function (siteId) {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, response, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        this.auditLogger.logInfo('getSiteInfo', "Fetching site information for: ".concat(siteId));
                        return [4 /*yield*/, this.getGraphClient()];
                    case 1:
                        graphClient = _a.sent();
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId))
                                .select('id,name,webUrl,sharingCapabilities,allowDownload,allowSharing')
                                .get()];
                    case 2:
                        response = _a.sent();
                        this.auditLogger.logInfo('getSiteInfo', "Successfully retrieved site information for: ".concat(siteId));
                        return [2 /*return*/, response];
                    case 3:
                        error_6 = _a.sent();
                        this.auditLogger.logError('getSiteInfo', "Failed to get site information for: ".concat(siteId), error_6);
                        throw new Error("Failed to get site information: ".concat(error_6.message));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Revoke external user access
     * Uses Graph API for advanced permission management
     */
    GraphApiService.prototype.revokeExternalUserAccess = function (siteId, permissionId) {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, error_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        this.auditLogger.logInfo('revokeExternalUserAccess', "Revoking access for permission: ".concat(permissionId));
                        return [4 /*yield*/, this.getGraphClient()];
                    case 1:
                        graphClient = _a.sent();
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId, "/permissions/").concat(permissionId))
                                .delete()];
                    case 2:
                        _a.sent();
                        this.auditLogger.logInfo('revokeExternalUserAccess', "Successfully revoked access for permission: ".concat(permissionId));
                        return [3 /*break*/, 4];
                    case 3:
                        error_7 = _a.sent();
                        this.auditLogger.logError('revokeExternalUserAccess', "Failed to revoke access for permission: ".concat(permissionId), error_7);
                        throw new Error("Failed to revoke access: ".concat(error_7.message));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get sharing analytics for a site
     * Provides insights into external sharing patterns
     */
    GraphApiService.prototype.getSharingAnalytics = function (siteId, days) {
        if (days === void 0) { days = 30; }
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, endDate, startDate, response, error_8;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        this.auditLogger.logInfo('getSharingAnalytics', "Fetching sharing analytics for site: ".concat(siteId));
                        return [4 /*yield*/, this.getGraphClient()];
                    case 1:
                        graphClient = _a.sent();
                        endDate = new Date();
                        startDate = new Date();
                        startDate.setDate(endDate.getDate() - days);
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId, "/analytics/allTime"))
                                .get()];
                    case 2:
                        response = _a.sent();
                        this.auditLogger.logInfo('getSharingAnalytics', "Successfully retrieved sharing analytics for site: ".concat(siteId));
                        return [2 /*return*/, response];
                    case 3:
                        error_8 = _a.sent();
                        this.auditLogger.logError('getSharingAnalytics', "Failed to get sharing analytics for site: ".concat(siteId), error_8);
                        throw new Error("Failed to get sharing analytics: ".concat(error_8.message));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    return GraphApiService;
}());
export { GraphApiService };
//# sourceMappingURL=GraphApiService.js.map