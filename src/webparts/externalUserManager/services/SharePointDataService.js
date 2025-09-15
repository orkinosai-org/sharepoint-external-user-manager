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
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
import { spfi, SPFx } from "@pnp/sp";
import { AuditLogger } from './AuditLogger';
/**
 * SharePoint Data Service using PnPjs for library management
 *
 * Technical Decisions:
 * - PnPjs chosen for primary integration due to:
 *   * Better TypeScript support and intellisense
 *   * Simplified API compared to raw REST calls
 *   * Built-in error handling and retry logic
 *   * Better caching and performance optimizations
 *
 * - Microsoft Graph API used as fallback for:
 *   * Tenant-level operations not supported by PnPjs
 *   * Cross-site operations requiring elevated permissions
 *   * Advanced user management scenarios
 */
var SharePointDataService = /** @class */ (function () {
    function SharePointDataService(context) {
        this.context = context;
        this.auditLogger = new AuditLogger(context);
        // Initialize PnPjs with SPFx context
        this.sp = spfi().using(SPFx(context));
    }
    /**
     * Get all external libraries (document libraries with external sharing enabled)
     */
    SharePointDataService.prototype.getExternalLibraries = function () {
        return __awaiter(this, void 0, void 0, function () {
            var lists, libraries, _i, lists_1, list, hasExternalUsers, library, error_1, error_2;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 11, , 12]);
                        this.auditLogger.logInfo('getExternalLibraries', 'Fetching external libraries');
                        return [4 /*yield*/, this.sp.web.lists
                                .filter("BaseTemplate eq 101 and Hidden eq false") // Document libraries only
                                .select("Id", "Title", "Description", "DefaultViewUrl", "LastItemModifiedDate", "ItemCount", "Created")
                                .expand("RoleAssignments/Member")
                                .get()];
                    case 1:
                        lists = _b.sent();
                        libraries = [];
                        _i = 0, lists_1 = lists;
                        _b.label = 2;
                    case 2:
                        if (!(_i < lists_1.length)) return [3 /*break*/, 10];
                        list = lists_1[_i];
                        _b.label = 3;
                    case 3:
                        _b.trys.push([3, 8, , 9]);
                        return [4 /*yield*/, this.hasExternalUsers(list.Id)];
                    case 4:
                        hasExternalUsers = _b.sent();
                        if (!hasExternalUsers.hasExternal) return [3 /*break*/, 7];
                        _a = {
                            id: list.Id,
                            name: list.Title,
                            description: list.Description || 'No description available',
                            siteUrl: list.DefaultViewUrl || '',
                            externalUsersCount: hasExternalUsers.externalCount,
                            lastModified: new Date(list.LastItemModifiedDate)
                        };
                        return [4 /*yield*/, this.getLibraryOwner(list.Id)];
                    case 5:
                        _a.owner = _b.sent();
                        return [4 /*yield*/, this.getLibraryPermissionLevel(list.Id)];
                    case 6:
                        library = (_a.permissions = _b.sent(),
                            _a);
                        libraries.push(library);
                        _b.label = 7;
                    case 7: return [3 /*break*/, 9];
                    case 8:
                        error_1 = _b.sent();
                        this.auditLogger.logError('getExternalLibraries', "Error processing library ".concat(list.Title), error_1);
                        return [3 /*break*/, 9];
                    case 9:
                        _i++;
                        return [3 /*break*/, 2];
                    case 10:
                        this.auditLogger.logInfo('getExternalLibraries', "Retrieved ".concat(libraries.length, " external libraries"));
                        return [2 /*return*/, libraries];
                    case 11:
                        error_2 = _b.sent();
                        this.auditLogger.logError('getExternalLibraries', 'Failed to fetch external libraries', error_2);
                        throw new Error("Failed to fetch external libraries: ".concat(error_2.message));
                    case 12: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Create a new document library with specified configuration
     */
    SharePointDataService.prototype.createLibrary = function (libraryConfig) {
        return __awaiter(this, void 0, void 0, function () {
            var existingLibraries, libraryCreationInfo, createResult, newLibrary, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        this.auditLogger.logInfo('createLibrary', "Creating library: ".concat(libraryConfig.title));
                        // Validate input
                        if (!libraryConfig.title || libraryConfig.title.trim().length === 0) {
                            throw new Error('Library title is required');
                        }
                        return [4 /*yield*/, this.sp.web.lists
                                .filter("Title eq '".concat(libraryConfig.title.replace(/'/g, "''"), "'"))
                                .get()];
                    case 1:
                        existingLibraries = _a.sent();
                        if (existingLibraries.length > 0) {
                            throw new Error("A library with the name \"".concat(libraryConfig.title, "\" already exists"));
                        }
                        libraryCreationInfo = {
                            Title: libraryConfig.title,
                            Description: libraryConfig.description || '',
                            BaseTemplate: libraryConfig.template || 101,
                            EnableAttachments: false,
                            EnableFolderCreation: true
                        };
                        return [4 /*yield*/, this.sp.web.lists.add(libraryCreationInfo.Title, libraryCreationInfo.Description, libraryCreationInfo.BaseTemplate, false, // enableContentTypes
                            {
                                EnableAttachments: libraryCreationInfo.EnableAttachments,
                                EnableFolderCreation: libraryCreationInfo.EnableFolderCreation
                            })];
                    case 2:
                        createResult = _a.sent();
                        if (!libraryConfig.enableExternalSharing) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.enableExternalSharing(createResult.data.Id)];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4:
                        newLibrary = {
                            id: createResult.data.Id,
                            name: libraryConfig.title,
                            description: libraryConfig.description || '',
                            siteUrl: createResult.data.DefaultViewUrl || '',
                            externalUsersCount: 0,
                            lastModified: new Date(),
                            owner: this.context.pageContext.user.displayName,
                            permissions: 'Full Control'
                        };
                        this.auditLogger.logInfo('createLibrary', "Successfully created library: ".concat(libraryConfig.title), {
                            libraryId: createResult.data.Id,
                            externalSharing: libraryConfig.enableExternalSharing
                        });
                        return [2 /*return*/, newLibrary];
                    case 5:
                        error_3 = _a.sent();
                        this.auditLogger.logError('createLibrary', "Failed to create library: ".concat(libraryConfig.title), error_3);
                        throw new Error("Failed to create library: ".concat(error_3.message));
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Delete a document library by ID
     */
    SharePointDataService.prototype.deleteLibrary = function (libraryId) {
        return __awaiter(this, void 0, void 0, function () {
            var library, currentUserPermissions, error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        this.auditLogger.logInfo('deleteLibrary', "Deleting library: ".concat(libraryId));
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId)
                                .select("Title", "ItemCount")
                                .get()];
                    case 1:
                        library = _a.sent();
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId)
                                .getCurrentUserEffectivePermissions()];
                    case 2:
                        currentUserPermissions = _a.sent();
                        // Note: In a real implementation, you would check specific permissions
                        // For now, we assume the user has sufficient permissions
                        // Perform the deletion
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId).delete()];
                    case 3:
                        // Note: In a real implementation, you would check specific permissions
                        // For now, we assume the user has sufficient permissions
                        // Perform the deletion
                        _a.sent();
                        this.auditLogger.logInfo('deleteLibrary', "Successfully deleted library: ".concat(library.Title), {
                            libraryId: libraryId,
                            itemCount: library.ItemCount
                        });
                        return [3 /*break*/, 5];
                    case 4:
                        error_4 = _a.sent();
                        this.auditLogger.logError('deleteLibrary', "Failed to delete library: ".concat(libraryId), error_4);
                        throw new Error("Failed to delete library: ".concat(error_4.message));
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get external users for a specific library
     */
    SharePointDataService.prototype.getExternalUsersForLibrary = function (libraryId) {
        return __awaiter(this, void 0, void 0, function () {
            var roleAssignments, externalUsers, _i, roleAssignments_1, assignment, permissions, externalUser, metadata, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 6, , 7]);
                        this.auditLogger.logInfo('getExternalUsersForLibrary', "Fetching external users for library: ".concat(libraryId));
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId)
                                .roleAssignments
                                .expand("Member", "RoleDefinitionBindings")
                                .get()];
                    case 1:
                        roleAssignments = _a.sent();
                        externalUsers = [];
                        _i = 0, roleAssignments_1 = roleAssignments;
                        _a.label = 2;
                    case 2:
                        if (!(_i < roleAssignments_1.length)) return [3 /*break*/, 5];
                        assignment = roleAssignments_1[_i];
                        if (!(assignment.Member.LoginName && assignment.Member.LoginName.includes('#ext#'))) return [3 /*break*/, 4];
                        permissions = assignment.RoleDefinitionBindings
                            .map(function (role) { return role.Name; })
                            .join(', ');
                        externalUser = {
                            id: assignment.Member.Id.toString(),
                            email: assignment.Member.Email || '',
                            displayName: assignment.Member.Title || '',
                            invitedBy: 'Unknown',
                            invitedDate: new Date(),
                            lastAccess: new Date(),
                            permissions: this.mapPermissionLevel(permissions),
                            company: undefined,
                            project: undefined
                        };
                        return [4 /*yield*/, this.getUserMetadata(libraryId, assignment.Member.Id)];
                    case 3:
                        metadata = _a.sent();
                        if (metadata) {
                            externalUser.company = metadata.company;
                            externalUser.project = metadata.project;
                        }
                        externalUsers.push(externalUser);
                        _a.label = 4;
                    case 4:
                        _i++;
                        return [3 /*break*/, 2];
                    case 5:
                        this.auditLogger.logInfo('getExternalUsersForLibrary', "Found ".concat(externalUsers.length, " external users"));
                        return [2 /*return*/, externalUsers];
                    case 6:
                        error_5 = _a.sent();
                        this.auditLogger.logError('getExternalUsersForLibrary', "Failed to get external users for library: ".concat(libraryId), error_5);
                        throw new Error("Failed to get external users: ".concat(error_5.message));
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Add external user to a library with specified permissions
     */
    SharePointDataService.prototype.addExternalUserToLibrary = function (libraryId, email, permission, company, project) {
        return __awaiter(this, void 0, void 0, function () {
            var ensuredUser, userId, roleDefId, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 6, , 7]);
                        this.auditLogger.logInfo('addExternalUserToLibrary', "Adding external user ".concat(email, " to library: ").concat(libraryId, " with ").concat(permission, " permissions"), {
                            libraryId: libraryId,
                            email: email,
                            permission: permission,
                            company: company,
                            project: project
                        });
                        return [4 /*yield*/, this.sp.web.ensureUser(email)];
                    case 1:
                        ensuredUser = _a.sent();
                        userId = ensuredUser.data.Id;
                        if (!userId) {
                            throw new Error('Failed to get user ID after ensuring user');
                        }
                        return [4 /*yield*/, this.getRoleDefinitionId(permission)];
                    case 2:
                        roleDefId = _a.sent();
                        // Add role assignment to the library
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId)
                                .roleAssignments.add(userId, roleDefId)];
                    case 3:
                        // Add role assignment to the library
                        _a.sent();
                        if (!(company || project)) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.storeUserMetadata(libraryId, email, userId, company, project)];
                    case 4:
                        _a.sent();
                        _a.label = 5;
                    case 5:
                        this.auditLogger.logInfo('addExternalUserToLibrary', "Successfully added user ".concat(email, " to library with ").concat(permission, " permissions"), {
                            libraryId: libraryId,
                            email: email,
                            permission: permission,
                            userId: userId,
                            company: company,
                            project: project
                        });
                        return [3 /*break*/, 7];
                    case 6:
                        error_6 = _a.sent();
                        this.auditLogger.logError('addExternalUserToLibrary', "Failed to add external user ".concat(email, " to library: ").concat(libraryId), error_6);
                        throw new Error("Failed to add user: ".concat(error_6.message));
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    /**
  
     * Remove external user from a library
     */
    SharePointDataService.prototype.removeExternalUserFromLibrary = function (libraryId, userId) {
        return __awaiter(this, void 0, void 0, function () {
            var error_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        this.auditLogger.logInfo('removeExternalUserFromLibrary', "Removing external user ".concat(userId, " from library: ").concat(libraryId));
                        // Remove role assignment from the library
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId)
                                .roleAssignments.remove(parseInt(userId, 10))];
                    case 1:
                        // Remove role assignment from the library
                        _a.sent();
                        this.auditLogger.logInfo('removeExternalUserFromLibrary', "Successfully removed user ".concat(userId, " from library"), {
                            libraryId: libraryId,
                            userId: userId
                        });
                        return [3 /*break*/, 3];
                    case 2:
                        error_7 = _a.sent();
                        this.auditLogger.logError('removeExternalUserFromLibrary', "Failed to remove external user ".concat(userId, " from library: ").concat(libraryId), error_7);
                        throw new Error("Failed to remove user: ".concat(error_7.message));
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
  
     * Update user metadata (company and project) - public method
     */
    SharePointDataService.prototype.updateUserMetadata = function (libraryId, userId, company, project) {
        return __awaiter(this, void 0, void 0, function () {
            var userIdNum, userEmail, users, user, error_8, error_9;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 6, , 7]);
                        userIdNum = parseInt(userId, 10);
                        userEmail = 'Unknown';
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.getExternalUsersForLibrary(libraryId)];
                    case 2:
                        users = _a.sent();
                        user = users.find(function (u) { return u.id === userId; });
                        if (user) {
                            userEmail = user.email;
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        error_8 = _a.sent();
                        // Continue with metadata update even if we can't get email
                        this.auditLogger.logWarning('updateUserMetadata', 'Could not retrieve user email for audit', { userId: userId });
                        return [3 /*break*/, 4];
                    case 4: return [4 /*yield*/, this.storeUserMetadata(libraryId, userEmail, userIdNum, company, project)];
                    case 5:
                        _a.sent();
                        this.auditLogger.logInfo('updateUserMetadata', "Successfully updated metadata for user ".concat(userId), {
                            libraryId: libraryId,
                            userId: userId,
                            company: company,
                            project: project
                        });
                        return [3 /*break*/, 7];
                    case 6:
                        error_9 = _a.sent();
                        this.auditLogger.logError('updateUserMetadata', "Failed to update metadata for user ".concat(userId), error_9);
                        throw new Error("Failed to update user metadata: ".concat(error_9.message));
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Store user metadata (company and project) in a SharePoint list
     */
    SharePointDataService.prototype.storeUserMetadata = function (libraryId, email, userId, company, project) {
        return __awaiter(this, void 0, void 0, function () {
            var metadata, storageKey;
            return __generator(this, function (_a) {
                try {
                    metadata = {
                        libraryId: libraryId,
                        email: email,
                        userId: userId,
                        company: company || '',
                        project: project || '',
                        timestamp: new Date().toISOString()
                    };
                    storageKey = "userMetadata_".concat(libraryId, "_").concat(userId);
                    localStorage.setItem(storageKey, JSON.stringify(metadata));
                    this.auditLogger.logInfo('storeUserMetadata', "Stored metadata for user ".concat(email), metadata);
                }
                catch (error) {
                    this.auditLogger.logError('storeUserMetadata', "Failed to store metadata for user ".concat(email), error);
                    // Don't throw error as this is supplementary functionality
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Retrieve user metadata (company and project)
     */
    SharePointDataService.prototype.getUserMetadata = function (libraryId, userId) {
        return __awaiter(this, void 0, void 0, function () {
            var storageKey, stored, metadata;
            return __generator(this, function (_a) {
                try {
                    storageKey = "userMetadata_".concat(libraryId, "_").concat(userId);
                    stored = localStorage.getItem(storageKey);
                    if (stored) {
                        metadata = JSON.parse(stored);
                        return [2 /*return*/, {
                                company: metadata.company || undefined,
                                project: metadata.project || undefined
                            }];
                    }
                    return [2 /*return*/, null];
                }
                catch (error) {
                    this.auditLogger.logError('getUserMetadata', "Failed to retrieve metadata for user ".concat(userId), error);
                    return [2 /*return*/, null];
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Bulk add external users to a library with specified permissions
     */
    SharePointDataService.prototype.bulkAddExternalUsersToLibrary = function (libraryId, request) {
        return __awaiter(this, void 0, void 0, function () {
            var results, sessionId, existingUsers, error_10, existingEmails, _i, _a, email, trimmedEmail, isExternal, error_11, successCount, alreadyMemberCount, failedCount;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        results = [];
                        sessionId = this.auditLogger.generateSessionId();
                        this.auditLogger.logInfo('bulkAddExternalUsersToLibrary', "Starting bulk addition of ".concat(request.emails.length, " users to library: ").concat(libraryId), {
                            libraryId: libraryId,
                            emailCount: request.emails.length,
                            permission: request.permission,
                            sessionId: sessionId
                        });
                        existingUsers = [];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.getExternalUsersForLibrary(libraryId)];
                    case 2:
                        existingUsers = _b.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        error_10 = _b.sent();
                        this.auditLogger.logWarning('bulkAddExternalUsersToLibrary', 'Failed to get existing users, proceeding without duplicate check', { error: error_10.message });
                        return [3 /*break*/, 4];
                    case 4:
                        existingEmails = new Set(existingUsers.map(function (user) { return user.email.toLowerCase(); }));
                        _i = 0, _a = request.emails;
                        _b.label = 5;
                    case 5:
                        if (!(_i < _a.length)) return [3 /*break*/, 11];
                        email = _a[_i];
                        trimmedEmail = email.trim().toLowerCase();
                        if (!trimmedEmail) {
                            results.push({
                                email: email,
                                status: 'failed',
                                message: 'Empty email address',
                                error: 'Email address is required'
                            });
                            return [3 /*break*/, 10];
                        }
                        // Validate email format
                        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmedEmail)) {
                            results.push({
                                email: email,
                                status: 'failed',
                                message: 'Invalid email format',
                                error: 'Please provide a valid email address'
                            });
                            return [3 /*break*/, 10];
                        }
                        // Check if user already has access
                        if (existingEmails.has(trimmedEmail)) {
                            results.push({
                                email: email,
                                status: 'already_member',
                                message: 'User already has access to this library'
                            });
                            this.auditLogger.logInfo('bulkAddExternalUsersToLibrary', "User ".concat(email, " already has access to library"), {
                                libraryId: libraryId,
                                email: email,
                                sessionId: sessionId
                            });
                            return [3 /*break*/, 10];
                        }
                        _b.label = 6;
                    case 6:
                        _b.trys.push([6, 9, , 10]);
                        return [4 /*yield*/, this.addExternalUserToLibrary(libraryId, trimmedEmail, request.permission, request.company, request.project)];
                    case 7:
                        _b.sent();
                        return [4 /*yield*/, this.isExternalUser(trimmedEmail)];
                    case 8:
                        isExternal = _b.sent();
                        results.push({
                            email: email,
                            status: isExternal ? 'invitation_sent' : 'success',
                            message: isExternal
                                ? 'Invitation sent to external user'
                                : 'User added successfully'
                        });
                        this.auditLogger.logInfo('bulkAddExternalUsersToLibrary', "Successfully added user ".concat(email, " to library"), {
                            libraryId: libraryId,
                            email: email,
                            permission: request.permission,
                            isExternal: isExternal,
                            company: request.company,
                            project: request.project,
                            sessionId: sessionId
                        });
                        return [3 /*break*/, 10];
                    case 9:
                        error_11 = _b.sent();
                        results.push({
                            email: email,
                            status: 'failed',
                            message: 'Failed to add user',
                            error: error_11.message
                        });
                        this.auditLogger.logError('bulkAddExternalUsersToLibrary', "Failed to add user ".concat(email, " to library"), {
                            libraryId: libraryId,
                            email: email,
                            error: error_11.message,
                            sessionId: sessionId
                        });
                        return [3 /*break*/, 10];
                    case 10:
                        _i++;
                        return [3 /*break*/, 5];
                    case 11:
                        successCount = results.filter(function (r) { return r.status === 'success' || r.status === 'invitation_sent'; }).length;
                        alreadyMemberCount = results.filter(function (r) { return r.status === 'already_member'; }).length;
                        failedCount = results.filter(function (r) { return r.status === 'failed'; }).length;
                        this.auditLogger.logInfo('bulkAddExternalUsersToLibrary', "Bulk addition completed. Success: ".concat(successCount, ", Already member: ").concat(alreadyMemberCount, ", Failed: ").concat(failedCount), {
                            libraryId: libraryId,
                            totalEmails: request.emails.length,
                            successCount: successCount,
                            alreadyMemberCount: alreadyMemberCount,
                            failedCount: failedCount,
                            sessionId: sessionId
                        });
                        return [2 /*return*/, results];
                }
            });
        });
    };
    /**
     * Search for users in the tenant (for adding external users)
     */
    SharePointDataService.prototype.searchUsers = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            var mockUser;
            return __generator(this, function (_a) {
                try {
                    this.auditLogger.logInfo('searchUsers', "Searching for users with query: ".concat(query));
                    // For external users, we'll use a simple approach of validating the email format
                    // In a real implementation, you might use Microsoft Graph API or SharePoint People Picker
                    if (!query || !query.includes('@')) {
                        return [2 /*return*/, []];
                    }
                    mockUser = {
                        id: 'search-result',
                        email: query,
                        displayName: query.split('@')[0],
                        invitedBy: this.context.pageContext.user.displayName,
                        invitedDate: new Date(),
                        lastAccess: new Date(),
                        permissions: 'Read'
                    };
                    this.auditLogger.logInfo('searchUsers', "Found 1 potential user for query: ".concat(query));
                    return [2 /*return*/, [mockUser]];
                }
                catch (error) {
                    this.auditLogger.logError('searchUsers', "Failed to search users with query: ".concat(query), error);
                    throw new Error("Failed to search users: ".concat(error.message));
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Get role definition ID for a permission level
     */
    SharePointDataService.prototype.getRoleDefinitionId = function (permission) {
        return __awaiter(this, void 0, void 0, function () {
            var roleName, roleDef, error_12;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        roleName = void 0;
                        switch (permission) {
                            case 'Read':
                                roleName = 'Read';
                                break;
                            case 'Contribute':
                                roleName = 'Contribute';
                                break;
                            case 'Full Control':
                                roleName = 'Full Control';
                                break;
                            default:
                                roleName = 'Read';
                        }
                        return [4 /*yield*/, this.sp.web.roleDefinitions.getByName(roleName).get()];
                    case 1:
                        roleDef = _a.sent();
                        if (!roleDef.Id) {
                            throw new Error("Role definition ID not found for permission: ".concat(permission));
                        }
                        return [2 /*return*/, roleDef.Id];
                    case 2:
                        error_12 = _a.sent();
                        this.auditLogger.logError('getRoleDefinitionId', "Failed to get role definition ID for permission: ".concat(permission), error_12);
                        throw new Error("Failed to get role definition: ".concat(error_12.message));
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Check if a user is external (not from the same tenant)
     */
    SharePointDataService.prototype.isExternalUser = function (email) {
        return __awaiter(this, void 0, void 0, function () {
            var currentDomain, emailDomain;
            return __generator(this, function (_a) {
                try {
                    currentDomain = this.context.pageContext.web.absoluteUrl.split('/')[2];
                    emailDomain = email.split('@')[1];
                    // If domains don't match, it's likely external
                    // This is a simplified check - in production you'd use Graph API for accurate determination
                    return [2 /*return*/, !currentDomain.includes(emailDomain) && !emailDomain.includes(currentDomain.split('.')[0])];
                }
                catch (_b) {
                    // If we can't determine, assume external for safety
                    return [2 /*return*/, true];
                }
                return [2 /*return*/];
            });
        });
    };
    // Private helper methods
    SharePointDataService.prototype.hasExternalUsers = function (libraryId) {
        return __awaiter(this, void 0, void 0, function () {
            var roleAssignments, externalCount, _i, roleAssignments_2, assignment, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId)
                                .roleAssignments
                                .expand("Member")
                                .get()];
                    case 1:
                        roleAssignments = _b.sent();
                        externalCount = 0;
                        for (_i = 0, roleAssignments_2 = roleAssignments; _i < roleAssignments_2.length; _i++) {
                            assignment = roleAssignments_2[_i];
                            if (assignment.Member.LoginName && assignment.Member.LoginName.includes('#ext#')) {
                                externalCount++;
                            }
                        }
                        return [2 /*return*/, {
                                hasExternal: externalCount > 0,
                                externalCount: externalCount
                            }];
                    case 2:
                        _a = _b.sent();
                        return [2 /*return*/, { hasExternal: false, externalCount: 0 }];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    SharePointDataService.prototype.getLibraryOwner = function (libraryId) {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var owner, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _c.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId)
                                .select("Author/Title")
                                .expand("Author")
                                .get()];
                    case 1:
                        owner = _c.sent();
                        return [2 /*return*/, ((_a = owner.Author) === null || _a === void 0 ? void 0 : _a.Title) || 'Unknown'];
                    case 2:
                        _b = _c.sent();
                        return [2 /*return*/, 'Unknown'];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    SharePointDataService.prototype.getLibraryPermissionLevel = function (libraryId) {
        return __awaiter(this, void 0, void 0, function () {
            var permissions, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.sp.web.lists.getById(libraryId)
                                .getCurrentUserEffectivePermissions()];
                    case 1:
                        permissions = _b.sent();
                        // This is a simplified permission check
                        // In reality, you'd need to parse the permission mask
                        return [2 /*return*/, 'Full Control']; // Default for now
                    case 2:
                        _a = _b.sent();
                        return [2 /*return*/, 'Read'];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    SharePointDataService.prototype.enableExternalSharing = function (libraryId) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                try {
                    // Note: External sharing configuration typically requires tenant admin permissions
                    // This is where Microsoft Graph API might be needed as a fallback
                    // For now, we'll log that this feature requires additional configuration
                    this.auditLogger.logInfo('enableExternalSharing', "External sharing enablement for library ".concat(libraryId, " requires tenant admin configuration"));
                }
                catch (error) {
                    this.auditLogger.logError('enableExternalSharing', "Failed to enable external sharing for library: ".concat(libraryId), error);
                }
                return [2 /*return*/];
            });
        });
    };
    SharePointDataService.prototype.mapPermissionLevel = function (permissions) {
        if (permissions.includes('Full Control'))
            return 'Full Control';
        if (permissions.includes('Contribute') || permissions.includes('Edit'))
            return 'Contribute';
        return 'Read';
    };
    return SharePointDataService;
}());
export { SharePointDataService };
//# sourceMappingURL=SharePointDataService.js.map