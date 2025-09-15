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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { Stack, Text, DetailsList, DetailsListLayoutMode, Selection, SelectionMode, CommandBar, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
import { MockDataService } from '../services/MockDataService';
import { SharePointDataService } from '../services/SharePointDataService';
import { CreateLibraryModal } from './CreateLibraryModal';
import { DeleteLibraryModal } from './DeleteLibraryModal';
import { ManageUsersModal } from './ManageUsersModal';
import styles from './ExternalUserManager.module.scss';
var ExternalUserManager = function (props) {
    var _a = useState([]), libraries = _a[0], setLibraries = _a[1];
    var _b = useState(true), loading = _b[0], setLoading = _b[1];
    var _c = useState([]), selectedLibraries = _c[0], setSelectedLibraries = _c[1];
    var _d = useState(false), showCreateModal = _d[0], setShowCreateModal = _d[1];
    var _e = useState(false), showDeleteModal = _e[0], setShowDeleteModal = _e[1];
    var _f = useState(false), showManageUsersModal = _f[0], setShowManageUsersModal = _f[1];
    var _g = useState(null), operationMessage = _g[0], setOperationMessage = _g[1];
    var dataService = useState(function () { return new SharePointDataService(props.context); })[0];
    var selection = useState(new Selection({
        onSelectionChanged: function () {
            setSelectedLibraries(selection.getSelection());
        }
    }))[0];
    var loadLibraries = useCallback(function (useMockData) {
        if (useMockData === void 0) { useMockData = false; }
        return __awaiter(void 0, void 0, void 0, function () {
            var librariesData, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        setLoading(true);
                        setOperationMessage(null);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 5, 6, 7]);
                        librariesData = void 0;
                        if (!useMockData) return [3 /*break*/, 2];
                        // For development/fallback, use mock data
                        librariesData = MockDataService.getExternalLibraries();
                        return [3 /*break*/, 4];
                    case 2: return [4 /*yield*/, dataService.getExternalLibraries()];
                    case 3:
                        // Use real SharePoint API
                        librariesData = _a.sent();
                        _a.label = 4;
                    case 4:
                        setLibraries(librariesData);
                        if (!useMockData && librariesData.length === 0) {
                            setOperationMessage({
                                message: 'No external libraries found. Libraries with external users will appear here.',
                                type: MessageBarType.info
                            });
                        }
                        return [3 /*break*/, 7];
                    case 5:
                        error_1 = _a.sent();
                        console.error('Error loading libraries:', error_1);
                        setOperationMessage({
                            message: "Failed to load libraries: ".concat(error_1.message, ". Falling back to demo data."),
                            type: MessageBarType.warning
                        });
                        // Fallback to mock data on error
                        setLibraries(MockDataService.getExternalLibraries());
                        return [3 /*break*/, 7];
                    case 6:
                        setLoading(false);
                        return [7 /*endfinally*/];
                    case 7: return [2 /*return*/];
                }
            });
        });
    }, [dataService]);
    useEffect(function () {
        // Try to load real data first, fallback to mock data if needed
        loadLibraries(false);
    }, [loadLibraries]);
    var handleCreateLibrary = function (config) { return __awaiter(void 0, void 0, void 0, function () {
        var newLibrary, error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.createLibrary(config)];
                case 1:
                    newLibrary = _a.sent();
                    return [2 /*return*/, newLibrary];
                case 2:
                    error_2 = _a.sent();
                    console.error('Error creating library:', error_2);
                    throw error_2;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleLibraryCreated = function (newLibrary) {
        setLibraries(function (prev) { return __spreadArray(__spreadArray([], prev, true), [newLibrary], false); });
        setOperationMessage({
            message: "Library \"".concat(newLibrary.name, "\" created successfully."),
            type: MessageBarType.success
        });
        // Clear message after 5 seconds
        setTimeout(function () { return setOperationMessage(null); }, 5000);
    };
    var handleDeleteLibrary = function (libraryId) { return __awaiter(void 0, void 0, void 0, function () {
        var error_3;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.deleteLibrary(libraryId)];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    error_3 = _a.sent();
                    console.error('Error deleting library:', error_3);
                    throw error_3;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleLibrariesDeleted = function (deletedIds) {
        setLibraries(function (prev) { return prev.filter(function (lib) { return deletedIds.indexOf(lib.id) === -1; }); });
        selection.setAllSelected(false);
        var count = deletedIds.length;
        setOperationMessage({
            message: "".concat(count, " librar").concat(count !== 1 ? 'ies' : 'y', " deleted successfully."),
            type: MessageBarType.success
        });
        // Clear message after 5 seconds
        setTimeout(function () { return setOperationMessage(null); }, 5000);
    };
    var handleAddUser = function (libraryId, email, permission, company, project) { return __awaiter(void 0, void 0, void 0, function () {
        var error_4;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.addExternalUserToLibrary(libraryId, email, permission, company, project)];
                case 1:
                    _a.sent();
                    // Update the external users count for the library
                    setLibraries(function (prev) { return prev.map(function (lib) {
                        return lib.id === libraryId
                            ? __assign(__assign({}, lib), { externalUsersCount: lib.externalUsersCount + 1 }) : lib;
                    }); });
                    return [3 /*break*/, 3];
                case 2:
                    error_4 = _a.sent();
                    console.error('Error adding user:', error_4);
                    throw error_4;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleBulkAddUsers = function (libraryId, emails, permission, company, project) { return __awaiter(void 0, void 0, void 0, function () {
        var results, successfulAdditions_1, error_5;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.bulkAddExternalUsersToLibrary(libraryId, {
                            emails: emails,
                            permission: permission,
                            company: company,
                            project: project
                        })];
                case 1:
                    results = _a.sent();
                    successfulAdditions_1 = results.filter(function (r) {
                        return r.status === 'success' || r.status === 'invitation_sent';
                    }).length;
                    // Update the external users count for the library
                    if (successfulAdditions_1 > 0) {
                        setLibraries(function (prev) { return prev.map(function (lib) {
                            return lib.id === libraryId
                                ? __assign(__assign({}, lib), { externalUsersCount: lib.externalUsersCount + successfulAdditions_1 }) : lib;
                        }); });
                    }
                    return [2 /*return*/, results];
                case 2:
                    error_5 = _a.sent();
                    console.error('Error bulk adding users:', error_5);
                    throw error_5;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleRemoveUser = function (libraryId, userId) { return __awaiter(void 0, void 0, void 0, function () {
        var error_6;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.removeExternalUserFromLibrary(libraryId, userId)];
                case 1:
                    _a.sent();
                    // Update the external users count for the library
                    setLibraries(function (prev) { return prev.map(function (lib) {
                        return lib.id === libraryId
                            ? __assign(__assign({}, lib), { externalUsersCount: Math.max(0, lib.externalUsersCount - 1) }) : lib;
                    }); });
                    return [3 /*break*/, 3];
                case 2:
                    error_6 = _a.sent();
                    console.error('Error removing user:', error_6);
                    throw error_6;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleUpdateUserMetadata = function (libraryId, userId, company, project) { return __awaiter(void 0, void 0, void 0, function () {
        var error_7;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.updateUserMetadata(libraryId, userId, company, project)];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    error_7 = _a.sent();
                    console.error('Error updating user metadata:', error_7);
                    throw error_7;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleGetUsers = function (libraryId) { return __awaiter(void 0, void 0, void 0, function () {
        var error_8;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.getExternalUsersForLibrary(libraryId)];
                case 1: return [2 /*return*/, _a.sent()];
                case 2:
                    error_8 = _a.sent();
                    console.error('Error getting users:', error_8);
                    throw error_8;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleSearchUsers = function (query) { return __awaiter(void 0, void 0, void 0, function () {
        var error_9;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.searchUsers(query)];
                case 1: return [2 /*return*/, _a.sent()];
                case 2:
                    error_9 = _a.sent();
                    console.error('Error searching users:', error_9);
                    throw error_9;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleUpdateUserMetadata = function (libraryId, userId, company, project) { return __awaiter(void 0, void 0, void 0, function () {
        var error_10;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, dataService.updateUserMetadata(libraryId, userId, company, project)];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    error_10 = _a.sent();
                    console.error('Error updating user metadata:', error_10);
                    throw error_10;
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var columns = [
        {
            key: 'name',
            name: 'Library Name',
            fieldName: 'name',
            minWidth: 200,
            maxWidth: 300,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Stack, null,
                React.createElement(Text, { variant: "medium", styles: { root: { fontWeight: 'semibold' } } }, item.name),
                React.createElement(Text, { variant: "small", styles: { root: { color: '#666' } } }, item.description))); }
        },
        {
            key: 'siteUrl',
            name: 'Site URL',
            fieldName: 'siteUrl',
            minWidth: 200,
            maxWidth: 250,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "small", styles: { root: { fontFamily: 'monospace' } } }, item.siteUrl)); }
        },
        {
            key: 'externalUsersCount',
            name: 'External Users',
            fieldName: 'externalUsersCount',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "medium", styles: { root: { textAlign: 'center' } } }, item.externalUsersCount)); }
        },
        {
            key: 'owner',
            name: 'Owner',
            fieldName: 'owner',
            minWidth: 150,
            maxWidth: 200,
            isResizable: true
        },
        {
            key: 'permissions',
            name: 'Permission Level',
            fieldName: 'permissions',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            onRender: function (item) {
                var permissionColor = item.permissions === 'Full Control' ? '#d73502' :
                    item.permissions === 'Contribute' ? '#f7630c' : '#0078d4';
                return (React.createElement(Text, { variant: "small", styles: {
                        root: {
                            color: permissionColor,
                            fontWeight: 'semibold',
                            padding: '4px 8px',
                            backgroundColor: "".concat(permissionColor, "15"),
                            borderRadius: '4px'
                        }
                    } }, item.permissions));
            }
        },
        {
            key: 'lastModified',
            name: 'Last Modified',
            fieldName: 'lastModified',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "small" }, item.lastModified.toLocaleDateString())); }
        }
    ];
    var commandBarItems = [
        {
            key: 'addLibrary',
            text: 'Add Library',
            iconProps: { iconName: 'Add' },
            onClick: function () { return setShowCreateModal(true); }
        },
        {
            key: 'removeLibrary',
            text: 'Remove',
            iconProps: { iconName: 'Delete' },
            disabled: selectedLibraries.length === 0,
            onClick: function () { return setShowDeleteModal(true); }
        },
        {
            key: 'manageUsers',
            text: 'Manage Users',
            iconProps: { iconName: 'People' },
            disabled: selectedLibraries.length !== 1,
            onClick: function () {
                // Open manage users modal for the selected library
                if (selectedLibraries.length === 1) {
                    setShowManageUsersModal(true);
                }
            }
        }
    ];
    var commandBarFarItems = [
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: function () {
                loadLibraries(false);
            }
        }
    ];
    return (React.createElement("div", { className: styles.externalUserManager },
        React.createElement(Stack, { tokens: { childrenGap: 20 } },
            React.createElement(Stack.Item, null,
                React.createElement(Text, { variant: "xxLarge", styles: { root: { fontWeight: 'semibold' } } }, "External User Manager"),
                React.createElement(Text, { variant: "medium", styles: { root: { color: '#666', marginTop: '8px' } } }, "Manage external users and shared libraries across SharePoint sites")),
            React.createElement(Stack.Item, null,
                React.createElement(MessageBar, { messageBarType: MessageBarType.info }, "Create and manage document libraries with external sharing capabilities. Use PnPjs for SharePoint integration with Microsoft Graph API fallback.")),
            operationMessage && (React.createElement(Stack.Item, null,
                React.createElement(MessageBar, { messageBarType: operationMessage.type }, operationMessage.message))),
            React.createElement(Stack.Item, null,
                React.createElement(CommandBar, { items: commandBarItems, farItems: commandBarFarItems, ariaLabel: "External Library Actions" })),
            React.createElement(Stack.Item, null, loading ? (React.createElement(Stack, { horizontalAlign: "center", tokens: { childrenGap: 10 } },
                React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading external libraries..." }))) : (React.createElement(DetailsList, { items: libraries, columns: columns, setKey: "set", layoutMode: DetailsListLayoutMode.justified, selection: selection, selectionPreservedOnEmptyClick: true, selectionMode: SelectionMode.multiple, ariaLabelForSelectionColumn: "Toggle selection", ariaLabelForSelectAllCheckbox: "Toggle selection for all items", checkButtonAriaLabel: "select row" }))),
            React.createElement(Stack.Item, null,
                React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 15 } },
                    React.createElement(Text, { variant: "small", styles: { root: { color: '#666' } } },
                        "Total Libraries: ",
                        libraries.length),
                    React.createElement(Text, { variant: "small", styles: { root: { color: '#666' } } },
                        "Selected: ",
                        selectedLibraries.length),
                    React.createElement(Text, { variant: "small", styles: { root: { color: '#666' } } },
                        "Total External Users: ",
                        libraries.reduce(function (sum, lib) { return sum + lib.externalUsersCount; }, 0))))),
        React.createElement(CreateLibraryModal, { isOpen: showCreateModal, onClose: function () { return setShowCreateModal(false); }, onLibraryCreated: handleLibraryCreated, onCreateLibrary: handleCreateLibrary }),
        React.createElement(DeleteLibraryModal, { isOpen: showDeleteModal, libraries: selectedLibraries, onClose: function () { return setShowDeleteModal(false); }, onLibrariesDeleted: handleLibrariesDeleted, onDeleteLibrary: handleDeleteLibrary }),
        React.createElement(ManageUsersModal, { isOpen: showManageUsersModal, library: selectedLibraries.length === 1 ? selectedLibraries[0] : null, onClose: function () { return setShowManageUsersModal(false); }, onAddUser: handleAddUser, onBulkAddUsers: handleBulkAddUsers, onRemoveUser: handleRemoveUser, onGetUsers: handleGetUsers, onSearchUsers: handleSearchUsers, onUpdateUserMetadata: handleUpdateUserMetadata })));
};
export default ExternalUserManager;
//# sourceMappingURL=ExternalUserManager.js.map