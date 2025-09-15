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
import { Modal, Stack, Text, TextField, PrimaryButton, DefaultButton, MessageBar, MessageBarType, Spinner, SpinnerSize, DetailsList, DetailsListLayoutMode, Selection, SelectionMode, CommandBar, Dropdown, IconButton, Dialog, DialogType, DialogFooter } from '@fluentui/react';
export var ManageUsersModal = function (_a) {
    var isOpen = _a.isOpen, library = _a.library, onClose = _a.onClose, onAddUser = _a.onAddUser, onBulkAddUsers = _a.onBulkAddUsers, onRemoveUser = _a.onRemoveUser, onGetUsers = _a.onGetUsers, onSearchUsers = _a.onSearchUsers, onUpdateUserMetadata = _a.onUpdateUserMetadata;
    var _b = useState([]), users = _b[0], setUsers = _b[1];
    var _c = useState([]), selectedUsers = _c[0], setSelectedUsers = _c[1];
    var _d = useState(false), loading = _d[0], setLoading = _d[1];
    var _e = useState(''), error = _e[0], setError = _e[1];
    var _f = useState(null), operationMessage = _f[0], setOperationMessage = _f[1];
    // Add User Form State
    var _g = useState(false), showAddUserForm = _g[0], setShowAddUserForm = _g[1];
    var _h = useState({
        email: '',
        emails: '',
        permission: 'Read',
        isBulkMode: false,
        company: '',
        project: ''
    }), addUserForm = _h[0], setAddUserForm = _h[1];
    var _j = useState(false), addingUser = _j[0], setAddingUser = _j[1];
    var _k = useState({}), validationErrors = _k[0], setValidationErrors = _k[1];
    var _l = useState(null), bulkResults = _l[0], setBulkResults = _l[1];
    // Remove User Confirmation
    var _m = useState(false), showRemoveConfirmation = _m[0], setShowRemoveConfirmation = _m[1];
    var _o = useState(false), removingUser = _o[0], setRemovingUser = _o[1];
    // Edit User Metadata
    var _p = useState(false), showEditModal = _p[0], setShowEditModal = _p[1];
    var _q = useState(null), editingUser = _q[0], setEditingUser = _q[1];
    var _r = useState({ company: '', project: '' }), editForm = _r[0], setEditForm = _r[1];
    var _s = useState(false), updatingUser = _s[0], setUpdatingUser = _s[1];
    var selection = useState(new Selection({
        onSelectionChanged: function () {
            setSelectedUsers(selection.getSelection());
        }
    }))[0];
    // Load users when modal opens
    useEffect(function () {
        if (isOpen && library) {
            loadUsers();
        }
        else if (!isOpen) {
            // Reset state when modal closes
            setUsers([]);
            setSelectedUsers([]);
            setError('');
            setOperationMessage(null);
            setShowAddUserForm(false);
            setShowRemoveConfirmation(false);
            setShowEditModal(false);
            setEditingUser(null);
            setBulkResults(null);
            setAddUserForm({ email: '', emails: '', permission: 'Read', isBulkMode: false, company: '', project: '' });
            selection.setAllSelected(false);
        }
    }, [isOpen, library]);
    var loadUsers = function () { return __awaiter(void 0, void 0, void 0, function () {
        var loadedUsers, err_1, errorMessage;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!library)
                        return [2 /*return*/];
                    setLoading(true);
                    setError('');
                    setOperationMessage(null);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, onGetUsers(library.id)];
                case 2:
                    loadedUsers = _a.sent();
                    setUsers(loadedUsers);
                    if (loadedUsers.length === 0) {
                        setOperationMessage({
                            message: 'No external users found for this library.',
                            type: MessageBarType.info
                        });
                    }
                    return [3 /*break*/, 5];
                case 3:
                    err_1 = _a.sent();
                    errorMessage = err_1.message || 'Failed to load users';
                    setError(errorMessage);
                    return [3 /*break*/, 5];
                case 4:
                    setLoading(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var handleAddUser = function () { return __awaiter(void 0, void 0, void 0, function () {
        var emails, results, successCount, alreadyMemberCount, failedCount, message, messageParts, err_2;
        var _a, _b, _c, _d;
        return __generator(this, function (_e) {
            switch (_e.label) {
                case 0:
                    if (!library)
                        return [2 /*return*/];
                    if (!validateAddUserForm()) {
                        return [2 /*return*/];
                    }
                    setAddingUser(true);
                    setError('');
                    setBulkResults(null);
                    _e.label = 1;
                case 1:
                    _e.trys.push([1, 7, 8, 9]);
                    if (!addUserForm.isBulkMode) return [3 /*break*/, 3];
                    emails = parseEmailsFromText(addUserForm.emails);
                    if (emails.length === 0) {
                        setError('Please enter at least one valid email address');
                        return [2 /*return*/];
                    }
                    return [4 /*yield*/, onBulkAddUsers(library.id, emails, addUserForm.permission, ((_a = addUserForm.company) === null || _a === void 0 ? void 0 : _a.trim()) || undefined, ((_b = addUserForm.project) === null || _b === void 0 ? void 0 : _b.trim()) || undefined)];
                case 2:
                    results = _e.sent();
                    setBulkResults(results);
                    successCount = results.filter(function (r) { return r.status === 'success' || r.status === 'invitation_sent'; }).length;
                    alreadyMemberCount = results.filter(function (r) { return r.status === 'already_member'; }).length;
                    failedCount = results.filter(function (r) { return r.status === 'failed'; }).length;
                    message = "Bulk operation completed: ";
                    messageParts = [];
                    if (successCount > 0)
                        messageParts.push("".concat(successCount, " added"));
                    if (alreadyMemberCount > 0)
                        messageParts.push("".concat(alreadyMemberCount, " already members"));
                    if (failedCount > 0)
                        messageParts.push("".concat(failedCount, " failed"));
                    message += messageParts.join(', ');
                    setOperationMessage({
                        message: message,
                        type: failedCount === 0 ? MessageBarType.success : MessageBarType.warning
                    });
                    return [3 /*break*/, 5];
                case 3: 
                // Single user addition
                return [4 /*yield*/, onAddUser(library.id, addUserForm.email.trim(), addUserForm.permission, ((_c = addUserForm.company) === null || _c === void 0 ? void 0 : _c.trim()) || undefined, ((_d = addUserForm.project) === null || _d === void 0 ? void 0 : _d.trim()) || undefined)];
                case 4:
                    // Single user addition
                    _e.sent();
                    setOperationMessage({
                        message: "Successfully added ".concat(addUserForm.email, " to ").concat(library.name),
                        type: MessageBarType.success
                    });
                    _e.label = 5;
                case 5:
                    // Reset form and reload users only on full success for single mode
                    // For bulk mode, keep the form open to show results
                    if (!addUserForm.isBulkMode) {
                        setAddUserForm({ email: '', emails: '', permission: 'Read', isBulkMode: false, company: '', project: '' });
                        setShowAddUserForm(false);
                    }
                    return [4 /*yield*/, loadUsers()];
                case 6:
                    _e.sent();
                    return [3 /*break*/, 9];
                case 7:
                    err_2 = _e.sent();
                    setError(err_2.message || 'Failed to add user(s)');
                    return [3 /*break*/, 9];
                case 8:
                    setAddingUser(false);
                    return [7 /*endfinally*/];
                case 9: return [2 /*return*/];
            }
        });
    }); };
    var parseEmailsFromText = function (text) {
        if (!text.trim())
            return [];
        // Split by comma, semicolon, or newline, then filter and trim
        return text
            .split(/[,;\n]/)
            .map(function (email) { return email.trim(); })
            .filter(function (email) { return email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email); });
    };
    var handleRemoveUsers = function () { return __awaiter(void 0, void 0, void 0, function () {
        var _i, selectedUsers_1, user, err_3;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!library || selectedUsers.length === 0)
                        return [2 /*return*/];
                    setRemovingUser(true);
                    setError('');
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 7, 8, 9]);
                    _i = 0, selectedUsers_1 = selectedUsers;
                    _a.label = 2;
                case 2:
                    if (!(_i < selectedUsers_1.length)) return [3 /*break*/, 5];
                    user = selectedUsers_1[_i];
                    return [4 /*yield*/, onRemoveUser(library.id, user.id)];
                case 3:
                    _a.sent();
                    _a.label = 4;
                case 4:
                    _i++;
                    return [3 /*break*/, 2];
                case 5:
                    setOperationMessage({
                        message: "Successfully removed ".concat(selectedUsers.length, " user(s) from ").concat(library.name),
                        type: MessageBarType.success
                    });
                    setShowRemoveConfirmation(false);
                    setSelectedUsers([]);
                    selection.setAllSelected(false);
                    return [4 /*yield*/, loadUsers()];
                case 6:
                    _a.sent();
                    return [3 /*break*/, 9];
                case 7:
                    err_3 = _a.sent();
                    setError(err_3.message || 'Failed to remove users');
                    return [3 /*break*/, 9];
                case 8:
                    setRemovingUser(false);
                    return [7 /*endfinally*/];
                case 9: return [2 /*return*/];
            }
        });
    }); };
    var validateAddUserForm = function () {
        var errors = {};
        if (addUserForm.isBulkMode) {
            // Bulk mode validation
            if (!addUserForm.emails.trim()) {
                errors.emails = 'Please enter at least one email address';
            }
            else {
                var emails = parseEmailsFromText(addUserForm.emails);
                if (emails.length === 0) {
                    errors.emails = 'Please enter valid email addresses (separated by commas, semicolons, or newlines)';
                }
            }
        }
        else {
            // Single mode validation
            if (!addUserForm.email.trim()) {
                errors.email = 'Email address is required';
            }
            else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(addUserForm.email.trim())) {
                errors.email = 'Please enter a valid email address';
            }
        }
        setValidationErrors(errors);
        return Object.keys(errors).length === 0;
    };
    var handleInputChange = function (field) {
        return function (event, newValue) {
            setAddUserForm(function (prev) {
                var _a;
                return (__assign(__assign({}, prev), (_a = {}, _a[field] = newValue || '', _a)));
            });
            // Clear validation error for this field
            if (validationErrors[field]) {
                setValidationErrors(function (prev) {
                    var newErrors = __assign({}, prev);
                    delete newErrors[field];
                    return newErrors;
                });
            }
        };
    };
    var handlePermissionChange = function (event, option) {
        if (option) {
            setAddUserForm(function (prev) { return (__assign(__assign({}, prev), { permission: option.key })); });
        }
    };
    var handleClose = function () {
        if (!addingUser && !removingUser && !updatingUser && !loading) {
            onClose();
        }
    };
    var handleEditUser = function (user) {
        setEditingUser(user);
        setEditForm({
            company: user.company || '',
            project: user.project || ''
        });
        setShowEditModal(true);
    };
    var handleUpdateUserMetadata = function () { return __awaiter(void 0, void 0, void 0, function () {
        var error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!editingUser || !library)
                        return [2 /*return*/];
                    setUpdatingUser(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    // Call the service method through the prop
                    return [4 /*yield*/, onUpdateUserMetadata(library.id, editingUser.id, editForm.company, editForm.project)];
                case 2:
                    // Call the service method through the prop
                    _a.sent();
                    // Update the local users list
                    setUsers(function (prev) { return prev.map(function (user) {
                        return user.id === editingUser.id
                            ? __assign(__assign({}, user), { company: editForm.company || undefined, project: editForm.project || undefined }) : user;
                    }); });
                    setOperationMessage({
                        message: "Successfully updated metadata for ".concat(editingUser.displayName),
                        type: MessageBarType.success
                    });
                    setShowEditModal(false);
                    setEditingUser(null);
                    return [3 /*break*/, 5];
                case 3:
                    error_1 = _a.sent();
                    setOperationMessage({
                        message: "Failed to update metadata: ".concat(error_1.message),
                        type: MessageBarType.error
                    });
                    return [3 /*break*/, 5];
                case 4:
                    setUpdatingUser(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    // Define columns for the users list
    var columns = [
        {
            key: 'displayName',
            name: 'Name',
            fieldName: 'displayName',
            minWidth: 150,
            maxWidth: 200,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "medium" }, item.displayName || 'Unknown')); }
        },
        {
            key: 'email',
            name: 'Email',
            fieldName: 'email',
            minWidth: 200,
            maxWidth: 250,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "small" }, item.email)); }
        },
        {
            key: 'permissions',
            name: 'Permissions',
            fieldName: 'permissions',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "small" }, item.permissions)); }
        },
        {
            key: 'company',
            name: 'Company',
            fieldName: 'company',
            minWidth: 120,
            maxWidth: 180,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "small" }, item.company || '-')); }
        },
        {
            key: 'project',
            name: 'Project',
            fieldName: 'project',
            minWidth: 120,
            maxWidth: 180,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "small" }, item.project || '-')); }
        },
        {
            key: 'invitedDate',
            name: 'Invited Date',
            fieldName: 'invitedDate',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Text, { variant: "small" }, item.invitedDate.toLocaleDateString())); }
        },
        {
            key: 'actions',
            name: 'Actions',
            fieldName: 'actions',
            minWidth: 100,
            maxWidth: 120,
            isResizable: false,
            onRender: function (item) { return (React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 5 } },
                React.createElement(IconButton, { iconProps: { iconName: 'Edit' }, title: "Edit Company/Project", ariaLabel: "Edit user metadata", onClick: function () { return handleEditUser(item); }, styles: {
                        root: { minWidth: 24, width: 24, height: 24 },
                        icon: { fontSize: 12 }
                    } }))); }
        }
    ];
    var commandBarItems = [
        {
            key: 'addUser',
            text: 'Add User',
            iconProps: { iconName: 'AddFriend' },
            onClick: function () {
                setAddUserForm(function (prev) { return (__assign(__assign({}, prev), { isBulkMode: false })); });
                setShowAddUserForm(true);
            }
        },
        {
            key: 'bulkAddUsers',
            text: 'Bulk Add Users',
            iconProps: { iconName: 'AddGroup' },
            onClick: function () {
                setAddUserForm(function (prev) { return (__assign(__assign({}, prev), { isBulkMode: true })); });
                setBulkResults(null);
                setShowAddUserForm(true);
            }
        },
        {
            key: 'removeUser',
            text: 'Remove User',
            iconProps: { iconName: 'RemoveFriend' },
            disabled: selectedUsers.length === 0,
            onClick: function () { return setShowRemoveConfirmation(true); }
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: loadUsers
        }
    ];
    var permissionOptions = [
        { key: 'Read', text: 'Read' },
        { key: 'Contribute', text: 'Contribute' },
        { key: 'Full Control', text: 'Full Control' }
    ];
    var modalProps = {
        isOpen: isOpen,
        onDismiss: handleClose,
        isBlocking: addingUser || removingUser || loading,
        containerClassName: 'manage-users-modal'
    };
    return (React.createElement(React.Fragment, null,
        React.createElement(Modal, __assign({}, modalProps),
            React.createElement("div", { style: { padding: '20px', minWidth: '600px', maxWidth: '800px' } },
                React.createElement(Stack, { tokens: { childrenGap: 20 } },
                    React.createElement(Stack.Item, null,
                        React.createElement(Stack, { horizontal: true, horizontalAlign: "space-between", verticalAlign: "center" },
                            React.createElement(Stack, null,
                                React.createElement(Text, { variant: "xLarge", styles: { root: { fontWeight: 'semibold' } } }, "Manage External Users"),
                                React.createElement(Text, { variant: "medium", styles: { root: { color: '#666', marginTop: '4px' } } }, library ? "Library: ".concat(library.name) : 'No library selected')),
                            React.createElement(IconButton, { iconProps: { iconName: 'Cancel' }, ariaLabel: "Close", onClick: handleClose, disabled: addingUser || removingUser || loading }))),
                    error && (React.createElement(Stack.Item, null,
                        React.createElement(MessageBar, { messageBarType: MessageBarType.error }, error))),
                    operationMessage && (React.createElement(Stack.Item, null,
                        React.createElement(MessageBar, { messageBarType: operationMessage.type }, operationMessage.message))),
                    React.createElement(Stack.Item, null,
                        React.createElement(CommandBar, { items: commandBarItems, ariaLabel: "User Management Actions" })),
                    React.createElement(Stack.Item, null, loading ? (React.createElement(Stack, { horizontalAlign: "center", tokens: { childrenGap: 10 } },
                        React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading users..." }))) : (React.createElement(DetailsList, { items: users, columns: columns, setKey: "set", layoutMode: DetailsListLayoutMode.justified, selection: selection, selectionPreservedOnEmptyClick: true, selectionMode: SelectionMode.multiple, ariaLabelForSelectionColumn: "Toggle selection", ariaLabelForSelectAllCheckbox: "Toggle selection for all items", checkButtonAriaLabel: "select row" }))),
                    React.createElement(Stack.Item, null,
                        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 15 } },
                            React.createElement(Text, { variant: "small", styles: { root: { color: '#666' } } },
                                "Total External Users: ",
                                users.length),
                            React.createElement(Text, { variant: "small", styles: { root: { color: '#666' } } },
                                "Selected: ",
                                selectedUsers.length))),
                    showAddUserForm && (React.createElement(Stack.Item, null,
                        React.createElement("div", { style: {
                                border: '1px solid #edebe9',
                                borderRadius: '4px',
                                padding: '16px',
                                backgroundColor: '#faf9f8'
                            } },
                            React.createElement(Stack, { tokens: { childrenGap: 15 } },
                                React.createElement(Stack.Item, null,
                                    React.createElement(Text, { variant: "mediumPlus", styles: { root: { fontWeight: 'semibold' } } }, addUserForm.isBulkMode ? 'Bulk Add External Users' : 'Add External User')),
                                React.createElement(Stack.Item, null,
                                    React.createElement(Stack, { horizontal: !addUserForm.isBulkMode, tokens: { childrenGap: 10 } },
                                        React.createElement(Stack.Item, { grow: true }, addUserForm.isBulkMode ? (React.createElement(TextField, { label: "Email Addresses *", multiline: true, rows: 6, value: addUserForm.emails, onChange: handleInputChange('emails'), disabled: addingUser, errorMessage: validationErrors.emails, placeholder: "Enter multiple email addresses separated by commas, semicolons, or new lines\n\nExample:\nuser1@external.com\nuser2@partner.com, user3@vendor.com", description: "Enter multiple email addresses for bulk invitation" })) : (React.createElement(TextField, { label: "Email Address *", value: addUserForm.email, onChange: handleInputChange('email'), disabled: addingUser, errorMessage: validationErrors.email, placeholder: "Enter user's email address", description: "Enter the email address of the external user to invite" }))),
                                        !addUserForm.isBulkMode && (React.createElement(Stack.Item, null,
                                            React.createElement(Dropdown, { label: "Permission Level *", options: permissionOptions, selectedKey: addUserForm.permission, onChange: handlePermissionChange, disabled: addingUser, styles: { dropdown: { width: 150 } } }))))),
                                addUserForm.isBulkMode && (React.createElement(Stack.Item, null,
                                    React.createElement(Dropdown, { label: "Permission Level for all users *", options: permissionOptions, selectedKey: addUserForm.permission, onChange: handlePermissionChange, disabled: addingUser, styles: { dropdown: { width: 200 } } }))),
                                React.createElement(Stack.Item, null,
                                    React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 } },
                                        React.createElement(Stack.Item, { grow: true },
                                            React.createElement(TextField, { label: "Company", value: addUserForm.company, onChange: handleInputChange('company'), disabled: addingUser, placeholder: "Enter company name", description: "Company or organization the user belongs to" })),
                                        React.createElement(Stack.Item, { grow: true },
                                            React.createElement(TextField, { label: "Project", value: addUserForm.project, onChange: handleInputChange('project'), disabled: addingUser, placeholder: "Enter project name", description: "Project or initiative the user is associated with" })))),
                                bulkResults && (React.createElement(Stack.Item, null,
                                    React.createElement("div", { style: {
                                            border: '1px solid #edebe9',
                                            borderRadius: '4px',
                                            padding: '12px',
                                            backgroundColor: '#fff'
                                        } },
                                        React.createElement(Text, { variant: "medium", styles: { root: { fontWeight: 'semibold', marginBottom: '8px' } } }, "Bulk Operation Results:"),
                                        React.createElement("div", { style: { maxHeight: '200px', overflowY: 'auto' } }, bulkResults.map(function (result, index) { return (React.createElement("div", { key: index, style: {
                                                display: 'flex',
                                                justifyContent: 'space-between',
                                                padding: '4px 0',
                                                borderBottom: index < bulkResults.length - 1 ? '1px solid #f3f2f1' : 'none'
                                            } },
                                            React.createElement(Text, { variant: "small" }, result.email),
                                            React.createElement(Text, { variant: "small", styles: {
                                                    root: {
                                                        color: result.status === 'success' || result.status === 'invitation_sent'
                                                            ? '#107c10'
                                                            : result.status === 'already_member'
                                                                ? '#797775'
                                                                : '#d13438',
                                                        fontWeight: 'semibold'
                                                    }
                                                } }, result.status === 'success' ? '✓ Added' :
                                                result.status === 'invitation_sent' ? '✓ Invited' :
                                                    result.status === 'already_member' ? '- Already member' :
                                                        '✗ Failed'))); }))))),
                                React.createElement(Stack.Item, null,
                                    React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 } },
                                        React.createElement(PrimaryButton, { text: addingUser
                                                ? (addUserForm.isBulkMode ? 'Adding Users...' : 'Adding...')
                                                : (addUserForm.isBulkMode ? 'Add Users' : 'Add User'), onClick: handleAddUser, disabled: addingUser || (addUserForm.isBulkMode
                                                ? !addUserForm.emails.trim()
                                                : !addUserForm.email.trim()), iconProps: addingUser ? undefined : { iconName: addUserForm.isBulkMode ? 'AddGroup' : 'AddFriend' } }),
                                        React.createElement(DefaultButton, { text: "Cancel", onClick: function () {
                                                setShowAddUserForm(false);
                                                setAddUserForm({ email: '', emails: '', permission: 'Read', isBulkMode: false, company: '', project: '' });
                                                setValidationErrors({});
                                                setBulkResults(null);
                                            }, disabled: addingUser }),
                                        bulkResults && (React.createElement(DefaultButton, { text: "Close Results", onClick: function () {
                                                setShowAddUserForm(false);
                                                setAddUserForm({ email: '', emails: '', permission: 'Read', isBulkMode: false, company: '', project: '' });
                                                setValidationErrors({});
                                                setBulkResults(null);
                                            } })))))))),
                    React.createElement(Stack.Item, null,
                        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 } },
                            React.createElement(DefaultButton, { text: "Close", onClick: handleClose, disabled: addingUser || removingUser || loading })))))),
        React.createElement(Dialog, { hidden: !showRemoveConfirmation, onDismiss: function () { return setShowRemoveConfirmation(false); }, dialogContentProps: {
                type: DialogType.normal,
                title: 'Remove External Users',
                subText: "Are you sure you want to remove ".concat(selectedUsers.length, " user(s) from \"").concat(library === null || library === void 0 ? void 0 : library.name, "\"? This action cannot be undone.")
            }, modalProps: {
                isBlocking: removingUser
            } },
            React.createElement(DialogFooter, null,
                React.createElement(PrimaryButton, { onClick: handleRemoveUsers, text: removingUser ? 'Removing...' : 'Remove', disabled: removingUser }),
                React.createElement(DefaultButton, { onClick: function () { return setShowRemoveConfirmation(false); }, text: "Cancel", disabled: removingUser }))),
        React.createElement(Dialog, { hidden: !showEditModal, onDismiss: function () { return setShowEditModal(false); }, dialogContentProps: {
                type: DialogType.normal,
                title: 'Edit User Metadata',
                subText: "Update company and project information for ".concat((editingUser === null || editingUser === void 0 ? void 0 : editingUser.displayName) || (editingUser === null || editingUser === void 0 ? void 0 : editingUser.email))
            }, modalProps: {
                isBlocking: updatingUser
            }, minWidth: 400 },
            React.createElement(Stack, { tokens: { childrenGap: 15 } },
                React.createElement(TextField, { label: "Company", value: editForm.company, onChange: function (event, newValue) { return setEditForm(function (prev) { return (__assign(__assign({}, prev), { company: newValue || '' })); }); }, disabled: updatingUser, placeholder: "Enter company name" }),
                React.createElement(TextField, { label: "Project", value: editForm.project, onChange: function (event, newValue) { return setEditForm(function (prev) { return (__assign(__assign({}, prev), { project: newValue || '' })); }); }, disabled: updatingUser, placeholder: "Enter project name" })),
            React.createElement(DialogFooter, null,
                React.createElement(PrimaryButton, { onClick: handleUpdateUserMetadata, text: updatingUser ? 'Updating...' : 'Update', disabled: updatingUser }),
                React.createElement(DefaultButton, { onClick: function () { return setShowEditModal(false); }, text: "Cancel", disabled: updatingUser })))));
};
//# sourceMappingURL=ManageUsersModal.js.map