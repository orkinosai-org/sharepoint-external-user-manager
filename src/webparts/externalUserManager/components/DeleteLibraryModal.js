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
import { useState } from 'react';
import { Modal, Stack, Text, PrimaryButton, DefaultButton, MessageBar, MessageBarType, Spinner, SpinnerSize, Checkbox, TextField } from '@fluentui/react';
export var DeleteLibraryModal = function (_a) {
    var isOpen = _a.isOpen, libraries = _a.libraries, onClose = _a.onClose, onLibrariesDeleted = _a.onLibrariesDeleted, onDeleteLibrary = _a.onDeleteLibrary;
    var _b = useState(false), isDeleting = _b[0], setIsDeleting = _b[1];
    var _c = useState(''), confirmationText = _c[0], setConfirmationText = _c[1];
    var _d = useState(false), acknowledgeDataLoss = _d[0], setAcknowledgeDataLoss = _d[1];
    var _e = useState(''), error = _e[0], setError = _e[1];
    var _f = useState({}), deletionProgress = _f[0], setDeletionProgress = _f[1];
    var isMultipleLibraries = libraries.length > 1;
    var confirmationPhrase = 'DELETE';
    var isConfirmationValid = confirmationText === confirmationPhrase;
    var canDelete = isConfirmationValid && acknowledgeDataLoss && !isDeleting;
    var totalExternalUsers = libraries.reduce(function (sum, lib) { return sum + lib.externalUsersCount; }, 0);
    var hasExternalUsers = totalExternalUsers > 0;
    var resetForm = function () {
        setConfirmationText('');
        setAcknowledgeDataLoss(false);
        setError('');
        setDeletionProgress({});
    };
    var handleClose = function () {
        if (!isDeleting) {
            resetForm();
            onClose();
        }
    };
    var handleDelete = function () { return __awaiter(void 0, void 0, void 0, function () {
        var initialProgress, deletedIds, hasErrors, _loop_1, _i, libraries_1, library;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!canDelete)
                        return [2 /*return*/];
                    setError('');
                    setIsDeleting(true);
                    initialProgress = {};
                    libraries.forEach(function (lib) {
                        initialProgress[lib.id] = 'pending';
                    });
                    setDeletionProgress(initialProgress);
                    deletedIds = [];
                    hasErrors = false;
                    _loop_1 = function (library) {
                        var err_1;
                        return __generator(this, function (_b) {
                            switch (_b.label) {
                                case 0:
                                    _b.trys.push([0, 2, , 3]);
                                    setDeletionProgress(function (prev) {
                                        var _a;
                                        return (__assign(__assign({}, prev), (_a = {}, _a[library.id] = 'deleting', _a)));
                                    });
                                    return [4 /*yield*/, onDeleteLibrary(library.id)];
                                case 1:
                                    _b.sent();
                                    setDeletionProgress(function (prev) {
                                        var _a;
                                        return (__assign(__assign({}, prev), (_a = {}, _a[library.id] = 'completed', _a)));
                                    });
                                    deletedIds.push(library.id);
                                    return [3 /*break*/, 3];
                                case 2:
                                    err_1 = _b.sent();
                                    hasErrors = true;
                                    setDeletionProgress(function (prev) {
                                        var _a;
                                        return (__assign(__assign({}, prev), (_a = {}, _a[library.id] = 'failed', _a)));
                                    });
                                    setError(function (prev) {
                                        var newError = "Failed to delete \"".concat(library.name, "\": ").concat(err_1.message);
                                        return prev ? "".concat(prev, "\n").concat(newError) : newError;
                                    });
                                    return [3 /*break*/, 3];
                                case 3: return [2 /*return*/];
                            }
                        });
                    };
                    _i = 0, libraries_1 = libraries;
                    _a.label = 1;
                case 1:
                    if (!(_i < libraries_1.length)) return [3 /*break*/, 4];
                    library = libraries_1[_i];
                    return [5 /*yield**/, _loop_1(library)];
                case 2:
                    _a.sent();
                    _a.label = 3;
                case 3:
                    _i++;
                    return [3 /*break*/, 1];
                case 4:
                    setIsDeleting(false);
                    // If some libraries were deleted successfully, notify parent
                    if (deletedIds.length > 0) {
                        onLibrariesDeleted(deletedIds);
                    }
                    // If all libraries were deleted successfully, close the modal
                    if (!hasErrors) {
                        setTimeout(function () {
                            handleClose();
                        }, 1000); // Give user time to see success message
                    }
                    return [2 /*return*/];
            }
        });
    }); };
    var getProgressIcon = function (status) {
        switch (status) {
            case 'pending': return 'Clock';
            case 'deleting': return 'Sync';
            case 'completed': return 'CheckMark';
            case 'failed': return 'ErrorBadge';
            default: return 'Clock';
        }
    };
    var getProgressColor = function (status) {
        switch (status) {
            case 'pending': return '#666';
            case 'deleting': return '#0078d4';
            case 'completed': return '#107c10';
            case 'failed': return '#d13438';
            default: return '#666';
        }
    };
    var modalProps = {
        isOpen: isOpen,
        onDismiss: handleClose,
        isBlocking: isDeleting,
        containerClassName: 'delete-library-modal'
    };
    return (React.createElement(Modal, __assign({}, modalProps),
        React.createElement("div", { style: { padding: '20px', minWidth: '450px', maxWidth: '600px' } },
            React.createElement(Stack, { tokens: { childrenGap: 20 } },
                React.createElement(Stack.Item, null,
                    React.createElement(Text, { variant: "xLarge", styles: { root: { fontWeight: 'semibold', color: '#d13438' } } },
                        "Delete ",
                        isMultipleLibraries ? 'Libraries' : 'Library'),
                    React.createElement(Text, { variant: "medium", styles: { root: { color: '#666', marginTop: '8px' } } }, "This action cannot be undone. All content and settings will be permanently deleted.")),
                React.createElement(Stack.Item, null,
                    React.createElement(Text, { variant: "mediumPlus", styles: { root: { fontWeight: 'semibold' } } }, isMultipleLibraries ? 'Libraries to be deleted:' : 'Library to be deleted:'),
                    libraries.map(function (library) { return (React.createElement(Stack, { key: library.id, horizontal: true, verticalAlign: "center", tokens: { childrenGap: 10 }, styles: { root: { marginTop: '8px', padding: '8px', backgroundColor: '#fef9f9', border: '1px solid #fde7e9' } } },
                        isDeleting && (React.createElement("div", { style: { color: getProgressColor(deletionProgress[library.id] || 'pending') } }, deletionProgress[library.id] === 'deleting' ? (React.createElement(Spinner, { size: SpinnerSize.small })) : (React.createElement(Text, null, getProgressIcon(deletionProgress[library.id] || 'pending'))))),
                        React.createElement(Stack, null,
                            React.createElement(Text, { variant: "medium", styles: { root: { fontWeight: 'semibold' } } }, library.name),
                            React.createElement(Text, { variant: "small", styles: { root: { color: '#666' } } },
                                library.externalUsersCount,
                                " external user",
                                library.externalUsersCount !== 1 ? 's' : '',
                                " \u2022 Owner: ",
                                library.owner)))); })),
                hasExternalUsers && (React.createElement(Stack.Item, null,
                    React.createElement(MessageBar, { messageBarType: MessageBarType.warning },
                        React.createElement("strong", null, "Warning:"),
                        " ",
                        totalExternalUsers,
                        " external user",
                        totalExternalUsers !== 1 ? 's' : '',
                        isMultipleLibraries ? ' have' : ' has',
                        " access to ",
                        isMultipleLibraries ? 'these libraries' : 'this library',
                        ". They will lose access when the ",
                        isMultipleLibraries ? 'libraries are' : 'library is',
                        " deleted."))),
                error && (React.createElement(Stack.Item, null,
                    React.createElement(MessageBar, { messageBarType: MessageBarType.error },
                        React.createElement("pre", { style: { whiteSpace: 'pre-wrap', margin: 0 } }, error)))),
                React.createElement(Stack.Item, null,
                    React.createElement(Checkbox, { label: "I understand that all content in ".concat(isMultipleLibraries ? 'these libraries' : 'this library', " will be permanently deleted"), checked: acknowledgeDataLoss, onChange: function (_, checked) { return setAcknowledgeDataLoss(checked || false); }, disabled: isDeleting })),
                React.createElement(Stack.Item, null,
                    React.createElement(TextField, { label: "Type \"".concat(confirmationPhrase, "\" to confirm deletion"), value: confirmationText, onChange: function (_, newValue) { return setConfirmationText(newValue || ''); }, disabled: isDeleting, placeholder: confirmationPhrase, styles: {
                            fieldGroup: {
                                borderColor: isConfirmationValid ? '#107c10' : undefined
                            }
                        } })),
                React.createElement(Stack.Item, null,
                    React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 } },
                        React.createElement(PrimaryButton, { text: isDeleting ? 'Deleting...' : "Delete ".concat(isMultipleLibraries ? 'Libraries' : 'Library'), onClick: handleDelete, disabled: !canDelete, iconProps: isDeleting ? undefined : { iconName: 'Delete' }, styles: {
                                root: {
                                    backgroundColor: '#d13438',
                                    borderColor: '#d13438'
                                },
                                rootHovered: {
                                    backgroundColor: '#b91b1b',
                                    borderColor: '#b91b1b'
                                }
                            } }),
                        isDeleting && React.createElement(Spinner, { size: SpinnerSize.small }),
                        React.createElement(DefaultButton, { text: "Cancel", onClick: handleClose, disabled: isDeleting }))),
                isDeleting && Object.keys(deletionProgress).length > 0 && (React.createElement(Stack.Item, null,
                    React.createElement(MessageBar, { messageBarType: MessageBarType.info },
                        "Deleting ",
                        libraries.length,
                        " librar",
                        libraries.length !== 1 ? 'ies' : 'y',
                        "...",
                        React.createElement("br", null),
                        "Completed: ",
                        Object.keys(deletionProgress).filter(function (id) { return deletionProgress[id] === 'completed'; }).length,
                        Object.keys(deletionProgress).some(function (id) { return deletionProgress[id] === 'failed'; }) &&
                            " \u2022 Failed: ".concat(Object.keys(deletionProgress).filter(function (id) { return deletionProgress[id] === 'failed'; }).length))))))));
};
//# sourceMappingURL=DeleteLibraryModal.js.map