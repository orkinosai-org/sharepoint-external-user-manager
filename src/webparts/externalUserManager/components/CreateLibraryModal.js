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
import { Modal, Stack, Text, TextField, PrimaryButton, DefaultButton, Checkbox, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
export var CreateLibraryModal = function (_a) {
    var isOpen = _a.isOpen, onClose = _a.onClose, onLibraryCreated = _a.onLibraryCreated, onCreateLibrary = _a.onCreateLibrary;
    var _b = useState({
        title: '',
        description: '',
        enableExternalSharing: false
    }), formData = _b[0], setFormData = _b[1];
    var _c = useState(false), isCreating = _c[0], setIsCreating = _c[1];
    var _d = useState(''), error = _d[0], setError = _d[1];
    var _e = useState({}), validationErrors = _e[0], setValidationErrors = _e[1];
    var validateForm = function () {
        var errors = {};
        // Title validation
        if (!formData.title.trim()) {
            errors.title = 'Library name is required';
        }
        else if (formData.title.length > 100) {
            errors.title = 'Library name must be less than 100 characters';
        }
        else if (!/^[a-zA-Z0-9\s\-_]+$/.test(formData.title)) {
            errors.title = 'Library name can only contain letters, numbers, spaces, hyphens, and underscores';
        }
        // Description validation
        if (formData.description.length > 500) {
            errors.description = 'Description must be less than 500 characters';
        }
        setValidationErrors(errors);
        return Object.keys(errors).length === 0;
    };
    var handleInputChange = function (field) {
        return function (event, newValue) {
            setFormData(function (prev) {
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
    var handleCheckboxChange = function (field) {
        return function (event, checked) {
            setFormData(function (prev) {
                var _a;
                return (__assign(__assign({}, prev), (_a = {}, _a[field] = checked || false, _a)));
            });
        };
    };
    var handleCreate = function () { return __awaiter(void 0, void 0, void 0, function () {
        var newLibrary, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setError('');
                    if (!validateForm()) {
                        return [2 /*return*/];
                    }
                    setIsCreating(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, onCreateLibrary({
                            title: formData.title.trim(),
                            description: formData.description.trim() || undefined,
                            enableExternalSharing: formData.enableExternalSharing
                        })];
                case 2:
                    newLibrary = _a.sent();
                    onLibraryCreated(newLibrary);
                    handleClose();
                    return [3 /*break*/, 5];
                case 3:
                    err_1 = _a.sent();
                    setError(err_1.message || 'Failed to create library');
                    return [3 /*break*/, 5];
                case 4:
                    setIsCreating(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var handleClose = function () {
        if (!isCreating) {
            setFormData({
                title: '',
                description: '',
                enableExternalSharing: false
            });
            setError('');
            setValidationErrors({});
            onClose();
        }
    };
    var modalProps = {
        isOpen: isOpen,
        onDismiss: handleClose,
        isBlocking: isCreating,
        containerClassName: 'create-library-modal'
    };
    return (React.createElement(Modal, __assign({}, modalProps),
        React.createElement("div", { style: { padding: '20px', minWidth: '400px', maxWidth: '500px' } },
            React.createElement(Stack, { tokens: { childrenGap: 20 } },
                React.createElement(Stack.Item, null,
                    React.createElement(Text, { variant: "xLarge", styles: { root: { fontWeight: 'semibold' } } }, "Create New Library"),
                    React.createElement(Text, { variant: "medium", styles: { root: { color: '#666', marginTop: '8px' } } }, "Create a new document library for sharing with external users")),
                error && (React.createElement(Stack.Item, null,
                    React.createElement(MessageBar, { messageBarType: MessageBarType.error }, error))),
                React.createElement(Stack.Item, null,
                    React.createElement(TextField, { label: "Library Name *", value: formData.title, onChange: handleInputChange('title'), disabled: isCreating, errorMessage: validationErrors.title, placeholder: "Enter a name for the library", maxLength: 100, description: "The name will be used in URLs and must be unique within this site" })),
                React.createElement(Stack.Item, null,
                    React.createElement(TextField, { label: "Description", value: formData.description, onChange: handleInputChange('description'), disabled: isCreating, errorMessage: validationErrors.description, placeholder: "Enter a description (optional)", multiline: true, rows: 3, maxLength: 500, description: "Provide a brief description of the library's purpose" })),
                React.createElement(Stack.Item, null,
                    React.createElement(Checkbox, { label: "Enable external sharing", checked: formData.enableExternalSharing, onChange: handleCheckboxChange('enableExternalSharing'), disabled: isCreating }),
                    React.createElement(Text, { variant: "small", styles: { root: { color: '#666', marginTop: '4px' } } }, "Allow this library to be shared with users outside your organization")),
                formData.enableExternalSharing && (React.createElement(Stack.Item, null,
                    React.createElement(MessageBar, { messageBarType: MessageBarType.warning }, "External sharing must be enabled at the tenant and site level. Contact your administrator if external sharing is not working."))),
                React.createElement(Stack.Item, null,
                    React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 } },
                        React.createElement(PrimaryButton, { text: isCreating ? 'Creating...' : 'Create Library', onClick: handleCreate, disabled: isCreating || !formData.title.trim(), iconProps: isCreating ? undefined : { iconName: 'Add' } }),
                        isCreating && React.createElement(Spinner, { size: SpinnerSize.small }),
                        React.createElement(DefaultButton, { text: "Cancel", onClick: handleClose, disabled: isCreating }))),
                React.createElement(Stack.Item, null,
                    React.createElement(MessageBar, { messageBarType: MessageBarType.info },
                        React.createElement(Text, { variant: "small" },
                            React.createElement("strong", null, "Note:"),
                            " The new library will be created with default permissions. You can manage permissions and external sharing after creation.")))))));
};
//# sourceMappingURL=CreateLibraryModal.js.map