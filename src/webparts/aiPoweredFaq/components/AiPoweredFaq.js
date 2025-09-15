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
import { useState, useEffect } from 'react';
import { Stack, SearchBox, CommandBar, DetailsList, SelectionMode, Spinner, SpinnerSize, MessageBar, MessageBarType, Panel, PanelType, Text, Rating, Separator, Pivot, PivotItem } from '@fluentui/react';
import { MockFaqDataService } from '../services/MockFaqDataService';
import { AzureAiService } from '../services/AzureAiService';
import styles from './AiPoweredFaq.module.scss';
var AiPoweredFaq = function (props) {
    var _a = useState([]), faqItems = _a[0], setFaqItems = _a[1];
    var _b = useState(''), searchQuery = _b[0], setSearchQuery = _b[1];
    var _c = useState(true), loading = _c[0], setLoading = _c[1];
    var _d = useState(null), selectedItem = _d[0], setSelectedItem = _d[1];
    var _e = useState(false), showPanel = _e[0], setShowPanel = _e[1];
    var _f = useState([]), chatMessages = _f[0], setChatMessages = _f[1];
    var _g = useState(false), aiProcessing = _g[0], setAiProcessing = _g[1];
    useEffect(function () {
        loadFaqData();
    }, []);
    var loadFaqData = function () { return __awaiter(void 0, void 0, void 0, function () {
        return __generator(this, function (_a) {
            setLoading(true);
            try {
                setTimeout(function () {
                    var mockFaqs = MockFaqDataService.getFaqItems();
                    setFaqItems(mockFaqs);
                    setLoading(false);
                }, 1000);
            }
            catch (error) {
                console.error('Error loading FAQ data:', error);
                setLoading(false);
            }
            return [2 /*return*/];
        });
    }); };
    var handleSearch = function (query) { return __awaiter(void 0, void 0, void 0, function () {
        var filteredItems, aiSuggestions, aiError_1, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setSearchQuery(query);
                    if (query.trim() === '') {
                        loadFaqData();
                        return [2 /*return*/];
                    }
                    setLoading(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 7, , 8]);
                    filteredItems = faqItems.filter(function (item) {
                        return item.question.toLowerCase().includes(query.toLowerCase()) ||
                            item.answer.toLowerCase().includes(query.toLowerCase()) ||
                            item.tags.some(function (tag) { return tag.toLowerCase().includes(query.toLowerCase()); });
                    });
                    setFaqItems(filteredItems);
                    if (!(props.enableAiSuggestions && props.azureOpenAiApiKey)) return [3 /*break*/, 6];
                    setAiProcessing(true);
                    _a.label = 2;
                case 2:
                    _a.trys.push([2, 4, , 5]);
                    return [4 /*yield*/, AzureAiService.getAnswerSuggestions(query, props.azureOpenAiEndpoint, props.azureOpenAiApiKey)];
                case 3:
                    aiSuggestions = _a.sent();
                    // Handle AI suggestions here
                    console.log('AI Suggestions:', aiSuggestions);
                    return [3 /*break*/, 5];
                case 4:
                    aiError_1 = _a.sent();
                    console.error('AI service error:', aiError_1);
                    return [3 /*break*/, 5];
                case 5:
                    setAiProcessing(false);
                    _a.label = 6;
                case 6:
                    setLoading(false);
                    return [3 /*break*/, 8];
                case 7:
                    error_1 = _a.sent();
                    console.error('Error searching FAQ:', error_1);
                    setLoading(false);
                    return [3 /*break*/, 8];
                case 8: return [2 /*return*/];
            }
        });
    }); };
    var handleAskAi = function (question) { return __awaiter(void 0, void 0, void 0, function () {
        var newMessage, aiResponse, aiMessage_1, mockResponse_1, error_2, errorMessage_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    newMessage = {
                        id: Date.now().toString(),
                        message: question,
                        isUser: true,
                        timestamp: new Date()
                    };
                    setChatMessages(__spreadArray(__spreadArray([], chatMessages, true), [newMessage], false));
                    setAiProcessing(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 5, , 6]);
                    if (!(props.enableAiSuggestions && props.azureOpenAiApiKey)) return [3 /*break*/, 3];
                    return [4 /*yield*/, AzureAiService.getAiAnswer(question, props.azureOpenAiEndpoint, props.azureOpenAiApiKey)];
                case 2:
                    aiResponse = _a.sent();
                    aiMessage_1 = {
                        id: (Date.now() + 1).toString(),
                        message: '',
                        isUser: false,
                        timestamp: new Date(),
                        aiResponse: {
                            answer: aiResponse.answer,
                            confidence: aiResponse.confidence,
                            sources: aiResponse.sources,
                            suggestedActions: aiResponse.suggestedActions
                        }
                    };
                    setChatMessages(function (prev) { return __spreadArray(__spreadArray([], prev, true), [aiMessage_1], false); });
                    return [3 /*break*/, 4];
                case 3:
                    mockResponse_1 = {
                        id: (Date.now() + 1).toString(),
                        message: '',
                        isUser: false,
                        timestamp: new Date(),
                        aiResponse: {
                            answer: 'I understand your question. However, AI services are not currently configured. Please check with your administrator or browse our FAQ section for answers.',
                            confidence: 0.5,
                            sources: ['Mock Response'],
                            suggestedActions: ['Browse FAQ', 'Contact Support']
                        }
                    };
                    setChatMessages(function (prev) { return __spreadArray(__spreadArray([], prev, true), [mockResponse_1], false); });
                    _a.label = 4;
                case 4: return [3 /*break*/, 6];
                case 5:
                    error_2 = _a.sent();
                    console.error('Error getting AI response:', error_2);
                    errorMessage_1 = {
                        id: (Date.now() + 1).toString(),
                        message: '',
                        isUser: false,
                        timestamp: new Date(),
                        aiResponse: {
                            answer: 'I apologize, but I encountered an error while processing your question. Please try again or contact support.',
                            confidence: 0,
                            sources: ['Error Handler'],
                            suggestedActions: ['Try Again', 'Contact Support']
                        }
                    };
                    setChatMessages(function (prev) { return __spreadArray(__spreadArray([], prev, true), [errorMessage_1], false); });
                    return [3 /*break*/, 6];
                case 6:
                    setAiProcessing(false);
                    return [2 /*return*/];
            }
        });
    }); };
    var columns = [
        {
            key: 'question',
            name: 'Question',
            fieldName: 'question',
            minWidth: 200,
            maxWidth: 400,
            isResizable: true,
            onRender: function (item) { return (React.createElement("span", { style: { cursor: 'pointer', color: '#0078d4' }, onClick: function () {
                    setSelectedItem(item);
                    setShowPanel(true);
                } }, item.question)); }
        },
        {
            key: 'category',
            name: 'Category',
            fieldName: 'category',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true
        },
        {
            key: 'rating',
            name: 'Rating',
            fieldName: 'rating',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true,
            onRender: function (item) { return (React.createElement(Rating, { rating: item.rating, max: 5, readOnly: true })); }
        },
        {
            key: 'viewCount',
            name: 'Views',
            fieldName: 'viewCount',
            minWidth: 60,
            maxWidth: 80,
            isResizable: true
        }
    ];
    var commandBarItems = [
        {
            key: 'addFaq',
            text: 'Add FAQ',
            iconProps: { iconName: 'Add' },
            onClick: function () { return alert('Add FAQ functionality will be implemented'); }
        },
        {
            key: 'askAi',
            text: 'Ask AI',
            iconProps: { iconName: 'Robot' },
            onClick: function () {
                var question = prompt('What would you like to ask?');
                if (question) {
                    handleAskAi(question);
                }
            }
        },
        {
            key: 'analytics',
            text: 'Analytics',
            iconProps: { iconName: 'BarChart4' },
            onClick: function () { return alert('Analytics dashboard will be implemented'); }
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: loadFaqData
        }
    ];
    var renderChatMessages = function () { return (React.createElement(Stack, { tokens: { childrenGap: 10 } },
        chatMessages.map(function (message) {
            var _a, _b, _c;
            return (React.createElement("div", { key: message.id, className: message.isUser ? styles.userMessage : styles.aiMessage }, message.isUser ? (React.createElement(Text, { variant: "medium" }, message.message)) : (React.createElement(Stack, { tokens: { childrenGap: 5 } },
                React.createElement(Text, { variant: "medium" }, (_a = message.aiResponse) === null || _a === void 0 ? void 0 : _a.answer),
                React.createElement(Text, { variant: "small" },
                    "Confidence: ",
                    Math.round((((_b = message.aiResponse) === null || _b === void 0 ? void 0 : _b.confidence) || 0) * 100),
                    "%"),
                ((_c = message.aiResponse) === null || _c === void 0 ? void 0 : _c.sources) && (React.createElement(Text, { variant: "small" },
                    "Sources: ",
                    message.aiResponse.sources.join(', ')))))));
        }),
        aiProcessing && React.createElement(Spinner, { size: SpinnerSize.small, label: "AI is thinking..." }))); };
    return (React.createElement("div", { className: styles.aiPoweredFaq },
        React.createElement(Stack, { tokens: { childrenGap: 20 } },
            React.createElement("h2", null, "AI-Powered FAQ & Knowledge Base"),
            !props.azureOpenAiApiKey && (React.createElement(MessageBar, { messageBarType: MessageBarType.warning }, "Azure OpenAI is not configured. AI features will use mock responses. Please configure Azure OpenAI settings in web part properties.")),
            React.createElement(SearchBox, { placeholder: "Search FAQ or ask a question...", value: searchQuery, onSearch: handleSearch, onChange: function (_, newValue) { return setSearchQuery(newValue || ''); } }),
            React.createElement(Pivot, null,
                React.createElement(PivotItem, { headerText: "FAQ", itemIcon: "Help" },
                    React.createElement(Stack, { tokens: { childrenGap: 15 } },
                        React.createElement(CommandBar, { items: commandBarItems }),
                        loading ? (React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading FAQ items..." })) : (React.createElement(DetailsList, { items: faqItems, columns: columns, selectionMode: SelectionMode.none, setKey: "set", layoutMode: 0, isHeaderVisible: true })))),
                React.createElement(PivotItem, { headerText: "AI Chat", itemIcon: "Robot" },
                    React.createElement(Stack, { tokens: { childrenGap: 15 } },
                        React.createElement(Text, { variant: "mediumPlus" }, "Ask AI Assistant"),
                        renderChatMessages()))),
            React.createElement(Panel, { isOpen: showPanel, type: PanelType.medium, onDismiss: function () { return setShowPanel(false); }, headerText: "FAQ Details" }, selectedItem && (React.createElement(Stack, { tokens: { childrenGap: 15 } },
                React.createElement(Text, { variant: "xLarge" }, selectedItem.question),
                React.createElement(Separator, null),
                React.createElement(Text, { variant: "medium" }, selectedItem.answer),
                React.createElement(Separator, null),
                React.createElement(Text, { variant: "small" },
                    "Category: ",
                    selectedItem.category),
                React.createElement(Text, { variant: "small" },
                    "Tags: ",
                    selectedItem.tags.join(', ')),
                React.createElement(Rating, { rating: selectedItem.rating, max: 5, readOnly: true }),
                React.createElement(Text, { variant: "small" },
                    "Views: ",
                    selectedItem.viewCount),
                React.createElement(Text, { variant: "small" },
                    "Last Updated: ",
                    selectedItem.lastUpdated.toLocaleDateString())))))));
};
export default AiPoweredFaq;
//# sourceMappingURL=AiPoweredFaq.js.map