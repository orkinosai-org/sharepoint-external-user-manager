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
var AzureAiService = /** @class */ (function () {
    function AzureAiService() {
    }
    /**
     * Get AI-powered answer suggestions using Azure OpenAI
     */
    AzureAiService.getAnswerSuggestions = function (query, endpoint, apiKey) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                try {
                    // Mock implementation - would integrate with Azure OpenAI API
                    if (!endpoint || !apiKey) {
                        return [2 /*return*/, this.getMockSuggestions(query)];
                    }
                    // Real Azure OpenAI integration would look like:
                    /*
                    const response = await fetch(`${endpoint}/openai/deployments/gpt-35-turbo/chat/completions?api-version=2023-12-01-preview`, {
                      method: 'POST',
                      headers: {
                        'Content-Type': 'application/json',
                        'api-key': apiKey
                      },
                      body: JSON.stringify({
                        messages: [
                          {
                            role: 'system',
                            content: 'You are a helpful assistant that provides FAQ answers based on company knowledge base.'
                          },
                          {
                            role: 'user',
                            content: query
                          }
                        ],
                        max_tokens: 500,
                        temperature: 0.7
                      })
                    });
              
                    const data = await response.json();
                    return this.processAzureResponse(data);
                    */
                    // For now, return mock suggestions
                    return [2 /*return*/, this.getMockSuggestions(query)];
                }
                catch (error) {
                    console.error('Error getting AI suggestions:', error);
                    return [2 /*return*/, this.getMockSuggestions(query)];
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Get direct AI answer for a question
     */
    AzureAiService.getAiAnswer = function (question, endpoint, apiKey) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                try {
                    // Mock implementation - would integrate with Azure OpenAI API
                    if (!endpoint || !apiKey) {
                        return [2 /*return*/, this.getMockAiResponse(question)];
                    }
                    // Real Azure OpenAI integration would look like:
                    /*
                    const response = await fetch(`${endpoint}/openai/deployments/gpt-35-turbo/chat/completions?api-version=2023-12-01-preview`, {
                      method: 'POST',
                      headers: {
                        'Content-Type': 'application/json',
                        'api-key': apiKey
                      },
                      body: JSON.stringify({
                        messages: [
                          {
                            role: 'system',
                            content: 'You are a helpful company assistant. Provide accurate, helpful answers based on company policies and procedures. If you don\'t know something, say so clearly.'
                          },
                          {
                            role: 'user',
                            content: question
                          }
                        ],
                        max_tokens: 800,
                        temperature: 0.5
                      })
                    });
              
                    const data = await response.json();
                    return this.processAiResponse(data);
                    */
                    // For now, return mock response
                    return [2 /*return*/, this.getMockAiResponse(question)];
                }
                catch (error) {
                    console.error('Error getting AI answer:', error);
                    return [2 /*return*/, {
                            answer: 'I apologize, but I encountered an error while processing your question. Please try again later or contact support.',
                            confidence: 0,
                            sources: ['Error Handler'],
                            suggestedActions: ['Try Again', 'Contact Support']
                        }];
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Analyze FAQ content and suggest improvements
     */
    AzureAiService.analyzeFaqContent = function (faqItems) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                try {
                    // Mock implementation - would use Azure Cognitive Services for text analytics
                    return [2 /*return*/, {
                            insights: [
                                'Consider adding more technical FAQ items',
                                'HR category needs more detailed answers',
                                'Add visual content to complex procedures'
                            ],
                            suggestions: [
                                'Group related questions together',
                                'Add step-by-step guides for complex processes',
                                'Include more examples in answers'
                            ],
                            topTopics: ['password reset', 'time off', 'software installation'],
                            sentimentScore: 0.85
                        }];
                }
                catch (error) {
                    console.error('Error analyzing FAQ content:', error);
                    return [2 /*return*/, null];
                }
                return [2 /*return*/];
            });
        });
    };
    AzureAiService.getMockSuggestions = function (query) {
        var suggestions = [];
        // Generate mock suggestions based on query content
        if (query.toLowerCase().includes('password')) {
            suggestions.push({
                id: '1',
                suggestion: 'Consider adding information about password complexity requirements',
                confidence: 0.85,
                reasoning: 'Users often ask about password rules after reset procedures',
                source: 'azure-ai'
            });
        }
        if (query.toLowerCase().includes('time off') || query.toLowerCase().includes('vacation')) {
            suggestions.push({
                id: '2',
                suggestion: 'Include information about advance notice requirements for time off requests',
                confidence: 0.78,
                reasoning: 'This is commonly asked follow-up question',
                source: 'cognitive-search'
            });
        }
        suggestions.push({
            id: '3',
            suggestion: 'Add related articles or links to comprehensive guides',
            confidence: 0.65,
            reasoning: 'Providing additional resources improves user satisfaction',
            source: 'custom-model'
        });
        return suggestions;
    };
    AzureAiService.getMockAiResponse = function (question) {
        // Generate different responses based on question content
        if (question.toLowerCase().includes('password')) {
            return {
                answer: 'To reset your password, go to the company login page and click "Forgot Password". You\'ll receive an email with reset instructions. New passwords must be at least 8 characters long and include uppercase, lowercase, numbers, and special characters.',
                confidence: 0.92,
                sources: ['IT Policy Document', 'User Manual'],
                suggestedActions: ['Go to Login Page', 'Contact IT Support', 'View Password Policy']
            };
        }
        if (question.toLowerCase().includes('time off') || question.toLowerCase().includes('vacation')) {
            return {
                answer: 'To request time off, use the HR portal to submit your request at least 2 weeks in advance. Include the dates, reason, and coverage plan. Your manager will review and approve the request.',
                confidence: 0.88,
                sources: ['HR Policy Manual', 'Employee Handbook'],
                suggestedActions: ['Open HR Portal', 'Check Time Off Balance', 'Contact HR']
            };
        }
        if (question.toLowerCase().includes('software') || question.toLowerCase().includes('install')) {
            return {
                answer: 'Only pre-approved software can be installed on company computers. Check the approved software catalog in the IT portal. For new software requests, submit a ticket to IT with business justification.',
                confidence: 0.85,
                sources: ['IT Security Policy', 'Software Catalog'],
                suggestedActions: ['View Software Catalog', 'Submit IT Ticket', 'Contact IT Support']
            };
        }
        // Default response
        return {
            answer: 'I understand your question, but I need more specific information to provide a detailed answer. Please check our FAQ section or contact the appropriate department for assistance.',
            confidence: 0.45,
            sources: ['General Knowledge Base'],
            suggestedActions: ['Browse FAQ', 'Contact Support', 'Refine Question']
        };
    };
    AzureAiService.processAzureResponse = function (data) {
        // Process actual Azure OpenAI response
        // This would parse the response and extract suggestions
        return [];
    };
    AzureAiService.processAiResponse = function (data) {
        var _a, _b;
        // Process actual Azure OpenAI response
        // This would parse the response and extract the answer
        return {
            answer: ((_b = (_a = data.choices[0]) === null || _a === void 0 ? void 0 : _a.message) === null || _b === void 0 ? void 0 : _b.content) || 'No response available',
            confidence: 0.8,
            sources: ['Azure OpenAI'],
            suggestedActions: []
        };
    };
    return AzureAiService;
}());
export { AzureAiService };
//# sourceMappingURL=AzureAiService.js.map