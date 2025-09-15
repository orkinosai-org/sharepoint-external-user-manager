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
var MockFaqDataService = /** @class */ (function () {
    function MockFaqDataService() {
    }
    MockFaqDataService.getFaqItems = function () {
        return [
            {
                id: '1',
                question: 'How do I reset my password?',
                answer: 'To reset your password, go to the login page and click "Forgot Password". Enter your email address and follow the instructions sent to your email.',
                category: 'IT Support',
                tags: ['password', 'login', 'security'],
                rating: 4.5,
                viewCount: 245,
                lastUpdated: new Date('2024-01-15'),
                createdBy: 'admin@company.com',
                aiGenerated: false,
                relatedQuestions: ['How to change password?', 'Account locked out']
            },
            {
                id: '2',
                question: 'What are the company holidays for 2024?',
                answer: 'Company holidays for 2024 include New Year\'s Day, Martin Luther King Jr. Day, Presidents\' Day, Memorial Day, Independence Day, Labor Day, Thanksgiving, and Christmas Day. Please refer to the HR portal for the complete list.',
                category: 'HR',
                tags: ['holidays', 'time-off', 'calendar'],
                rating: 4.8,
                viewCount: 189,
                lastUpdated: new Date('2024-01-01'),
                createdBy: 'hr@company.com',
                aiGenerated: false,
                relatedQuestions: ['How to request time off?', 'Vacation policy']
            },
            {
                id: '3',
                question: 'How do I submit an expense report?',
                answer: 'To submit an expense report, log into the expense management system, click "New Report", add your expenses with receipts, and submit for approval. Reports must be submitted within 30 days of the expense date.',
                category: 'Billing',
                tags: ['expenses', 'reimbursement', 'finance'],
                rating: 4.2,
                viewCount: 156,
                lastUpdated: new Date('2024-01-10'),
                createdBy: 'finance@company.com',
                aiGenerated: false,
                relatedQuestions: ['Expense policy', 'Receipt requirements']
            },
            {
                id: '4',
                question: 'What software can I install on my work computer?',
                answer: 'Only approved software from the company catalog can be installed. Contact IT for software requests or to add new software to the approved list. Personal software installation is not permitted for security reasons.',
                category: 'IT Support',
                tags: ['software', 'installation', 'security', 'policy'],
                rating: 4.0,
                viewCount: 201,
                lastUpdated: new Date('2024-01-08'),
                createdBy: 'it@company.com',
                aiGenerated: false,
                relatedQuestions: ['Software approval process', 'Security policies']
            },
            {
                id: '5',
                question: 'How do I access SharePoint from home?',
                answer: 'To access SharePoint from home, use your company credentials to log in through the web portal at portal.company.com. Ensure you have VPN access if required. For mobile access, download the SharePoint mobile app.',
                category: 'Technical',
                tags: ['sharepoint', 'remote-access', 'vpn', 'mobile'],
                rating: 4.6,
                viewCount: 312,
                lastUpdated: new Date('2024-01-12'),
                createdBy: 'it@company.com',
                aiGenerated: true,
                confidence: 0.89,
                relatedQuestions: ['VPN setup', 'Mobile app installation']
            }
        ];
    };
    MockFaqDataService.searchFaqItems = function (query) {
        return __awaiter(this, void 0, void 0, function () {
            var allItems;
            return __generator(this, function (_a) {
                allItems = this.getFaqItems();
                return [2 /*return*/, allItems.filter(function (item) {
                        return item.question.toLowerCase().includes(query.toLowerCase()) ||
                            item.answer.toLowerCase().includes(query.toLowerCase()) ||
                            item.tags.some(function (tag) { return tag.toLowerCase().includes(query.toLowerCase()); });
                    })];
            });
        });
    };
    MockFaqDataService.createFaqItem = function (faq) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                return [2 /*return*/, {
                        id: Math.random().toString(36).substr(2, 9),
                        question: faq.question || '',
                        answer: faq.answer || '',
                        category: faq.category || 'General',
                        tags: faq.tags || [],
                        rating: 0,
                        viewCount: 0,
                        lastUpdated: new Date(),
                        createdBy: 'user@company.com',
                        aiGenerated: false
                    }];
            });
        });
    };
    MockFaqDataService.updateFaqItem = function (id, updates) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                console.log("Updating FAQ item ".concat(id), updates);
                return [2 /*return*/];
            });
        });
    };
    MockFaqDataService.deleteFaqItem = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                console.log("Deleting FAQ item ".concat(id));
                return [2 /*return*/];
            });
        });
    };
    MockFaqDataService.incrementViewCount = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would track analytics
                console.log("Incrementing view count for FAQ ".concat(id));
                return [2 /*return*/];
            });
        });
    };
    MockFaqDataService.rateFaqItem = function (id, rating) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would store user ratings
                console.log("Rating FAQ ".concat(id, " with ").concat(rating, " stars"));
                return [2 /*return*/];
            });
        });
    };
    return MockFaqDataService;
}());
export { MockFaqDataService };
//# sourceMappingURL=MockFaqDataService.js.map