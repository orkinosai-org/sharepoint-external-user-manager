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
var MockInventoryDataService = /** @class */ (function () {
    function MockInventoryDataService() {
    }
    MockInventoryDataService.getProducts = function () {
        return [
            {
                id: '1',
                name: 'Wireless Bluetooth Headphones',
                sku: 'WBH-001',
                category: 'Electronics',
                description: 'High-quality wireless headphones with noise cancellation',
                price: 199.99,
                currency: '$',
                currentStock: 25,
                minimumStock: 10,
                maximumStock: 100,
                supplier: 'TechSupplier Inc.',
                lastRestocked: new Date('2024-01-10'),
                status: 'Active',
                location: 'Warehouse A - Shelf 1'
            },
            {
                id: '2',
                name: 'Ergonomic Office Chair',
                sku: 'EOC-002',
                category: 'Furniture',
                description: 'Adjustable ergonomic chair with lumbar support',
                price: 299.99,
                currency: '$',
                currentStock: 5,
                minimumStock: 8,
                maximumStock: 50,
                supplier: 'Office Solutions Ltd.',
                lastRestocked: new Date('2024-01-05'),
                status: 'Low Stock',
                location: 'Warehouse B - Section 2'
            },
            {
                id: '3',
                name: 'Laptop Stand Aluminum',
                sku: 'LSA-003',
                category: 'Accessories',
                description: 'Adjustable aluminum laptop stand for better ergonomics',
                price: 49.99,
                currency: '$',
                currentStock: 0,
                minimumStock: 15,
                maximumStock: 75,
                supplier: 'Accessories Plus',
                lastRestocked: new Date('2023-12-20'),
                status: 'Out of Stock',
                location: 'Warehouse A - Shelf 3'
            },
            {
                id: '4',
                name: 'Smart Water Bottle',
                sku: 'SWB-004',
                category: 'Health & Wellness',
                description: 'Smart water bottle with hydration tracking',
                price: 79.99,
                currency: '$',
                currentStock: 45,
                minimumStock: 20,
                maximumStock: 80,
                supplier: 'Health Tech Co.',
                lastRestocked: new Date('2024-01-15'),
                status: 'Active',
                location: 'Warehouse C - Zone 1'
            }
        ];
    };
    MockInventoryDataService.getStockAlerts = function () {
        return [
            {
                id: '1',
                productId: '2',
                productName: 'Ergonomic Office Chair',
                alertType: 'Low Stock',
                currentStock: 5,
                thresholdStock: 8,
                dateGenerated: new Date('2024-01-16'),
                acknowledged: false
            },
            {
                id: '2',
                productId: '3',
                productName: 'Laptop Stand Aluminum',
                alertType: 'Out of Stock',
                currentStock: 0,
                thresholdStock: 15,
                dateGenerated: new Date('2024-01-14'),
                acknowledged: false
            }
        ];
    };
    MockInventoryDataService.getInventoryTransactions = function () {
        return [
            {
                id: '1',
                productId: '1',
                type: 'Stock In',
                quantity: 50,
                date: new Date('2024-01-10'),
                reference: 'PO-2024-001',
                notes: 'Monthly restock from supplier',
                performedBy: 'warehouse@company.com'
            },
            {
                id: '2',
                productId: '1',
                type: 'Stock Out',
                quantity: 25,
                date: new Date('2024-01-15'),
                reference: 'SO-2024-005',
                notes: 'Sales order fulfillment',
                performedBy: 'sales@company.com'
            }
        ];
    };
    MockInventoryDataService.createProduct = function (product) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                return [2 /*return*/, {
                        id: Math.random().toString(36).substr(2, 9),
                        name: product.name || '',
                        sku: product.sku || '',
                        category: product.category || '',
                        description: product.description || '',
                        price: product.price || 0,
                        currency: product.currency || '$',
                        currentStock: product.currentStock || 0,
                        minimumStock: product.minimumStock || 0,
                        maximumStock: product.maximumStock || 0,
                        supplier: product.supplier || '',
                        lastRestocked: new Date(),
                        status: 'Active',
                        location: product.location || ''
                    }];
            });
        });
    };
    MockInventoryDataService.updateStock = function (productId, quantity, type) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would integrate with SharePoint Lists or external API
                console.log("Updating stock for product ".concat(productId, ": ").concat(type, " ").concat(quantity, " units"));
                return [2 /*return*/];
            });
        });
    };
    MockInventoryDataService.generateInventoryReport = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                // Mock implementation - would generate actual report
                return [2 /*return*/, 'inventory-report-' + new Date().toISOString().split('T')[0] + '.xlsx'];
            });
        });
    };
    return MockInventoryDataService;
}());
export { MockInventoryDataService };
//# sourceMappingURL=MockInventoryDataService.js.map