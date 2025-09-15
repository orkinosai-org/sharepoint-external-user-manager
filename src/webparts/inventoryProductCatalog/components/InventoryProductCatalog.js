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
import { Stack, CommandBar, DetailsList, SelectionMode, Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';
import { MockInventoryDataService } from '../services/MockInventoryDataService';
import styles from './InventoryProductCatalog.module.scss';
var InventoryProductCatalog = function (props) {
    var _a = useState([]), products = _a[0], setProducts = _a[1];
    var _b = useState([]), stockAlerts = _b[0], setStockAlerts = _b[1];
    var _c = useState(true), loading = _c[0], setLoading = _c[1];
    useEffect(function () {
        loadInventoryData();
    }, []);
    var loadInventoryData = function () { return __awaiter(void 0, void 0, void 0, function () {
        return __generator(this, function (_a) {
            setLoading(true);
            try {
                // Simulate API call
                setTimeout(function () {
                    var mockProducts = MockInventoryDataService.getProducts();
                    var mockAlerts = MockInventoryDataService.getStockAlerts();
                    setProducts(mockProducts);
                    setStockAlerts(mockAlerts);
                    setLoading(false);
                }, 1000);
            }
            catch (error) {
                console.error('Error loading inventory data:', error);
                setLoading(false);
            }
            return [2 /*return*/];
        });
    }); };
    var columns = [
        {
            key: 'name',
            name: 'Product Name',
            fieldName: 'name',
            minWidth: 150,
            maxWidth: 200,
            isResizable: true
        },
        {
            key: 'sku',
            name: 'SKU',
            fieldName: 'sku',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true
        },
        {
            key: 'category',
            name: 'Category',
            fieldName: 'category',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true
        },
        {
            key: 'currentStock',
            name: 'Current Stock',
            fieldName: 'currentStock',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true
        },
        {
            key: 'price',
            name: 'Price',
            fieldName: 'price',
            minWidth: 80,
            maxWidth: 100,
            isResizable: true,
            onRender: function (item) { return "".concat(item.currency).concat(item.price.toFixed(2)); }
        },
        {
            key: 'status',
            name: 'Status',
            fieldName: 'status',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true
        }
    ];
    var commandBarItems = [
        {
            key: 'addProduct',
            text: 'Add Product',
            iconProps: { iconName: 'Add' },
            onClick: function () { return alert('Add Product functionality will be implemented'); }
        },
        {
            key: 'stockIn',
            text: 'Stock In',
            iconProps: { iconName: 'Upload' },
            onClick: function () { return alert('Stock In functionality will be implemented'); }
        },
        {
            key: 'stockOut',
            text: 'Stock Out',
            iconProps: { iconName: 'Download' },
            onClick: function () { return alert('Stock Out functionality will be implemented'); }
        },
        {
            key: 'generateReport',
            text: 'Generate Report',
            iconProps: { iconName: 'BarChart4' },
            onClick: function () { return alert('Generate Report functionality will be implemented'); }
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: loadInventoryData
        }
    ];
    var renderStockAlerts = function () {
        if (stockAlerts.length === 0)
            return null;
        return (React.createElement(Stack, { tokens: { childrenGap: 10 } },
            React.createElement("h3", null, "Stock Alerts"),
            stockAlerts.map(function (alert) { return (React.createElement(MessageBar, { key: alert.id, messageBarType: alert.alertType === 'Out of Stock' ? MessageBarType.severeWarning : MessageBarType.warning, isMultiline: false },
                alert.productName,
                ": ",
                alert.alertType,
                " - Current: ",
                alert.currentStock,
                ", Threshold: ",
                alert.thresholdStock)); })));
    };
    return (React.createElement("div", { className: styles.inventoryProductCatalog },
        React.createElement(Stack, { tokens: { childrenGap: 20 } },
            React.createElement("h2", null, "Inventory & Product Catalog"),
            renderStockAlerts(),
            React.createElement(CommandBar, { items: commandBarItems }),
            loading ? (React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading inventory data..." })) : (React.createElement(DetailsList, { items: products, columns: columns, selectionMode: SelectionMode.multiple, setKey: "set", layoutMode: 0, isHeaderVisible: true })))));
};
export default InventoryProductCatalog;
//# sourceMappingURL=InventoryProductCatalog.js.map