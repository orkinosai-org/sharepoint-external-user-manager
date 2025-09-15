var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'AiPoweredFaqWebPartStrings';
import AiPoweredFaq from './components/AiPoweredFaq';
var AiPoweredFaqWebPart = /** @class */ (function (_super) {
    __extends(AiPoweredFaqWebPart, _super);
    function AiPoweredFaqWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    AiPoweredFaqWebPart.prototype.render = function () {
        var element = React.createElement(AiPoweredFaq, {
            description: this.properties.description,
            context: this.context,
            azureOpenAiEndpoint: this.properties.azureOpenAiEndpoint,
            azureOpenAiApiKey: this.properties.azureOpenAiApiKey,
            enableAiSuggestions: this.properties.enableAiSuggestions,
            enableAnalytics: this.properties.enableAnalytics
        });
        ReactDom.render(element, this.domElement);
    };
    AiPoweredFaqWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(AiPoweredFaqWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    AiPoweredFaqWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        },
                        {
                            groupName: strings.AzureAiGroupName,
                            groupFields: [
                                PropertyPaneTextField('azureOpenAiEndpoint', {
                                    label: strings.AzureOpenAiEndpointLabel
                                }),
                                PropertyPaneTextField('azureOpenAiApiKey', {
                                    label: strings.AzureOpenAiApiKeyLabel
                                }),
                                PropertyPaneToggle('enableAiSuggestions', {
                                    label: strings.EnableAiSuggestionsLabel,
                                    onText: 'Enabled',
                                    offText: 'Disabled'
                                }),
                                PropertyPaneToggle('enableAnalytics', {
                                    label: strings.EnableAnalyticsLabel,
                                    onText: 'Enabled',
                                    offText: 'Disabled'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return AiPoweredFaqWebPart;
}(BaseClientSideWebPart));
export default AiPoweredFaqWebPart;
//# sourceMappingURL=AiPoweredFaqWebPart.js.map