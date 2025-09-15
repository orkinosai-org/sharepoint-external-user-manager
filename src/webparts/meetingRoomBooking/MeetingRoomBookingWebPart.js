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
import { PropertyPaneTextField, PropertyPaneToggle, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'MeetingRoomBookingWebPartStrings';
import MeetingRoomBooking from './components/MeetingRoomBooking';
var MeetingRoomBookingWebPart = /** @class */ (function (_super) {
    __extends(MeetingRoomBookingWebPart, _super);
    function MeetingRoomBookingWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    MeetingRoomBookingWebPart.prototype.render = function () {
        var element = React.createElement(MeetingRoomBooking, {
            description: this.properties.description,
            context: this.context,
            defaultLocation: this.properties.defaultLocation,
            enableTeamsIntegration: this.properties.enableTeamsIntegration,
            enableRecurringBookings: this.properties.enableRecurringBookings,
            maxBookingDuration: this.properties.maxBookingDuration,
            requireApproval: this.properties.requireApproval
        });
        ReactDom.render(element, this.domElement);
    };
    MeetingRoomBookingWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(MeetingRoomBookingWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    MeetingRoomBookingWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                }),
                                PropertyPaneTextField('defaultLocation', {
                                    label: strings.DefaultLocationLabel
                                })
                            ]
                        },
                        {
                            groupName: strings.BookingSettingsGroupName,
                            groupFields: [
                                PropertyPaneToggle('enableTeamsIntegration', {
                                    label: strings.EnableTeamsIntegrationLabel,
                                    onText: 'Enabled',
                                    offText: 'Disabled'
                                }),
                                PropertyPaneToggle('enableRecurringBookings', {
                                    label: strings.EnableRecurringBookingsLabel,
                                    onText: 'Enabled',
                                    offText: 'Disabled'
                                }),
                                PropertyPaneDropdown('maxBookingDuration', {
                                    label: strings.MaxBookingDurationLabel,
                                    options: [
                                        { key: 2, text: '2 hours' },
                                        { key: 4, text: '4 hours' },
                                        { key: 8, text: '8 hours' },
                                        { key: 24, text: '1 day' }
                                    ]
                                }),
                                PropertyPaneToggle('requireApproval', {
                                    label: strings.RequireApprovalLabel,
                                    onText: 'Required',
                                    offText: 'Not Required'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return MeetingRoomBookingWebPart;
}(BaseClientSideWebPart));
export default MeetingRoomBookingWebPart;
//# sourceMappingURL=MeetingRoomBookingWebPart.js.map