var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as ko from 'knockout';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'KnockoutCrudFormWebPartStrings';
import KnockoutCrudFormViewModel from './KnockoutCrudFormViewModel';
import { sp } from "@pnp/sp";
var _instance = 0;
var KnockoutCrudFormWebPart = /** @class */ (function (_super) {
    __extends(KnockoutCrudFormWebPart, _super);
    function KnockoutCrudFormWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.intKOEditNumber = ko.observable();
        _this.strKoDescription = ko.observable();
        _this.koFirstName = ko.observable('');
        _this.koGender = ko.observable('');
        _this.koDOB = ko.observable('');
        _this.strFileName = "KnockoutCrudFormWebPart";
        /**
         * Shouter is used to communicate between web part and view model.
         */
        _this._shouter = new ko.subscribable();
        return _this;
    }
    /**
     * Initialize the web part.
     */
    KnockoutCrudFormWebPart.prototype.onInit = function () {
        var _this = this;
        /// <summary>onInit function</summary>
        try {
            this.id = _instance++;
            var tagName = "ComponentElement-" + this.id;
            this.componentElement = this._createComponentElement(tagName);
            this._registerComponent(tagName);
            this.intKOEditNumber.subscribe(function (newValue) {
                _this._shouter.notifySubscribers(newValue, 'intEditNumber');
            });
            this.strKoDescription.subscribe(function (newValue) {
                _this._shouter.notifySubscribers(newValue, 'strDescription');
            });
            this.koFirstName.subscribe(function (newValue) {
                _this._shouter.notifySubscribers(newValue, 'firstname');
            });
            this.koDOB.subscribe(function (newValue) {
                _this._shouter.notifySubscribers(newValue, 'dob');
            });
            this.koGender.subscribe(function (newValue) {
                _this._shouter.notifySubscribers(newValue, 'gender');
            });
            var bindings = {
                firstname: this.properties.firstname,
                intEditNumber: this.properties.intEditNumber,
                strDescription: this.properties.strDescription,
                gender: this.properties.gender,
                dob: this.properties.dob,
                context: this.context,
                shouter: this._shouter
            };
            ko.applyBindings(bindings, this.componentElement);
            sp.setup({
                spfxContext: this.context
            });
            return _super.prototype.onInit.call(this);
        }
        catch (Exception) {
            console.log(this.strFileName + " onInit() : " + Exception.message);
        }
    };
    KnockoutCrudFormWebPart.prototype.render = function () {
        /// <summary>render function</summary>
        try {
            if (!this.renderedOnce) {
                this.domElement.appendChild(this.componentElement);
            }
            this.strKoDescription(this.properties.strDescription);
            this.intKOEditNumber(this.properties.intEditNumber);
            this.koFirstName(this.properties.firstname);
            this.koGender(this.properties.gender);
            this.koDOB(this.properties.dob);
        }
        catch (Exception) {
            console.log(this.strFileName + " render() : " + Exception.message);
        }
    };
    KnockoutCrudFormWebPart.prototype._createComponentElement = function (tagName) {
        /// <summary>_createComponentElement function</summary>
        /// <param name="tagName">TagName of the HTML element</param>
        try {
            var componentElement = document.createElement('div');
            componentElement.setAttribute('data-bind', "component: { name: \"" + tagName + "\", params: $data }");
            return componentElement;
        }
        catch (Exception) {
            console.log(this.strFileName + " _createComponentElement() : " + Exception.message);
        }
    };
    KnockoutCrudFormWebPart.prototype._registerComponent = function (tagName) {
        /// <summary>_registerComponent function</summary>
        /// <param name="tagName">TagName of the HTML Components</param>
        try {
            ko.components.register(tagName, {
                viewModel: KnockoutCrudFormViewModel,
                template: require('./KnockoutCrudForm.template.html'),
                synchronous: false
            });
        }
        catch (Exception) {
            console.log(this.strFileName + " _registerComponent() : " + Exception.message);
        }
    };
    Object.defineProperty(KnockoutCrudFormWebPart.prototype, "dataVersion", {
        get: function () {
            /// <summary>dataVersion function</summary>
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    KnockoutCrudFormWebPart.prototype.getPropertyPaneConfiguration = function () {
        /// <summary>this function is used to add the UI for Property Pane.</summary>
        try {
            return {
                pages: [
                    {
                        groups: [
                            {
                                groupFields: [
                                    PropertyPaneTextField('strDescription', {
                                        label: strings.DescriptionFieldLabel
                                    }),
                                ]
                            }
                        ]
                    }
                ]
            };
        }
        catch (Exception) {
            console.log(this.strFileName + " getPropertyPaneConfiguration() : " + Exception.message);
        }
    };
    return KnockoutCrudFormWebPart;
}(BaseClientSideWebPart));
export default KnockoutCrudFormWebPart;
//# sourceMappingURL=KnockoutCrudFormWebPart.js.map