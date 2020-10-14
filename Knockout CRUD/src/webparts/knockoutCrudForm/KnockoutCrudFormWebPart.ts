import * as ko from 'knockout';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'KnockoutCrudFormWebPartStrings';
import KnockoutCrudFormViewModel, { IKnockoutCrudFormBindingContext } from './KnockoutCrudFormViewModel';
import { sp } from "@pnp/sp";
let _instance: number = 0;

export interface IKnockoutCrudFormWebPartProps {
  firstname: string;
  strDescription: string;
  gender: string;
  dob: string;
  intEditNumber: number;
}

export default class KnockoutCrudFormWebPart extends BaseClientSideWebPart<IKnockoutCrudFormWebPartProps> {
  /// <summary>KnockoutCrudFormWebPart class</summary>
  /// <param name="IKnockoutCrudFormWebPartProps"></param>
  private id: number;
  private intKOEditNumber: KnockoutObservable<Number> = ko.observable();
  private strKoDescription: KnockoutObservable<String> = ko.observable();
  private componentElement: HTMLElement;
  private koFirstName: KnockoutObservable<string> = ko.observable('');
  private koGender: KnockoutObservable<string> = ko.observable('');
  private koDOB: KnockoutObservable<string> = ko.observable('');
  private lists: IPropertyPaneDropdownOption[];
  public strFileName: string = "KnockoutCrudFormWebPart";
  /**
   * Shouter is used to communicate between web part and view model.
   */
  private _shouter: KnockoutSubscribable<{}> = new ko.subscribable();

  /**
   * Initialize the web part.
   */
  protected onInit(): Promise<void> {
    /// <summary>onInit function</summary>
    try {
      this.id = _instance++;

      const tagName: string = `ComponentElement-${this.id}`;
      this.componentElement = this._createComponentElement(tagName);
      this._registerComponent(tagName);

      this.intKOEditNumber.subscribe((newValue: Number) => {
        this._shouter.notifySubscribers(newValue, 'intEditNumber');
      });

      this.strKoDescription.subscribe((newValue: string) => {
        this._shouter.notifySubscribers(newValue, 'strDescription');
      });

      this.koFirstName.subscribe((newValue: string) => {
        this._shouter.notifySubscribers(newValue, 'firstname');
      });

      this.koDOB.subscribe((newValue: string) => {
        this._shouter.notifySubscribers(newValue, 'dob');
      });

      this.koGender.subscribe((newValue: string) => {
        this._shouter.notifySubscribers(newValue, 'gender');
      });

      const bindings: IKnockoutCrudFormBindingContext = {
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

      return super.onInit();
    } catch (Exception) {
      console.log(this.strFileName + " onInit() : " + Exception.message);
    }
  }

  public render(): void {
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
    } catch (Exception) {
      console.log(this.strFileName + " render() : " + Exception.message);
    }
  }

  private _createComponentElement(tagName: string): HTMLElement {
    /// <summary>_createComponentElement function</summary>
    /// <param name="tagName">TagName of the HTML element</param>
    try {
      const componentElement: HTMLElement = document.createElement('div');
      componentElement.setAttribute('data-bind', `component: { name: "${tagName}", params: $data }`);
      return componentElement;
    } catch (Exception) {
      console.log(this.strFileName + " _createComponentElement() : " + Exception.message);
    }
  }

  private _registerComponent(tagName: string): void {
    /// <summary>_registerComponent function</summary>
    /// <param name="tagName">TagName of the HTML Components</param>
    try {
      ko.components.register(
        tagName,
        {
          viewModel: KnockoutCrudFormViewModel,
          template: require('./KnockoutCrudForm.template.html'),
          synchronous: false
        }
      );
    } catch (Exception) {
      console.log(this.strFileName + " _registerComponent() : " + Exception.message);
    }
  }

  protected get dataVersion(): Version {
    /// <summary>dataVersion function</summary>
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
    } catch (Exception) {
      console.log(this.strFileName + " getPropertyPaneConfiguration() : " + Exception.message);
    }
  }
}
