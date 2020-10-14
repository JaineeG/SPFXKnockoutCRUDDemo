import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IKnockoutCrudFormWebPartProps {
    firstname: string;
    strDescription: string;
    gender: string;
    dob: string;
    intEditNumber: number;
}
export default class KnockoutCrudFormWebPart extends BaseClientSideWebPart<IKnockoutCrudFormWebPartProps> {
    private id;
    private intKOEditNumber;
    private strKoDescription;
    private componentElement;
    private koFirstName;
    private koGender;
    private koDOB;
    private lists;
    strFileName: string;
    /**
     * Shouter is used to communicate between web part and view model.
     */
    private _shouter;
    /**
     * Initialize the web part.
     */
    protected onInit(): Promise<void>;
    render(): void;
    private _createComponentElement;
    private _registerComponent;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=KnockoutCrudFormWebPart.d.ts.map