import { IKnockoutCrudFormWebPartProps } from './KnockoutCrudFormWebPart';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
export interface IKnockoutCrudFormBindingContext extends IKnockoutCrudFormWebPartProps {
    shouter: KnockoutSubscribable<{}>;
    context: IWebPartContext;
}
export interface Employee {
    intID: number;
    FirstName: string;
    Gender: string;
    DOB: string;
}
export default class KnockoutCrudFormViewModel {
    strDescription: KnockoutObservable<string>;
    firstname: KnockoutObservable<string>;
    lastname: KnockoutObservable<string>;
    gender: KnockoutObservable<string>;
    dob: KnockoutObservable<string>;
    availableGenders: KnockoutObservableArray<string>;
    lstEmployees: KnockoutObservableArray<Employee>;
    intEditNumber: KnockoutObservable<number>;
    knockoutCrudFormClass: string;
    containerClass: string;
    rowClass: string;
    columnClass: string;
    titleClass: string;
    subTitleClass: string;
    descriptionClass: string;
    firstnameClass: string;
    genderClass: string;
    dobClass: string;
    buttonClass: string;
    labelClass: string;
    context: any;
    strFileName: string;
    tempEmployee: Employee;
    constructor(bindings: IKnockoutCrudFormBindingContext);
    private getItems;
    private ensureList;
    addItem(): void;
    deleteItem(data: any): void;
    editItem(data: any): Promise<void>;
}
//# sourceMappingURL=KnockoutCrudFormViewModel.d.ts.map