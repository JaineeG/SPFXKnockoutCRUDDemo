import * as ko from 'knockout';
import styles from './KnockoutCrudForm.module.scss';
import { IKnockoutCrudFormWebPartProps } from './KnockoutCrudFormWebPart';
import { sp, List, ItemAddResult, ListEnsureResult } from "sp-pnp-js";
import { IWebPartContext } from '@microsoft/sp-webpart-base';
export interface IKnockoutCrudFormBindingContext extends IKnockoutCrudFormWebPartProps {
  shouter: KnockoutSubscribable<{}>;
  context: IWebPartContext;
}

export interface Employee {
  /// <summary>Employee interface based on list columns</summary>
  intID: number;
  FirstName: string;
  Gender: string;
  DOB: string;
}

export default class KnockoutCrudFormViewModel {
  /// <summary>KnockoutCrudFormViewModel class</summary>

  public strDescription: KnockoutObservable<string> = ko.observable('');
  public firstname: KnockoutObservable<string> = ko.observable('');
  public lastname: KnockoutObservable<string> = ko.observable('');
  public gender: KnockoutObservable<string> = ko.observable('');
  public dob: KnockoutObservable<string> = ko.observable('');
  public availableGenders: KnockoutObservableArray<string> = ko.observableArray(['Select', 'Male', 'Female', 'Others']);
  public lstEmployees: KnockoutObservableArray<Employee> = ko.observableArray([]);
  public intEditNumber: KnockoutObservable<number> = ko.observable();

  public knockoutCrudFormClass: string = styles.knockoutCrudForm;
  public containerClass: string = styles.container;
  public rowClass: string = styles.row;
  public columnClass: string = styles.column;
  public titleClass: string = styles.title;
  public subTitleClass: string = styles.subTitle;
  public descriptionClass: string = styles.strDescription;
  public firstnameClass: string = styles.firstname;
  public genderClass: string = styles.gender;
  public dobClass: string = styles.dob;
  public buttonClass: string = styles.button;
  public labelClass: string = styles.label;
  public context: any;
  public strFileName: string = "KnockoutCrudFormViewModel";
  // used for storing old value while updating an item
  public tempEmployee: Employee;

  constructor(bindings: IKnockoutCrudFormBindingContext) {
    /// <summary>constructor</summary>
    /// <param name="IKnockoutCrudFormBindingContext">bindings values</param>

    this.firstname(bindings.firstname);
    this.gender(bindings.gender);
    this.dob(bindings.dob);
    this.intEditNumber(bindings.intEditNumber);
    this.strDescription(bindings.strDescription);
    this.context = bindings.context;

    // When web part fields is updated, change this view model's values
    bindings.shouter.subscribe((value: number) => {
      this.intEditNumber(value);
    }, this, 'intEditNumber');
    this.intEditNumber(0);
    bindings.shouter.subscribe((value: string) => {
      this.firstname(value);
    }, this, 'firstname');
    bindings.shouter.subscribe((value: string) => {
      this.dob(value);
    }, this, 'dob');
    bindings.shouter.subscribe((value: string) => {
      this.gender(value);
    }, this, 'gender');
    bindings.shouter.subscribe((value: string) => {
      this.strDescription(value);
    }, this, 'strDescription');

    this.getItems().then(items => {
      this.lstEmployees(items);
    });
  }

  private getItems(): Promise<Employee[]> {
    /// <summary>Gives the data items from the this.strDescription(). </summary>
    try {
      return new Promise((resolve, reject) => {
        if (sp !== null && sp !== undefined) {
          const items = sp.web.lists.getByTitle(this.strDescription()).items.getAll();
          resolve(items);
        } else {
          reject('Failed getting list data...');
        }
      });
    } catch (Exception) {
      console.log(this.strFileName + " getItems() : " + Exception.message);
    }
  }

  private ensureList(): Promise<List> {
    /// <summary>used for creating batch for database operation. </summary>
    try {
      return new Promise<List>((resolve, reject) => {
        sp.web.lists.ensure(this.strDescription()).then((ler: ListEnsureResult) => {
          if (ler.created) {
            ler.list.fields.addText("FirstName").then(_ => {
              let batch = sp.web.createBatch();
              ler.list.getListItemEntityTypeFullName().then(typeName => {
                batch.execute().then(_ => {
                  resolve(ler.list);
                }).catch(e => reject(e));
              }).catch(e => reject(e));
            }).catch(e => reject(e));
          } else {
            resolve(ler.list);
          }
        }).catch(e => reject(e));
      });
    } catch (Exception) {
      console.log(this.strFileName + " ensureList() : " + Exception.message);
    }
  }

  public addItem(): void {
    /// <summary>used for add and update the Listitems using Item ID ((this.intEditNumber() == 0) => ADD Operation, (this.intEditNumber() > 0) => UPDATE operation). </summary>
    try {
      var submitButtons = document.getElementById("btnAddId");
      submitButtons.innerText = "Add";
      // intEditNumber = 0 in Add Mode
      if (this.intEditNumber() == 0) {
        if (this.firstname() !== "" && this.gender() !== "" && this.dob() != null) {

          this.ensureList().then(list => {

            // add the new item to the SharePoint list
            list.items.add({
              FirstName: this.firstname(),
              Gender: this.gender(),
              DOB: this.dob(),
            }).then((iar: ItemAddResult) => {

              // add the new item to the display
              this.lstEmployees.push({
                intID: iar.data.Id,
                FirstName: iar.data.FirstName,
                Gender: iar.data.Gender,
                DOB: iar.data.DOB,
              });

              // clear the form 
              this.firstname("");
              this.gender("Select");
              this.dob(null);
            });
          });
        }
      }
      // intEditNumber > 0 in Edit Mode whihc stores the Item Id which has to be updated
      else if (this.intEditNumber() > 0) {

        let updatedEmployee: Employee = {
          intID: this.intEditNumber(),
          FirstName: this.firstname(),
          Gender: this.gender(),
          DOB: this.dob()
        };

        this.ensureList().then(list => {
          list.items.getById(this.intEditNumber())
            .update({
              FirstName: this.firstname(),
              Gender: this.gender(),
              DOB: this.dob()
            }).then(_ => {
              this.lstEmployees.replace(this.tempEmployee, updatedEmployee);
            });

          // clear the form
          this.firstname("");
          this.gender("Select");
          this.dob(null);
          this.intEditNumber(0);
        });
      }
    } catch (Exception) {
      console.log(this.strFileName + " editItem() : " + Exception.message);
    }
  }

  public deleteItem(data): void {
    /// <summary>This function deletes the Item which is supplied in the parameter</summary>
    /// <param name="data">Employee Item which is to be deleted</param>
    if (confirm("Are you sure you want to delete this item?")) {
      this.ensureList().then(list => {
        list.items.getById(data.Id).delete().then(_ => {
          this.lstEmployees.remove(data);
        });
      }).catch((e: Error) => {
        console.log(this.strFileName + " deleteItem() : " + e.message);
      });
    }
    this.intEditNumber(0);
  }

  public async editItem(data): Promise<void> {
    /// <summary>This function is used to prefetch the values in the specified fields</summary>
    /// <param name="data">Employee Item which is to be edited</param>
    var submitButtons = document.getElementById("btnAddId");
    submitButtons.innerText = "Update";
    this.tempEmployee = data;
    try {
      this.intEditNumber(data.Id);
      this.firstname(data.FirstName);
      this.gender(data.Gender);
      this.dob(new Date(data.DOB).toISOString().substring(0, 10));
    } catch (Exception) {
      console.log(this.strFileName + " editItem() : " + Exception.message);
    }
  }
}
