var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
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
import * as ko from 'knockout';
import styles from './KnockoutCrudForm.module.scss';
import { sp } from "sp-pnp-js";
var KnockoutCrudFormViewModel = /** @class */ (function () {
    function KnockoutCrudFormViewModel(bindings) {
        /// <summary>constructor</summary>
        /// <param name="IKnockoutCrudFormBindingContext">bindings values</param>
        var _this = this;
        /// <summary>KnockoutCrudFormViewModel class</summary>
        this.strDescription = ko.observable('');
        this.firstname = ko.observable('');
        this.lastname = ko.observable('');
        this.gender = ko.observable('');
        this.dob = ko.observable('');
        this.availableGenders = ko.observableArray(['Select', 'Male', 'Female', 'Others']);
        this.lstEmployees = ko.observableArray([]);
        this.intEditNumber = ko.observable();
        this.knockoutCrudFormClass = styles.knockoutCrudForm;
        this.containerClass = styles.container;
        this.rowClass = styles.row;
        this.columnClass = styles.column;
        this.titleClass = styles.title;
        this.subTitleClass = styles.subTitle;
        this.descriptionClass = styles.strDescription;
        this.firstnameClass = styles.firstname;
        this.genderClass = styles.gender;
        this.dobClass = styles.dob;
        this.buttonClass = styles.button;
        this.labelClass = styles.label;
        this.strFileName = "KnockoutCrudFormViewModel";
        this.firstname(bindings.firstname);
        this.gender(bindings.gender);
        this.dob(bindings.dob);
        this.intEditNumber(bindings.intEditNumber);
        this.strDescription(bindings.strDescription);
        this.context = bindings.context;
        // When web part fields is updated, change this view model's values
        bindings.shouter.subscribe(function (value) {
            _this.intEditNumber(value);
        }, this, 'intEditNumber');
        this.intEditNumber(0);
        bindings.shouter.subscribe(function (value) {
            _this.firstname(value);
        }, this, 'firstname');
        bindings.shouter.subscribe(function (value) {
            _this.dob(value);
        }, this, 'dob');
        bindings.shouter.subscribe(function (value) {
            _this.gender(value);
        }, this, 'gender');
        bindings.shouter.subscribe(function (value) {
            _this.strDescription(value);
        }, this, 'strDescription');
        this.getItems().then(function (items) {
            _this.lstEmployees(items);
        });
    }
    KnockoutCrudFormViewModel.prototype.getItems = function () {
        var _this = this;
        /// <summary>Gives the data items from the this.strDescription(). </summary>
        try {
            return new Promise(function (resolve, reject) {
                if (sp !== null && sp !== undefined) {
                    var items = sp.web.lists.getByTitle(_this.strDescription()).items.getAll();
                    resolve(items);
                }
                else {
                    reject('Failed getting list data...');
                }
            });
        }
        catch (Exception) {
            console.log(this.strFileName + " getItems() : " + Exception.message);
        }
    };
    KnockoutCrudFormViewModel.prototype.ensureList = function () {
        var _this = this;
        /// <summary>used for creating batch for database operation. </summary>
        try {
            return new Promise(function (resolve, reject) {
                sp.web.lists.ensure(_this.strDescription()).then(function (ler) {
                    if (ler.created) {
                        ler.list.fields.addText("FirstName").then(function (_) {
                            var batch = sp.web.createBatch();
                            ler.list.getListItemEntityTypeFullName().then(function (typeName) {
                                batch.execute().then(function (_) {
                                    resolve(ler.list);
                                }).catch(function (e) { return reject(e); });
                            }).catch(function (e) { return reject(e); });
                        }).catch(function (e) { return reject(e); });
                    }
                    else {
                        resolve(ler.list);
                    }
                }).catch(function (e) { return reject(e); });
            });
        }
        catch (Exception) {
            console.log(this.strFileName + " ensureList() : " + Exception.message);
        }
    };
    KnockoutCrudFormViewModel.prototype.addItem = function () {
        var _this = this;
        /// <summary>used for add and update the Listitems using Item ID ((this.intEditNumber() == 0) => ADD Operation, (this.intEditNumber() > 0) => UPDATE operation). </summary>
        try {
            var submitButtons = document.getElementById("btnAddId");
            submitButtons.innerText = "Add";
            // intEditNumber = 0 in Add Mode
            if (this.intEditNumber() == 0) {
                if (this.firstname() !== "" && this.gender() !== "" && this.dob() != null) {
                    this.ensureList().then(function (list) {
                        // add the new item to the SharePoint list
                        list.items.add({
                            FirstName: _this.firstname(),
                            Gender: _this.gender(),
                            DOB: _this.dob(),
                        }).then(function (iar) {
                            // add the new item to the display
                            _this.lstEmployees.push({
                                intID: iar.data.Id,
                                FirstName: iar.data.FirstName,
                                Gender: iar.data.Gender,
                                DOB: iar.data.DOB,
                            });
                            // clear the form 
                            _this.firstname("");
                            _this.gender("Select");
                            _this.dob(null);
                        });
                    });
                }
            }
            // intEditNumber > 0 in Edit Mode whihc stores the Item Id which has to be updated
            else if (this.intEditNumber() > 0) {
                var updatedEmployee_1 = {
                    intID: this.intEditNumber(),
                    FirstName: this.firstname(),
                    Gender: this.gender(),
                    DOB: this.dob()
                };
                this.ensureList().then(function (list) {
                    list.items.getById(_this.intEditNumber())
                        .update({
                        FirstName: _this.firstname(),
                        Gender: _this.gender(),
                        DOB: _this.dob()
                    }).then(function (_) {
                        _this.lstEmployees.replace(_this.tempEmployee, updatedEmployee_1);
                    });
                    // clear the form
                    _this.firstname("");
                    _this.gender("Select");
                    _this.dob(null);
                    _this.intEditNumber(0);
                });
            }
        }
        catch (Exception) {
            console.log(this.strFileName + " editItem() : " + Exception.message);
        }
    };
    KnockoutCrudFormViewModel.prototype.deleteItem = function (data) {
        var _this = this;
        /// <summary>This function deletes the Item which is supplied in the parameter</summary>
        /// <param name="data">Employee Item which is to be deleted</param>
        if (confirm("Are you sure you want to delete this item?")) {
            this.ensureList().then(function (list) {
                list.items.getById(data.Id).delete().then(function (_) {
                    _this.lstEmployees.remove(data);
                });
            }).catch(function (e) {
                console.log(_this.strFileName + " deleteItem() : " + e.message);
            });
        }
        this.intEditNumber(0);
    };
    KnockoutCrudFormViewModel.prototype.editItem = function (data) {
        return __awaiter(this, void 0, void 0, function () {
            var submitButtons;
            return __generator(this, function (_a) {
                submitButtons = document.getElementById("btnAddId");
                submitButtons.innerText = "Update";
                this.tempEmployee = data;
                try {
                    this.intEditNumber(data.Id);
                    this.firstname(data.FirstName);
                    this.gender(data.Gender);
                    this.dob(new Date(data.DOB).toISOString().substring(0, 10));
                }
                catch (Exception) {
                    console.log(this.strFileName + " editItem() : " + Exception.message);
                }
                return [2 /*return*/];
            });
        });
    };
    return KnockoutCrudFormViewModel;
}());
export default KnockoutCrudFormViewModel;
//# sourceMappingURL=KnockoutCrudFormViewModel.js.map