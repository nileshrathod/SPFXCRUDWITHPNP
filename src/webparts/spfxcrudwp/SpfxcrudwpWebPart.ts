import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './SpfxcrudwpWebPart.module.scss';
import * as strings from 'SpfxcrudwpWebPartStrings';
import objEmployee from './EmployeeRegistration'
import pnp, { log } from "sp-pnp-js";
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as toastr from 'toastr';
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.css');

export interface ISpfxcrudwpWebPartProps {
  description: string;
}

export default class SpfxcrudwpWebPart extends BaseClientSideWebPart<ISpfxcrudwpWebPartProps> {

  private AddEventListeners(): void {
    document.getElementById('AddItem').addEventListener('click', () => this.AddItem());
    document.getElementById('ddlState').addEventListener('change', () => this.BindCity());
  }

  public render(): void {

    this.domElement.innerHTML = objEmployee.templateHTML;

    this.BindState();

    this.AddEventListeners();
  }

  BindState(){
    let listName:string = "IndiaState";
    let queryCondition:string ="";
    let listColumns:string ="ID,Title";
    let listLookUpColumns:string ="";
    objEmployee.GetListItems(listName,listColumns, queryCondition,listLookUpColumns)
    .then((response) => {                       
            response.forEach(item => {
              $('#ddlState').append($('<option></option>').val(item.Title).html(item.Title));
            }); 
    });
  }

  BindCity() {
    $("#ddlCity").empty();      
    $("#ddlCity").append('<option value="">-Select-</option>');
    let listName:string = "IndiaCity";
    let queryCondition:string ="StateName/Title eq '" + $("#ddlState").val() + "'" ;
    let listColumns:string ="ID,Title,StateName/Title,StateName/ID";
    let listLookUpColumns:string="StateName";
    objEmployee.GetListItems(listName,listColumns,queryCondition,listLookUpColumns)
    .then((response) => {                       
            response.forEach(item => {
              $('#ddlCity').append($('<option></option>').val(item.Title).html(item.Title));
            }); 
    });
  }

  /*Add List Item Event Handler Start */
  AddItem() {
    pnp.sp.web.lists.getByTitle('EmployeeDetails').items.add({
      Title: $("#txtName").val(),
      State: $("#ddlState").val(),
      City: $("#ddlCity").val(),
      Email: $("#txtEmail").val(),
      Contact: $("#txtContact").val(),
      JoiningDate: $("#datepicker").val()
    })
     .then(function() {  toastr.success("Record  Added !", "Success");})
     .catch(function() { toastr.error("An error occured while adding record.", "Error"); })
    

  
  }
  /*Add List Item Event Handler End */

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                },
                )
              ]
            }
          ]
        }
      ]
    };
  }
}

