import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EditEmployeeWebPart.module.scss';
import * as strings from 'EditEmployeeWebPartStrings';


/*Custom Import*/
import {  UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import * as $ from 'jquery';
import * as toastr from 'toastr';
import pnp, { Util } from "sp-pnp-js";
import { SPComponentLoader } from '@microsoft/sp-loader';
import Utility from '../Utility';
import 'jqueryui';
import * as moment from 'moment';
SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.css');
/*Custom Import*/

$(document).ready(function() {        
  $("#datepicker").datepicker();   
})


export interface IEditEmployeeWebPartProps {
  description: string;
}

export default class EditEmployeeWebPart extends BaseClientSideWebPart<IEditEmployeeWebPartProps> {

  private AddEventListeners(): void {    
    document.getElementById('btnUpdate').addEventListener('click', () => this.UpdateItem());    
    document.getElementById('ddlState').addEventListener('change', () => this.BindCityOnTrigger());
  }

  public render(): void {
    this.domElement.innerHTML = `</div><div id="lists"></div>
    <div class="container">
    <h2>Update</h2>
    <p>Update SharePoint List using SPFX and pnp.js</p>
        <form>
            <div class="form-group">
                <label for="inputdefault">Name</label>
                <input class="form-control" id="txtName" type="text">
            </div>
            <div class="form-group">
                <label for="inputdefault">Joining Date</label>
                <input type="text" name="selected_date" class="form-control" id="datepicker" readonly/>
            </div>
            <div class="form-group">
                <label for="sel1">State</label>
                <select class="form-control" id="ddlState" OnChange={this.handleChange} >
                <option value='0'>-Select-</option>             
                </select>
            </div>
            <div class="form-group">
                <label for="sel1">City</label>
                <select class="form-control" id="ddlCity">                  
                <option value='0'>-Select-</option>                  
                </select>
            </div>
            <div class="form-group">
                <label for="inputdefault">Email</label>
                <input class="form-control" id="txtEmail" type="text">
            </div>
            <div class="form-group">
                <label for="inputdefault">Contact</label>
                <input class="form-control" id="txtContact" type="text">
            </div>
         
            <div class="form-group">              
            <input type="button" id="btnUpdate" class="btn btn" value="Update"/>
                <button type="btnCancel" class="btn btn">Cancel</button>
            </div>
        </form>
    </div>`; this.AddEventListeners();this.GetEmployeeDetails();    
  }

  GetEmployeeDetails(){
    this.BindState();
  }

    BindPageControl()
    {
      let listName: string = "EmployeeDetails";        
      let listColumns: string = "ID,Title,State,City,Email,Contact,JoiningDate";
      let listLookUpColumns: string = "";
      var varTable: string = '';
  
      const url : any = new URL(window.location.href);
      const itemId = url.searchParams.get("id");

      Utility.GetListItemsByID(listName,itemId, listColumns, listLookUpColumns)
      .then(function success(data){new EditEmployeeWebPart().AssignValesToControl(data)})
      .catch(function failed(){
        $("#txtName").val("");
        $("#datepicker").val("");
        $("ddlState").val("0")
        $("#ddlCity").val("0");
        $("#txtEmail").val("");
        $("#txtContact").val("");
      });   
    }

  AssignValesToControl(item)
  {
    $("#ddlState").val(item["State"]); 
    $("#txtName").val(item["Title"]);  
    $("#datepicker").val(moment(item["JoiningDate"]).format("MM/DD/YYYY"));     
    this.BindCity(item["State"],item["City"]);     
    $("#ddlCity").val(item["City"]);  
    $("#txtEmail").val(item["Email"]);  
    $("#txtContact").val(item["Contact"]);  
  }

  BindState(){
    $("#ddlState").empty();      
    $("#ddlState").append('<option value="0">-Select-</option>');
    let listName:string = "IndiaState";
    let queryCondition:string ="";
    let listColumns:string ="ID,Title";
    let listLookUpColumns:string ="";
    Utility.GetListItems(listName,listColumns, queryCondition,listLookUpColumns)
    .then((response) => {                       
            response.forEach(item => {
              $('#ddlState').append($('<option></option>').val(item.Title).html(item.Title));
            });          
            this.BindPageControl();     
    });
  }

  BindCityOnTrigger() {
    $("#ddlCity").empty();      
    $("#ddlCity").append('<option value="0">-Select-</option>');
    let listName:string = "IndiaCity";
    let queryCondition:string ="StateName/Title eq '" + $("#ddlState").val() + "'" ;
    let listColumns:string ="ID,Title,StateName/Title,StateName/ID";
    let listLookUpColumns:string="StateName";
    Utility.GetListItems(listName,listColumns,queryCondition,listLookUpColumns)
    .then((response) => {                       
            response.forEach(item => {
              $('#ddlCity').append($('<option></option>').val(item.Title).html(item.Title));
            });   
    });
  }


  BindCity(State,City) {
    $("#ddlCity").empty();      
    $("#ddlCity").append('<option value="0">-Select-</option>');
    let listName:string = "IndiaCity";
    let queryCondition:string ="StateName/Title eq '" + State + "'" ;
    let listColumns:string ="ID,Title,StateName/Title,StateName/ID";
    let listLookUpColumns:string="StateName";
    Utility.GetListItems(listName,listColumns,queryCondition,listLookUpColumns)
    .then((response) => {                       
            response.forEach(item => {
              $('#ddlCity').append($('<option></option>').val(item.Title).html(item.Title));
            }); 
            $("#ddlCity").val(City);
    });
  }

  UpdateItem()
  {
    const url : any = new URL(window.location.href);
    const itemid = url.searchParams.get("id");
    pnp.sp.web.lists.getByTitle('EmployeeDetails').items.getById(itemid).update({
      Title: $("#txtName").val(),
      State: $("#ddlState").val(),
      City: $("#ddlCity").val(),
      Email: $("#txtEmail").val(),
      Contact: $("#txtContact").val(),
      JoiningDate: $("#datepicker").val()
    })
     .then(function() {  toastr.success("Record  Updated !", "Success"); 
     $("#txtName").val("");
     $("#datepicker").val("");
     $("#ddlState").val("0")
     $("#ddlCity").val("0");
     $("#txtEmail").val("");
     $("#txtContact").val("");
    
    })
     .catch(function() { toastr.error("An error occured while adding record.", "Error"); })
  }

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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
