import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import * as strings from 'ViewEmployeeWebPartStrings';


import * as $ from 'jquery';
import Utility from '../Utility';
import * as moment from 'moment';
import 'datatables.net';
import * as toastr from 'toastr';


SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.15/css/jquery.dataTables.min.css');
SPComponentLoader.loadScript("https://code.jquery.com/jquery-1.12.4.min.js");
SPComponentLoader.loadScript("https://cdn.datatables.net/1.10.15/js/jquery.dataTables.min.js");
SPComponentLoader.loadScript("https://momentjs.com/downloads/moment.min.js");
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.css');


$(document).ready(function () {

  $('#example').on('click', 'a.editor_edit', function (e) {
    var id = $(this).parent().siblings(":first").text();
    //alert("List Item Id to Edit Record " + id);
    e.preventDefault();
    var editpageurl = "https://sharepointheliossolutions.sharepoint.com/sites/Intranet/Pages/UpdateEmployee.aspx?id=" + id;
    window.location.assign(editpageurl);
  });

  $('#example').on('click', 'a.editor_remove', function (e) {
    var id = $(this).parent().siblings(":first").text();
    if (confirm("Are you sure you want to delete this item?")) {
      //alert("List Item Id to Delete Record " + id);
      //Remove Row From DataTable  
      $(this).closest("tr").remove();
      new ViewEmployeeWebPart().DeleteListItemByID(id);
    }
    e.preventDefault();
  });

});


export interface IViewEmployeeWebPartProps {
  description: string;
}

export interface IListItem {
  Title: string;
  State: string;
  City: string;
  Email: string;
  Contac: string;
  JoiningDate: Date;
}
export default class ViewEmployeeWebPart extends BaseClientSideWebPart<IViewEmployeeWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = ` <div class="container">
    <h2>View</h2>
    <p>View SharePoint List using SPFX and pnp.js</p>

    <div>
    <hr/>
     <p id='divEmployee'></p> <table id="example" class="display" cellspacing="0" width="100%">
     <thead>
         <tr>
         <th class="hide_me" >Id</th>
             <th>Title</th>
             <th>State</th>
             <th>City</th>
             <th>Email</th>
             <th>Contact</th>
             <th>JoiningDate</th>
             <th>Edit</th>
         </tr>
     </thead>
 </table>`; this.GetEmployeeDetails();
  }

  //Get EmployeeDetails List using PNP.Js 
  GetEmployeeDetails() {

    let listName: string = "EmployeeDetails";
    let queryCondition: string = "";
    let listColumns: string = "ID,Title,State,City,Email,Contact,JoiningDate";
    let listLookUpColumns: string = "";
    var varTable: string = '';

    Utility.GetListItems(listName, listColumns, queryCondition, listLookUpColumns)
      .then((response) => {
        this._renderList(response);
      });
  }

  //Create jQuery Datatable and Format it using Jquery Datatable CSS
  private _renderList(items: IListItem[]): void {
    $('#example').DataTable({
      data: items,
      columns: [
        { "data": "ID", className: "hide_me" },
        { "data": "Title" },
        { "data": "State" },
        { "data": "City" },
        { "data": "Email" },
        { "data": "Contact" },
        { "data": "JoiningDate", render: function (data, type, row) { if (type === "sort" || type === "type") { return data; } return moment(data).format("MM/DD/YYYY"); } },
        { "data": null, className: "center", defaultContent: '<a href=""  class="editor_edit">Edit</a> / <a href=""  class="editor_remove">Delete</a>' }
      ]
    });

    //Make First Column Visible False 
    //$('#example').DataTable().column( 0 ).visible( false );
    //Hide Column by class
    $('.hide_me').css('display', 'none');
  }

  DeleteListItemByID(itemid): void {

    Utility.DeleteListItemsByID("EmployeeDetails", itemid)
      .then(function sucess() { toastr.warning("Record Deleted Successfully", "Success"); })
      .catch(function error() { toastr.error("An Error Occues while deleting record", "Error"); })

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
