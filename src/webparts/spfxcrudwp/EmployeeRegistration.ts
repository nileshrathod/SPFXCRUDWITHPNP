import pnp from "sp-pnp-js";
import { SPComponentLoader } from '@microsoft/sp-loader';
/* Jquery and BootStrap Reference Start */
import * as $ from 'jquery';
import * as toastr from 'toastr';
import 'jqueryui';

SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');

/* Jquery and BootStrap Reference End*/

$(document).ready(function() {        
    $("#datepicker").datepicker();   
  })

export default class EmployeeRegistration {

  public static GetListItems(listName,listColumns,queryCondition,LookupColumns):Promise<any[]> {

    /*Condotion is used for where or Filter like condition */
    /*LookupColumns is used to expand lookup Columns*/
    if ((queryCondition.length > 0 && LookupColumns.length > 0 ) )
    {
        return pnp.sp.web.lists.getByTitle(listName).items.select(listColumns).expand(LookupColumns).filter(queryCondition).get().then((response) => {
            console.log(response);
            return response;});
    }
    else if(queryCondition.length >0)
    {
        return pnp.sp.web.lists.getByTitle(listName).items.select(listColumns).filter(queryCondition).get().then((response) => {
            console.log(response);
            return response;});
    }
    else 
    {
        return pnp.sp.web.lists.getByTitle(listName).items.select(listColumns).get().then((response) => {
            console.log(response);
            return response;});
    }
    
  }


    public static templateHTML: string = `
    
    </div><div id="lists"></div>
      <div class="container">
      <h2>Add</h2>
      <p>Add SharePoint List Item using SPFX and pnp.js</p>
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
                      <option>-Select-</option>                  
                  </select>
              </div>
              <div class="form-group">
                  <label for="sel1">City</label>
                  <select class="form-control" id="ddlCity">                  
                  <option>-Select-</option>                  
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
              <input type="button" id="AddItem" class="btn btn" value="Submit"/>
                  <button type="btnCancel" class="btn btn">Cancel</button>
              </div>
          </form>
      </div>
      `



}


