
import pnp from "sp-pnp-js";

export  default class  Utility{
    
    public static GetListItems(listName,listColumns,queryCondition,LookupColumns):Promise<any[]> {

        /*Condotion is used for where or Filter like condition */
        /*LookupColumns is used to expand lookup Columns*/
        if ((queryCondition.length > 0 && LookupColumns.length > 0 ) )
        {
            return pnp.sp.web.lists.getByTitle(listName).items.select(listColumns).expand(LookupColumns).filter(queryCondition).get().then((response) => {
                //console.log(response);
                return response;});
        }
        else if(queryCondition.length >0)
        {
            return pnp.sp.web.lists.getByTitle(listName).items.select(listColumns).filter(queryCondition).get().then((response) => {
                //console.log(response);
                return response;});
        }
        else 
        {
            return pnp.sp.web.lists.getByTitle(listName).items.select(listColumns).get().then((response) => {
                //console.log(response);
                return response;});
        }
        
      }

      public static DeleteListItemsByID(listName ,itemID):Promise<void> {
        return  pnp.sp.web.lists.getByTitle(listName).items.getById(itemID).delete().then(response =>{
            //console.log(response);
            return response ;
        });        
      }

      public static GetListItemsByID(listName,itemID,listColumns,LookupColumns):Promise<any[]> {

        return pnp.sp.web.lists.getByTitle(listName).items.getById(itemID).get().then((response) => {
            //console.log(response);
            return response;});
        
      }


}