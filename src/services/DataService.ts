import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";


let _sp = null;
let _wpctx = null;

// Set the WP Context
export const initDS = (context: WebPartContext) => {
    _wpctx = context;

    getSP(context);
}

// Initialize a PnP SP Object for SPFX
const getSP = (context: WebPartContext) => {
    if(_sp==null) {
        _sp = spfi().using(SPFx(_wpctx));
        return _sp;
    }

    return _sp;
} 

//------------------------ LIST FUNCTIONS --------------------------------------//
export const getItems = async (listName: string) => {
       
    let items = await _sp.web.lists.getByTitle(listName).items();

    return items;
}

export const getItemById = async (listName: string, id: number) => {

    let item = await _sp.web.lists.getByTitle(listName).getItemById(id);

    return item;
}

export const addItem = async (listName: string, newItem: any) => {
    const result = await _sp.web.lists.getByTitle(listName).items.add(newItem);

    console.log("IAR : " + result);

    return result.data;
}