import { spfi,SPFI, SPFx } from '@pnp/sp';
import "@pnp/sp";
import "@pnp/sp/site-users";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";

import "@pnp/logging";

import PnPTelemetry from "@pnp/telemetry-js";

import { LogLevel, PnPLogging } from '@pnp/logging';

import { WebPartContext } from '@microsoft/sp-webpart-base';
import ITimeSheet from '../models/ITimeSheet';


 // Initialize & return SP Object
 let _sp : SPFI = null;

export default function getSP(wpContext: WebPartContext) : SPFI {
    // Turnoff telemetry
    PnPTelemetry.getInstance().optOut();

   if(_sp==null) {
     _sp = spfi().using(SPFx(wpContext)).using(PnPLogging(LogLevel.Warning));
   }

   return _sp;
}

export async function getCurrentUser(wpContext) {
  const sp = getSP(wpContext);

  const user = await sp.web.currentUser();

  return user;
}

export async function getTimeSheetItems(wpContext: WebPartContext) : Promise<ITimeSheet[]> {
  const sp = getSP(wpContext);

  const user = await sp.web.currentUser();

  let data = await sp.web.lists.getByTitle("TimeSheet").items
                    .expand("Person")
                    .select("ID, Title, From, To, Hours, Person/Id, Person/Name, Notes")
                    .filter(`Person/Name eq '${user.LoginName}'`)
                    .getAll();

  return data;
 }

 export async function deleteTimeSheetItem(id: number, wpContext:WebPartContext) : Promise<string> {
   try {
    const sp = getSP(wpContext);
    let data = await sp.web.lists.getByTitle("TimeSheet").items.getById(id).delete();

    return "";
   } catch(ex) {
     return "Error deleting item: " + ex.toString();
   }
 }

