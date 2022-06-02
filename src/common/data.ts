import { spfi,SPFI, SPFx } from '@pnp/sp';
import "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";

import "@pnp/logging";

import PnPTelemetry from "@pnp/telemetry-js";

import { LogLevel, PnPLogging } from '@pnp/logging';

import { WebPartContext } from '@microsoft/sp-webpart-base';


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

