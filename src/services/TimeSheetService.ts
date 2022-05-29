import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import ITimeSheet from  "../models/ITimeSheet";

export default class TimeSheetService {
    private sp : any;
    private listName : string = "TimeSheet"

    constructor(private context: any) {
        
    }

    public async init() : Promise<boolean> {
        this.sp = spfi().using(SPFx(this.context));

        return Promise.resolve(true);
    }

    public async getItems(lim: number) : Promise<ITimeSheet[]> {
        return new Promise(async (res,rej) => {
            try {
                let items : ITimeSheet[] = [];
                
                items = await this.sp.web.lists.getByTitle(this.listName).items();

                console.log(`TIMESHEETSERVICE: items fetched : ${ items }`);
                console.log(items);
        
                res(items);
            } catch(ex) {
                rej(ex);
            }
        });
    } 

}