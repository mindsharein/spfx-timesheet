import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/presets/all";


import IProject from "../models/IProject";

export class ProjectService {
    private sp : any;
    private listName : string = "Projects"

    constructor(context: any) {
        this.sp = spfi().using(SPFx(context));
    }

    public async getItems(lim: number) {
        const items : IProject[] = await this.sp.web.lists.getByTitle(this.listName).getItems().limit(lim);

        return items;
    } 

}