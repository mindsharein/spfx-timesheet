export default interface ITimeSheet {
    Title : string;
    ProjectTaskId : number;
    From: Date;
    To: Date;
    Hours: number;
    Person?: any;
    ProjectTask?: any;
    PersonId: number;
    Notes: string;
}