export default interface ITimeSheet {
    Title : string;
    ProjectTaskId : number;
    From: Date;
    To: Date;
    Hours: number;
    PersonId: number;
    Notes: string;
}