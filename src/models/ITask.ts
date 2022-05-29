import IProject from "./IProject";

export default interface ITask {
    Title: string;
    Description: string;
    Project: IProject;
    AssignedTo: number;
    Priority: string;
    CompleteBy: Date;
    Hours: number;
    Comments: string;
}