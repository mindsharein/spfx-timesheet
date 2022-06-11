export default interface IConfirmDialogProps {
    show: boolean;
    title: string;
    message: string;
    onClick?(result: boolean) : void;
}