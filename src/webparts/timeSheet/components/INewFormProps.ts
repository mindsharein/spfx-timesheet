import { ISiteUserInfo } from "@pnp/sp/site-users";

export interface INewFormProps {
    wpContext: any;
    currentUser: ISiteUserInfo;
    isOpen: boolean;
    onClosed(msg: boolean): void;
    onItemAdded?(refreshData: boolean): void;
}
  