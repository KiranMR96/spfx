import { ServiceScope } from "@microsoft/sp-core-library";

export interface IUserProfileProps {
  description: string;
  userName: string;
  serviceScope: ServiceScope;
}
