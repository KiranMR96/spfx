import { IUserProfile } from "../birthdayMessage/components/IUserProfile";

export interface IDataService {
  getSPUserProfileProperties: () => Promise<IUserProfile>;
}
