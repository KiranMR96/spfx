import * as React from "react";
import styles from "./BirthdayMessage.module.scss";
import { IUserProfileProps } from "./IUserProfileProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { IUserProfileState } from "./IUserProfileState";
import { UserProfileService } from "../../Services/UserProfileService";
import { IUserProfile } from "./IUserProfile";
import { IDataService } from "../../Services/IDataService";

export class UserProfile implements IUserProfile {
  public FirstName: string;
  public LastName: string;
  public BirthDate: string;
  public HiredDate: string;
  public UserProfileProperties: any[];
}

export default class BirthdayMessage extends React.Component<
  IUserProfileProps,
  IUserProfileState,
  {}
> {
  private dataCenterServiceInstance: IDataService;

  constructor(props: IUserProfileProps, state: IUserProfileState) {
    super(props);

    let userProfile: IUserProfile = new UserProfile();
    (userProfile.BirthDate = ""),
      (userProfile.FirstName = ""),
      (userProfile.LastName = ""),
      (userProfile.UserProfileProperties = []);

    this.state = { userProfileItems: userProfile };
  }

  public componentWillUnmount(): void {
    let serviceScope = this.props.serviceScope;
    this.dataCenterServiceInstance = serviceScope.consume(
      UserProfileService.serviceKey
    );

    this.dataCenterServiceInstance
      .getSPUserProfileProperties()
      .then((userProfileItems: IUserProfile) => {
        console.log(userProfileItems);
        for (
          let i: number = 0;
          i < userProfileItems.UserProfileProperties.length;
          i++
        ) {
          if (userProfileItems.UserProfileProperties[i].Key == "FirstName") {
            userProfileItems.FirstName =
              userProfileItems.UserProfileProperties[i].Value;
          }

          if (userProfileItems.UserProfileProperties[i].Key == "LastName") {
            userProfileItems.LastName =
              userProfileItems.UserProfileProperties[i].Value;
          }

          // if (userProfileItems.UserProfileProperties[i].Key == "WorkPhone") {
          //   userProfileItems.BirthDate =
          //     userProfileItems.UserProfileProperties[i].Value;
          // }

          // if (userProfileItems.UserProfileProperties[i].Key == "Department") {
          //   userProfileItems.Department =
          //     userProfileItems.UserProfileProperties[i].Value;
          // }

          // if (userProfileItems.UserProfileProperties[i].Key == "PictureURL") {
          //   userProfileItems.PictureURL =
          //     userProfileItems.UserProfileProperties[i].Value;
          // }
        }
        this.setState({ userProfileItems: userProfileItems });
        console.log(userProfileItems);
      });
  }

  public render(): React.ReactElement<IUserProfileProps> {
    return (
      <div className={styles.birthdayMessage}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p>
                Name : {this.state.userProfileItems.LastName},
                {this.state.userProfileItems.FirstName}
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
