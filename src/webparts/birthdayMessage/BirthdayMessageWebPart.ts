import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "BirthdayMessageWebPartStrings";
import BirthdayMessage from "./components/BirthdayMessage";
import { IUserProfileProps } from "./components/IUserProfileProps";

export interface IBirthdayMessageWebPartProps {
  description: string;
}

export default class BirthdayMessageWebPart extends BaseClientSideWebPart<IBirthdayMessageWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IUserProfileProps> = React.createElement(
      BirthdayMessage,
      {
        description: this.properties.description,
        userName: encodeURIComponent(
          "i:0#.f|membership|" + this.context.pageContext.user.loginName
        ),
        serviceScope: this.context.serviceScope,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
