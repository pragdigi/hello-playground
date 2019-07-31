import {
  Version,
  DisplayMode,
  Environment,
  EnvironmentType,
  Log
} from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./HelloPlaygroundWebPart.module.scss";
import * as strings from "HelloPlaygroundWebPartStrings";

export interface IHelloPlaygroundWebPartProps {
  description: string;
}

export default class HelloPlaygroundWebPart extends BaseClientSideWebPart<
  IHelloPlaygroundWebPartProps
> {
  public render(): void {
    const pageMode: string =
      this.displayMode === DisplayMode.Edit
        ? "You are currently in edit mode"
        : "You are currently in read mode";

    const environmentType: string =
      Environment.type === EnvironmentType.Local
        ? "This is your local environment"
        : "This is your sharepoint environment";

    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      "message"
    );
    setTimeout(() => {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);

      this.domElement.innerHTML = `
      <div class="${styles.helloPlayground}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to SharePoint!</span>
              <p class="${styles.subTitle}">Playing around with this WebPart</p>
              <p class="${
                styles.subTitle
              }"><strong>Page mode:</strong> ${pageMode}</p>
              <p class="${
                styles.subTitle
              }"><strong>Environment:</strong> ${environmentType}</p>
              <p class="${styles.description}">${escape(
        this.properties.description
      )}</p>
              <a href="#" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
      this.domElement
        .getElementsByClassName(`${styles.button}`)[0]
        .addEventListener("click", (event: any) => {
          event.preventDefault();
          alert("Welcome to my Playground!");
        });
    }, 5000);
    
    Log.info("HelloPlayground", "message", this.context.serviceScope);
    Log.warn("HelloPlayground", "WARNING message", this.context.serviceScope);
    Log.error(
      "HelloPlayground",
      new Error("Error message"),
      this.context.serviceScope
    );
    Log.verbose(
      "HelloPlayground",
      "VERBOSE message",
      this.context.serviceScope
    );
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
