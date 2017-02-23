import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { App } from 'botframework-webchat';
import { DirectLine } from 'botframework-directlinejs';
require('../../../node_modules/BotFramework-WebChat/botchat.css');
import styles from './EchoBot.module.scss';
import * as strings from 'echoBotStrings';
import { IEchoBotWebPartProps } from './IEchoBotWebPartProps';

export default class EchoBotWebPart extends BaseClientSideWebPart<IEchoBotWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `<div id="${this.context.instanceId}" class="${styles.echobot}"></div>`;

    // Get userprofile from SharePoint REST endpoint
    var req = new XMLHttpRequest();
    req.open("GET", "/_api/SP.UserProfiles.PeopleManager/GetMyProperties", false);
    req.setRequestHeader("Accept", "application/json");
    req.send();
    var user = { id: "userid", name: "unknown" };
    if (req.status == 200) {
      var result = JSON.parse(req.responseText);
      user.id = result.Email;
      user.name = result.DisplayName;
    }

    // Initialize DirectLine connection
    var botConnection = new DirectLine({
      secret: "AAos-s9yFEI.cwA.atA.qMoxsYRlWzZPgKBuo5ZfsRpASbo6XsER9i6gBOORIZ8"
    });

    // Initialize the BotChat.App with basic config data and the wrapper element
    App({
        user: user,
        botConnection: botConnection
      }, document.getElementById(this.context.instanceId));

    // Call the bot backchannel to give it user information
    botConnection
      .postActivity({ type: "event", name: "sendUserInfo", value: user.name, from: user })
      .subscribe(id => console.log("success"));                
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
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
