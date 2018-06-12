import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FileUploadWebPart.module.scss';
import * as strings from 'FileUploadWebPartStrings';
import { ConsoleListener, Web, Logger, LogLevel, ODataRaw, CheckinType } from "sp-pnp-js";


export interface IFileUploadWebPartProps {
  description: string;
}

export default class FileUploadWebPart extends BaseClientSideWebPart<IFileUploadWebPartProps> {

  private AddEventListeners(): void {
    document.getElementById('btnUpdate').addEventListener('click', () => this.UploadFiles());
  }

  public render(): void {
    this.domElement.innerHTML = `
    <input type="file" id="myFile" value="Upload File" />
    <input type="button" id="btnUpdate" class="btn btn" value="Update"/>
      `; this.AddEventListeners();
  }

  UploadFiles() {

    var siteUrl =this.context.pageContext.web.absoluteUrl ;
    let web = new Web(siteUrl);
    let folderpath = "/sites/Intranet/Documents/";

    let input = <HTMLInputElement>document.getElementById("myFile");
    let file = input.files[0];

    var filename = folderpath + file.name;

    // you can adjust this number to control what size files are uploaded in chunks
    if (file.size <= 10485760) {

      //small size file upload

      /* UPLOAD FILE TO GIVEN SHAREPOINT LOCATION */
      web.getFolderByServerRelativeUrl(folderpath).files.add(file.name, file, true).then(function () {

        /* CHECK IN FILE TO GIVEN SHAREPOINT LOCATION */
        web.getFileByServerRelativeUrl(filename).checkin().then(function () {

          /* PUBLISH FILE TO GIVEN SHAREPOINT LOCATION */
          web.getFileByServerRelativeUrl(filename).publish("Publish comment").then(_ => {
            console.log("File published!");
          });

        });

      });


    }

    else {

      //large size file upload
      web.getFolderByServerRelativeUrl(folderpath).files.addChunked(file.name, file, data => {
        Logger.log({ data: data, level: LogLevel.Verbose, message: "progress" });
      }, true).then(function () {

        /* CHECK IN FILE TO GIVEN SHAREPOINT LOCATION */
        web.getFileByServerRelativeUrl(filename).checkin().then(function () {

          /* PUBLISH FILE TO GIVEN SHAREPOINT LOCATION */
          web.getFileByServerRelativeUrl(filename).publish("Publish comment").then(_ => {
            console.log("File published!");
          });

        });

      });;



    }


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
