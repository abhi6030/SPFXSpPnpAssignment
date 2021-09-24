import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './SpPnpCrudWebPart.module.scss';
import * as strings from 'SpPnpCrudWebPartStrings';
import * as pnp from 'sp-pnp-js';

export interface ISpPnpCrudWebPartProps {
  description: string;
}

export default class SpPnpCrudWebPart extends BaseClientSideWebPart<ISpPnpCrudWebPartProps> {

    public render(): void {
    this.domElement.innerHTML = `
    <div>
    <div>
      <table border='5' bgcolor='aqua'>
      <tr>
      <td>Please Enter Software ID </td>
      <td><input type='text' id='txtID' />
      </td>
      </tr>
    </table>
    <hr/>
    <div>
      <input type='submit'  value='Accept' id='btnAccept' />
      <input type='submit'  value='Reject' id='btnReject' />
    </div>
    </div>
      <div id="divStatus"/>
      <h2>Get All List Items</h2>
    <hr/>
      <div id="spListData" />
    </div>`;

    this._bindEvents();
    this.readAllItems();
  }

  public readAllItems(): void {

    let html: string = '<table border=1 width=100% style="bordercollapse: collapse;">';
    html += `<th>ID</th><th>Title</th><th>Status</th><th>Name</th><th>Version</th><th>Description</th>`;

    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.get().then((items: any[]) => {
      items.forEach(function (item) {
        if (item.SoftwareVendor === "In-Progress") {
          html += `
      <tr>
      <td>${item["ID"]}</td>
      <td>${item["Title"]}</td>
      <td>${item["SoftwareVendor"]}</td>
      <td>${item["SoftwareName"]}</td>
      <td>${item["SoftwareVersion"]}</td>
      <td>${item["SoftwareDescription"]}</td>
      </tr>
      `;
        }
      });
      html += `</table>`;
      const allitems: Element = this.domElement.querySelector('#spListData');
      allitems.innerHTML = html;
    });
  }
    private _bindEvents(): void {

    this.domElement.querySelector('#btnAccept').addEventListener('click', () => { this.acceptitem(); });
    this.domElement.querySelector('#btnReject').addEventListener('click', () => { this.rejectitem(); });
  }
  private deleteListItem(): void {
    const id = document.getElementById("txtID")["value"];
    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.getById(id).delete();
    alert("list item Deleted");
  }
  private acceptitem(): void {
    var softwareVendor = "Accept";
    let id: number = document.getElementById("txtID")["value"];
    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.getById(id).update({

      SoftwareVendor: softwareVendor,

    }).then(r => {
      alert(" Record-Accepted ");
    });
  }
    private rejectitem(): void {
    var softwareVendor = "Reject";
    let id: number = document.getElementById("txtID")["value"];
    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.getById(id).update({

      SoftwareVendor: softwareVendor,

    }).then(r => {
      alert(" Record-Rejected ");
    });
  }
  private readListItem(): void {

    const id = document.getElementById("txtID")["value"];

    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.getById(id).get().then((item: any) => {
      document.getElementById("txtSoftwareTitle")["value"] = item["Title"];
      document.getElementById("txtSoftwareName")["value"] = item["SoftwareName"];
      document.getElementById("txtSoftwareVersion")["value"] = item["SoftwareVersion"];
      document.getElementById("txtSoftwareDescription")["value"] = item["SoftwareDescription"];
      document.getElementById("ddlSoftwareVendor")["value"] = item["SoftwareVendor"];
    });
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
