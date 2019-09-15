import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudOperationWebPart.module.scss';
import * as strings from 'CrudOperationWebPartStrings';

import { SPHttpClientResponse, SPHttpClient, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { ICrudOperationWebPartProps } from './ICrudOperationWebPartProps';
import { IListItem } from './CrudOperationWebPartModels';

const logo: any = require('../../images/add.png');

export default class CrudOperationWebPart extends BaseClientSideWebPart<ICrudOperationWebPartProps> {

  private listItemEntityTypeName: string = undefined;

  public render(): void {

    this.domElement.innerHTML = `
      <div class="${ styles.crudOperation }">
        <div class="${ styles.container }">
          
          <div class="${ styles.row }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.description }">SharePoint list title: ${ escape(this.properties.listTitle) }</p>
              <p class="${ styles.description }">
                  Loading from '${escape(this.context.pageContext.web.title)}'
              </p>
          </div>

          <div id="userDetail" class="${ styles.row }">
            <span>Hello ${this.context.pageContext.user.displayName}</span>
            <a href="mailto:${this.context.pageContext.user.email}?Subject=Hello%20again" target="_top">
                ${this.context.pageContext.user.email}
            </a>
            <div class="status"></div>
          </div>

          <div id="getAll" class="${ styles.row }">
            <button class="${styles.button} create-Button">
              <span class="${styles.label}">Create item</span>
              <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAALEwAACxMBAJqcGAAAAaRJREFUWIXtls1KAzEQgL9VD4pnexGEtigtPoMgFqlFsb5c8VLEg+B7VKyiN38uvYjiAxSVUqFlPWSWjSW7203SenFgCOwkM19mk5kAhJ70AVjDQiIHI0vVIR5tIKLFthICY+DZFsIHwAgoaBBPeSB8AWCAKMwbwASRKoEWPHAAyIqRKEuWQXUZJ3xfnNaB6y9w8rswg8C55B/AdAhdz0Ou2/TnGQB/tyAAjoBL4FXz+wJcAA0SsuMDYBu453dnNOkNUM0CcHkTfAEtoA6URA+AU2Agcz6A3VkAdID1lCxtAF0NopIEMK0EwJ0WfHnCbvK7gvoNIXDtCnBInHbTzpP8FoGh2PZNE/OmvpUAmLaxttjOfQDULQCaYutlTUyT6EFaygkO6gCGwNClEnpp4S4A7zJuyRhMKCnfoyv45gLQlbFpsfZYxmuwPwMNWTdAFZlJSfJbBr7FVnMBCIBbWdtFFZksgFXintHJIp1GqsAncaMppswta8H7wKYPAIA9VG0PURWujToXFdET4Iw47X1gR3fg0v3y6pW+83kB9CQDNVP6fgBo7S4oS9qI5wAAAABJRU5ErkJggg==" alt="Add" title="Add user" />
            </button>
            <br/>
            <div class='myTable'>
            </div>
          </div>

          <div id="addNew" class="${ styles.row }">
            <h3 id="addNew_Label">Add new user</h3>
            <hr/>
            <table>
              <tr>
                  <td>Title</td>
                  <td><input type="text" class="newField" id="titleNew" /></td>
              </tr>
              <tr>
                  <td>Full Name</td>
                  <td><input type="text" class="newField" id="fullNameNew" /></td>
              </tr>
              <tr>
                  <td>
                      <input title="Save" type="button" value="Save" class="${styles.button} create-save-Button" />
                  </td>
                  <td>
                    <input title="Cancel" type="reset" class="${styles.button} reset-Button" value="Cancel" />
                  </td>
              </tr>
            </table>

          </div>

          <div id="update" class="${ styles.row }">
            <label id="update_Label">Edit user form</label>
            <hr/>
            <table>
              <tr>
                  <td>Title</td>
                  <td><input type="text" class="newField" id="titleEdit" /></td>
              </tr>
              <tr>
                  <td>Full Name</td>
                  <td><input type="text" class="newField" id="fullNameEdit" /></td>
              </tr>
              <tr>
                  <td>
                      <input title="Save" type="button" value="Update" class="${styles.button} update-Button" />
                  </td>
                  <td>
                    <input title="Cancel" type="reset" class="${styles.button} reset-Button" value="Cancel" />
                  </td>
              </tr>
            </table>

          </div>

        </div>
      </div>
      `;

      this.listItemEntityTypeName = undefined;
      
      this.controlFormsDisplay('#addNew, #update', '#getAll');
      this.setButtonsState();
      this.setButtonsEventHandlers();
      
      if(!this.listNotConfigured()) {
        this.updateStatus('Ready');
        this.readAllItems();
      }
      else {
        this.updateStatus('Please configure list in Web Part properties');
      }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private resetForm():void {
    const inputs : Array<HTMLInputElement> = this.domElement.querySelectorAll('input[type="text"][class="newField"]') as any as Array<HTMLInputElement>;
    for (let input of inputs) {
      input.value = '';
    }
    
    this.controlFormsDisplay('#addNew, #update', '#getAll');
  }

  private addUser(): void {
    this.updateStatus('Creating item...');
    this.getListItemEntityTypeName()
      .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
          
          const body = JSON.stringify({
              '__metadata': {
                'type': listItemEntityTypeName
              },
              'Title': (this.domElement.querySelector('#addNew #titleNew') as HTMLInputElement).value,
              'FullName': (this.domElement.querySelector('#addNew #fullNameNew') as HTMLInputElement).value
          });
          
          const restUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${escape(this.properties.listTitle)}')/items`;
          
          return this.context.spHttpClient.post(restUrl,
            SPHttpClient.configurations.v1,
            {
              headers:{
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': ''
              },
              body: body
            }
          );
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        if(response.ok) {          
          setTimeout(() => null, 0);
          return response.json();
        }
      })
      .then((item: IListItem): void => {
        this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
        this.readAllItems();
        this.resetForm();
      },(error: any): void => {
        console.error(error);
        this.updateStatus('Error while creating the item: ' + error);
      });
  }

  private createItem(): void {
    this.updateStatus('Create new item');
    this.controlFormsDisplay('#getAll, #update', '#addNew');
  }

  private getListItemEntityTypeName(): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      if (this.listItemEntityTypeName) {
        resolve(this.listItemEntityTypeName);
        return;
      }

      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${ escape(this.properties.listTitle) }')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
          return response.json();
        }, (error: any): void => {
          console.error(error);
          reject(error);
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          this.listItemEntityTypeName = response.ListItemEntityTypeFullName;
          resolve(this.listItemEntityTypeName);
        });
    });
  }

  private setButtonsEventHandlers(): void {
    const webPart : CrudOperationWebPart = this;
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.createItem(); });
    this.domElement.querySelector('input.create-save-Button').addEventListener('click', () => { webPart.addUser(); });
    this.domElement.querySelector('input.update-Button').addEventListener('click', () => { webPart.updateItem(); });
    
    (this.domElement.querySelectorAll('input.reset-Button') as any as Array<Element>)
      .forEach(
          bt => bt.addEventListener('click', () => { webPart.resetForm(); })
        );
    
  }

  private controlFormsDisplay(formsToHide: string, formsToDisplay: string): void {
    (this.domElement.querySelectorAll(formsToHide) as any as Array<HTMLElement>)
      .forEach(elem => { 
        elem.style.display = 'none'; 
      });

    (this.domElement.querySelector(formsToDisplay) as HTMLElement).style.display = 'block';
  }

  private setButtonsState(): void {
    const buttons: NodeListOf<Element> = this.domElement.querySelectorAll(`button.${styles.button}`);
    const listNotConfigured: boolean = this.listNotConfigured();

    for (let index: number = 0; index < buttons.length; index++) {
      const button: Element = buttons.item(index);
      if (listNotConfigured) {
        button.setAttribute('disabled', 'disabled');
      }
      else {
        button.removeAttribute('disabled');
      }
    }
  }

  private updateStatus(status: string): void {
    this.domElement.querySelector('.status').innerHTML = status;
  }

  private updateItemsHtml(items: IListItem[]): void {
    let webpartHtmlString : string = `
          <table class="${ styles.myTable }">
          <tr>
              <th class="${ styles.myTable_th }">ID</th>
              <th class="${ styles.myTable_th }">Title</th>
              <th class="${ styles.myTable_th }">Full Name</th>
              <th class="${ styles.myTable_th }">Action</th>
          </tr>
          `;

    if(items.length > 0) {
      webpartHtmlString +=
          items.map((item : IListItem, key: any) => `
            <tr>
              <td class="${ styles.myTable_td }">${ item.Id }</td>
              <td class="${ styles.myTable_td }">${ item.Title }</td>
              <td class="${ styles.myTable_td }">${ item.FullName }</td>
              <td class="${ styles.myTable_td }">
                <img data-id="${ item.Id }" class="imgEditListItem" src="data:image/png;base64,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" alt="Edit" title="Edit user" />
                &nbsp;
                <img data-id="${ item.Id }" class="imgDeleteListItem" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyBpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMC1jMDYwIDYxLjEzNDc3NywgMjAxMC8wMi8xMi0xNzozMjowMCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNSBXaW5kb3dzIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjNBMTNGMUFCQkU0RDExRTRBMkM1ODYwNUUwQTA2NjUyIiB4bXBNTTpEb2N1bWVudElEPSJ4bXAuZGlkOjNBMTNGMUFDQkU0RDExRTRBMkM1ODYwNUUwQTA2NjUyIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9InhtcC5paWQ6M0ExM0YxQTlCRTREMTFFNEEyQzU4NjA1RTBBMDY2NTIiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6M0ExM0YxQUFCRTREMTFFNEEyQzU4NjA1RTBBMDY2NTIiLz4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz5xIJAcAAAAd0lEQVR42mL8//8/AxqQE3ErfIgm9ujNrn55BiyAhQEHAGpghLGxGAgHjCAX+JdM/X/00h0GYoG1ngrDxp5ssAVMIOLy3acMpIBHL9+huoASwILkT2JNQglQJgYKARO+0MfFxmsAxS4YgQbAE9KApQOKkzJAgAEA0Lc04+2ethsAAAAASUVORK5CYII=" alt="Delete" title="Delete user" />
              </td>
            </tr>
          `)
          .join("");

    }
    else {
      webpartHtmlString += `
        <tr>
          <td colspan="4">No records...</td>
        </tr>
        `;
    }

    webpartHtmlString += `</table>`;
    
    this.domElement.querySelector('.myTable').innerHTML = webpartHtmlString;

    //// ES3 pattern
    //const imgs: NodeListOf<Element> = this.domElement.querySelectorAll('img.imgEditListItem');
    //for (let i: number = 0; i < imgs.length; i++) {
    //  const img: Element = imgs.item(i);
    //  const currentListItemId: string = img.getAttribute("data-id");
    //  img.addEventListener('click', (e: Event) => { this.logConsole(e, currentListItemId); });
    //}

    //// ES6 pattern - below is short hand notation
    //(this.domElement.querySelectorAll('img.imgEditListItem') as any as Array<Element>)
    //.forEach(bt => bt.addEventListener('click', (e: Event) => this.logConsole(e, "2")));

    // below is another form of above but have more statements in the func.
    (this.domElement.querySelectorAll('img.imgEditListItem') as any as Array<Element>)
      .forEach(imgIcon => {
        const currentListItemId: number = parseInt(imgIcon.getAttribute("data-id"));
        imgIcon.addEventListener('click', (e: Event) => this.editItem(e, currentListItemId));
      });

    (this.domElement.querySelectorAll('img.imgDeleteListItem') as any as Array<Element>)
      .forEach(imgIcon => {
        const currentListItemId: number = parseInt(imgIcon.getAttribute("data-id"));
        imgIcon.addEventListener('click', (e: Event) => this.deleteItem(e, currentListItemId));
      });

  }

  private editItem(event: Event, listItemId: number): void {
    this.updateStatus('Editing item');
    
    this.updateStatus(`Loading information in Edit form item ID: ${listItemId}`);
    const selectColumnsCommaSeparated = 'Title,Id,FullName';
    this.getItemById(listItemId, selectColumnsCommaSeparated)
      .then((listItemObj: IListItem): void => {
        this.updateStatus(`Item loaded ${listItemId} in edit form`);
        this.controlFormsDisplay('#getAll, #addNew', '#update');
        (this.domElement.querySelector('#update #titleEdit') as HTMLInputElement).value = listItemObj.Title;
        (this.domElement.querySelector('#update #fullNameEdit') as HTMLInputElement).value = listItemObj.FullName;
      });
  }

  private getItemById(listItemId: number, selectColumns: string): Promise<IListItem> {
    this.updateStatus(`Loading item ${listItemId} in edit form`);

    return new Promise<IListItem>((resolve: (listItemObj: IListItem) => void, reject: (error: any) => void): void => {
        let restURL: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${ escape(this.properties.listTitle) }')/items(${listItemId})?$select=${selectColumns}`;
        
        this.context.spHttpClient.get(restURL,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<IListItem> => {
            if(response.ok) {
              
              if (window.sessionStorage){
                let listItemIdkey = escape(this.properties.listTitle) + '_Id',
                    listItemEtagkey = escape(this.properties.listTitle) + '_etag';

                sessionStorage.setItem(listItemIdkey, listItemId.toString());
                sessionStorage.setItem(listItemEtagkey, response.headers.get('ETag'));
              }

              return response.json();
            }
            else {
              console.log(response);
            }
          },
          (error: any): void => {
            console.error(error);
            reject(error);
          }
        )
        .then((response: IListItem): void => {
          if(response != null) {
            resolve(response);
          }
        }, (error: any): void => {
          console.error(error);
          this.updateStatus(`Loading item ${listItemId} failed with error: ` + error);
        });
    });
    
  }

  private updateItem(): void {
    this.updateStatus('Updating item...');
    let itemId: number = undefined, eTag: string = undefined;

    let listItemIdkey = escape(this.properties.listTitle) + '_Id',
        listItemEtagkey = escape(this.properties.listTitle) + '_etag';

    this.getListItemEntityTypeName()
      .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
        
        const body = JSON.stringify({
          '__metadata': {
            'type': listItemEntityTypeName
          },
          'Title': (this.domElement.querySelector('#update #titleEdit') as HTMLInputElement).value,
          'FullName': (this.domElement.querySelector('#update #fullNameEdit') as HTMLInputElement).value
        });

        itemId = Number(sessionStorage.getItem(listItemIdkey));
        eTag = sessionStorage.getItem(listItemEtagkey);

        // Build a REST endpoint URL
        const restUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${escape(this.properties.listTitle)}')/items(${itemId})`;
        
        return this.context.spHttpClient.post(restUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': eTag,
              'X-HTTP-Method': 'MERGE'
            },
            body: body
          }
        );
      })
      .then((response: SPHttpClientResponse): void => {
        if(response.ok) {
          this.updateStatus(`Item with ID: ${itemId} successfully updated`);
          sessionStorage.removeItem(listItemIdkey);
          sessionStorage.removeItem(listItemEtagkey);

          this.readAllItems();
          this.resetForm();
        }
      }, (error: any): void => {
        console.error(error);
        this.updateStatus(`Error updating item: ${error}`);
      });
  }

  private deleteItem(event: Event, listItemId: number): void {
    
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }
    
    this.updateStatus(`Deleting item ID: ${listItemId}`);

    let itemId: number = undefined, eTag: string = undefined;
    let listItemIdkey = escape(this.properties.listTitle) + '_Id',
        listItemEtagkey = escape(this.properties.listTitle) + '_etag';

    const selectColumnsCommaSeparated = 'Id';
    this.getItemById(listItemId, selectColumnsCommaSeparated)
      .then((listItemObj: IListItem): Promise<SPHttpClientResponse> => {
        itemId = Number(sessionStorage.getItem(listItemIdkey));
        eTag = sessionStorage.getItem(listItemEtagkey);

        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listTitle}')/items(${listItemId})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': eTag,
              'X-HTTP-Method': 'DELETE'
            }
          });

      })
      .then((response: SPHttpClientResponse): void => {
        if(response.ok) {
          this.updateStatus(`Item with ID: ${listItemId} deleted successfully`);
          sessionStorage.removeItem(listItemIdkey);
          sessionStorage.removeItem(listItemEtagkey);

          this.readAllItems();
          this.resetForm();
        }
      }, (error: any): void => {
        console.error(error);
        this.updateStatus(`Error deleting item: ${error}`);
      });
  }

  private listNotConfigured(): boolean {
    return this.properties.listTitle === undefined ||
      this.properties.listTitle === null ||
      this.properties.listTitle.length === 0;
  }
  
  private readAllItems(): void {
    this.updateStatus('Loading all items...');
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${ escape(this.properties.listTitle) }')/items?$select=Title,Id,FullName`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      })
      .then((response: SPHttpClientResponse): Promise<{ value: IListItem[] }> => {
        return response.json();
      })
      .then((response: { value: IListItem[] }): void => {
        this.updateStatus(`Successfully loaded ${response.value.length} items`);
        this.updateItemsHtml(response.value);
      }, (error: any): void => {
        console.error(error);
        this.updateStatus('Loading all items failed with error: ' + error);
      });
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
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel,
                  placeholder: 'enter list title',
                  description: 'enter list title'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

}