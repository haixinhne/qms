import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import styles from "./HelloWorldWebPart.module.scss";
import * as strings from "HelloWorldWebPartStrings";
import * as XLSX from "xlsx";
import { getIdGroup, manageRoles, manageRolesFolder } from "./SetPermissions";
import {
  handleClick,
  historyLog,
  getUserName,
  displayJsonContent,
} from "./ActivityLog";
import {
  childSubFolders,
  onCountFileUpdateSharepointList,
  onProgressFiles,
  onCountFileUpdateFolders,
  onCountFileUpdateFoldersOption2,
} from "./CountFiles";

import { updateNationColumn } from "./UpdateNation";

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";

import { __metadata } from "tslib";

//Url file Excel
const excelUrl = "/sites/QMS/ProjectFolder/ADMIN/Book1.xlsx";
const sharepointUrl = "https://iscapevn.sharepoint.com/sites/QMS";
const nameSharepointSite = "QMS";
const nameSharepointList = "ProjectList";
const manageRolesValue = [
  { nameItems: "Japan-JP", groupId: 26, newRoleId: 1073741826 },
  { nameItems: "Viet Nam-VN", groupId: 25, newRoleId: 1073741826 },
];

export interface IHelloWorldWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  //DOM------------------------------------------------------------------------------------------------------------------------------------------
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }">

     <div class="${styles.welcome}">
        <h3>Management Function</h3>
     </div>

     <div class=${styles.qms_btn}>
     <button class="${
       styles.qms_button
     }" id="createUpdateProjectList">Create/Update ProjectList</button>

        <button class="${
          styles.qms_button
        }" id="createProjectFolder">Create ProjectFoder</button>

         <button class="${
           styles.qms_button
         }" id="setFolderPermissions">Set Folder Permissions</button>

         <button class="${
           styles.qms_button
         }" id="setListPermissions">Set List Permissions</button>


         <button class="${
           styles.qms_button
         }" id="updateProgressProject_Op1">Update Progress Project Op1</button>

          <button class="${
            styles.qms_button
          }" id="updateProgressProject_Op2">Update Progress Project Op2</button>
         </div>

     <div class="${styles.qms_actions}" id= "qms_actions">
     <p id= "qms_desc"></p>
     </div>
     </section>`;

    const clickCreateSharepoint = this.domElement.querySelector(
      "#createUpdateProjectList"
    );
    const clickCreateFolder = this.domElement.querySelector(
      "#createProjectFolder"
    );
    const setFolderPermissions = this.domElement.querySelector(
      "#setFolderPermissions"
    );
    const setListPermissions = this.domElement.querySelector(
      "#setListPermissions"
    );
    const actionsContainer = this.domElement.querySelector("#qms_actions");
    const clickCountFilesOp1 = this.domElement.querySelector(
      "#updateProgressProject_Op1"
    );
    const clickCountFilesOp2 = this.domElement.querySelector(
      "#updateProgressProject_Op2"
    );

    if (!actionsContainer) {
      console.error("Error");
      return;
    }

    //Event click button-------------------------------------------------------------------------------------------------------------------------
    //Tạo sharepoint
    if (clickCreateSharepoint) {
      clickCreateSharepoint.addEventListener("click", () => {
        this.onClickButtonCreateSharepoint();
        handleClick(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointSite,
          "Create/Update ProjectList"
        );
      });
    }

    //Tạo folder
    if (clickCreateFolder) {
      clickCreateFolder.addEventListener("click", () => {
        this.onCreateFolder();
        handleClick(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointSite,
          "Create ProjectFoder"
        );
        this.onUpdateNationColumn();
      });
    }

    //Set Permissions
    //Folder
    if (setFolderPermissions) {
      setFolderPermissions.addEventListener("click", () => {
        this.getDataFromSharePointList().then((folderData) => {
          if (!folderData || folderData.length === 0) {
            console.warn("Error");
            return;
          }

          const allPromises: Promise<void>[] = [];

          folderData.forEach(({ subFolderName, folderName }) => {
            const folderPath = `/sites/QMS/ProjectFolder/PROJECT/${subFolderName}`;
            const roleInfo = [];

            for (let i = 0; i < manageRolesValue.length; i++) {
              if (manageRolesValue[i].nameItems === folderName) {
                roleInfo.push({
                  groupId: manageRolesValue[i].groupId,
                  newRoleId: manageRolesValue[i].newRoleId,
                });
                break;
              }
            }

            if (roleInfo.length === 0) {
              console.warn(`Error: ${folderName}`);
            }

            roleInfo.push({ groupId: 36, newRoleId: 1073741829 });

            roleInfo.forEach(({ groupId, newRoleId }) => {
              allPromises.push(
                manageRolesFolder(
                  this.context.spHttpClient,
                  sharepointUrl,
                  folderPath,
                  groupId,
                  newRoleId,
                  this.context.pageContext.legacyPageContext.formDigestValue
                )
              );
            });
          });

          Promise.all(allPromises)
            .then(() => {
              alert("The folder permissions were set successfully");
              historyLog(
                this.context.spHttpClient,
                sharepointUrl,
                nameSharepointSite,
                "The folder permissions were set successfully"
              );
            })
            .catch((error) => {
              console.error("Error:", error);
            });

          handleClick(
            this.context.spHttpClient,
            sharepointUrl,
            nameSharepointSite,
            "Set Folder Permissions"
          );
        });
      });
    }

    //Sharepoint list
    if (setListPermissions) {
      setListPermissions.addEventListener("click", () => {
        getIdGroup(this.context.spHttpClient, sharepointUrl);

        const allPromises: Promise<void>[] = [];

        manageRolesValue.forEach(({ nameItems, groupId, newRoleId }) => {
          const roleInfo = [
            { groupId, newRoleId },
            { groupId: 36, newRoleId: 1073741829 },
          ];

          roleInfo.forEach(({ groupId, newRoleId }) => {
            allPromises.push(
              manageRoles(
                this.context.spHttpClient,
                sharepointUrl,
                nameSharepointList,
                nameItems,
                groupId,
                newRoleId,
                this.context.pageContext.legacyPageContext.formDigestValue
              )
            );
          });
        });

        Promise.all(allPromises)
          .then(() => {
            alert("The list permissions were set successfully");
            historyLog(
              this.context.spHttpClient,
              sharepointUrl,
              nameSharepointSite,
              "The list permissions were set successfully"
            );
          })
          .catch((error) => {
            console.error(
              "An error occurred while setting permissions:",
              error
            );
          });

        handleClick(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointSite,
          "Set List Permissions"
        );
      });
    }

    //Count files
    if (clickCountFilesOp1) {
      clickCountFilesOp1.addEventListener("click", () => {
        onCountFileUpdateSharepointList(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList
        );
        onProgressFiles(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList
        );
        onCountFileUpdateFolders(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList
        );
        handleClick(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointSite,
          "Update Progress Project Op1"
        );
        historyLog(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointSite,
          `The Progress column was updated successfully in ${nameSharepointList} and ProjectFolder`
        );
      });
    }

    if (clickCountFilesOp2) {
      clickCountFilesOp2.addEventListener("click", () => {
        onCountFileUpdateSharepointList(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList
        );
        onProgressFiles(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList
        );
        onCountFileUpdateFoldersOption2(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList
        );
        handleClick(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointSite,
          "Update Progress Project Op2"
        );
        historyLog(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointSite,
          `The Progress column was updated successfully in ${nameSharepointList}`
        );
      });
    }

    this.renderListAsync();
  }

  //Hàm defaults--------------------------------------------------------------------------------------------------------------------------------
  protected onInit(): Promise<any> {
    return this.getEnvironmentMessage().then(() => {
      return getUserName(this.context.spHttpClient, sharepointUrl).then(() => {
        displayJsonContent(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointSite
        );
      });
    });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = "";
    items.forEach((item: ISPList) => {
      html += `
  <ul class="${styles.list}">
    <li class="${styles.listItem}">
      <span class="ms-font-l">${item.Title}</span>
    </li>
  </ul>`;
    });

    if (this.domElement.querySelector("#spListContainer") != null) {
      this.domElement.querySelector("#spListContainer")!.innerHTML = html;
    }
  }

  private renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      })
      .catch(() => {});
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .catch(() => {});
  }

  //Tạo sharepoint list từ excel, add, update, xóa items từ sharepoint---------------------------------------------------------------------------
  //Lấy file excel
  private getFileExcelFromSharePoint(excelUrl: string): Promise<ArrayBuffer> {
    return this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${excelUrl}')/$value`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.arrayBuffer();
      })
      .catch((error) => {
        console.error("Error:", error);
        return Promise.reject(error);
      });
  }

  //Đọc nội dung file excel (lấy tên các cột)
  private readFileExcelFromSharePoint(fileContent: ArrayBuffer): {
    nameColumnSharepoint: string[];
    nameItems: Record<string, any>[];
  } {
    const data = new Uint8Array(fileContent);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const nameColumnSharepoint = (jsonData[0] as any[]).filter(Boolean);
    const nameItems = jsonData.slice(1).map((row: any[]) => {
      const rowObject: Record<string, any> = {};
      nameColumnSharepoint.forEach((colName, index) => {
        rowObject[colName] = row[index] || null;
      });
      return rowObject;
    });

    return { nameColumnSharepoint, nameItems };
  }

  //Check sharepoint list đã tồn tại
  private checkNameSharepointList(listName: string): Promise<boolean> {
    return this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return true;
        } else {
          console.log(`The list does not exist: ${listName}`);
          return false;
        }
      })
      .catch((error) => {
        console.error("Error:", error);
        return false;
      });
  }

  //Tạo sharepoint list
  private async createSharePointList(listName: string): Promise<any> {
    const listNameExists = await this.checkNameSharepointList(listName);
    if (listNameExists) {
      console.log(`The list already exists: ${listName}`);
      alert(`The list already exists: ${listName}`);
      return;
    }
    const body = JSON.stringify({
      __metadata: { type: "SP.List" },
      BaseTemplate: 100,
      Title: listName,
    });
    const optionsHTTP: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "",
      },
      body: body,
    };

    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log(`The list was created successfully: ${listName}`);
          alert(`The list was created successfully: ${listName}`);
          historyLog(
            this.context.spHttpClient,
            sharepointUrl,
            nameSharepointSite,
            `The list was created successfully: ${listName}`
          );
          return response.json();
        } else {
          return response.json().then((errorResponse) => {
            console.error("Error:", errorResponse);
          });
        }
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }

  //Check các cột đã tồn tại ở sharepoint
  private async getExistingColumns(listName: string): Promise<string[]> {
    const response = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/fields?$select=Title`,
      SPHttpClient.configurations.v1
    );

    if (response.ok) {
      const data = await response.json();
      return data.value.map((field: { Title: string }) => field.Title);
    } else {
      console.error("Error");
      return [];
    }
  }

  //Tạo cột sharepoint list
  private async createColumnInSharePoint(
    listName: string,
    columnNames: string
  ): Promise<any> {
    const existingColumns = await this.getExistingColumns(listName);

    if (existingColumns.indexOf(columnNames) !== -1) {
      console.log(`The columns were already exists: ${columnNames}`);
      return;
    }

    const body = JSON.stringify({
      __metadata: { type: "SP.Field" },
      Title: columnNames,
      FieldTypeKind: 2,
    });

    const optionsHTTP: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "",
      },
      body: body,
    };

    return await this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/fields`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log(`The columns were created successfully: ${columnNames}`);
          return response.json();
        } else {
          return response.json().then((errorResponse) => {
            console.error("Error:", errorResponse);
          });
        }
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }

  //Tạo items, update items
  private async createItemsInSharePointList(
    listName: string,
    itemData: any
  ): Promise<void> {
    const capsLocksFirstLetter = (text: string): string => {
      return text.charAt(0).toUpperCase() + text.slice(1);
    };
    const listNameUpdate = capsLocksFirstLetter(listName);

    //Check sự tồn tại của item dựa vào cột CustomID
    const checkExistingItem = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items?$filter=CustomID eq '${itemData.CustomID}'`,
      SPHttpClient.configurations.v1
    );

    const existingItems = await checkExistingItem.json();
    const saveExistingItem = existingItems.value && existingItems.value[0];

    let endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items`;
    let method = "POST"; // default là để tạo mới

    if (saveExistingItem) {
      endpoint += `(${saveExistingItem.Id})`;
      method = "MERGE"; //update
    }

    const bodyObject: Record<string, any> = {
      __metadata: {
        type: `SP.Data.${listNameUpdate}ListItem`,
      },
      Title: itemData.Title || nameSharepointList,
    };

    let hasChanges = false;
    let createdFields: string[] = [];
    let updatedFields: string[] = [];
    //Lặp qua các itemData để xem các thay đổi (True-False)
    for (const key in itemData) {
      if (itemData.hasOwnProperty(key)) {
        const newValue = String(itemData[key] || "");
        const existingValue = saveExistingItem
          ? String(saveExistingItem[key] || "")
          : "";

        if (newValue !== existingValue) {
          bodyObject[key] = newValue;
          hasChanges = true;

          if (saveExistingItem) {
            updatedFields.push(key); //Update
          } else {
            createdFields.push(key); //Create
          }
        }
      }
    }

    if (!hasChanges && method === "MERGE") {
      console.log("No items were changed");
      return;
    }

    const body = JSON.stringify(bodyObject);
    const optionsHTTP: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "",
        "If-Match": "*",
        "X-HTTP-Method": method,
      },
      body: body,
    };

    return await this.context.spHttpClient
      .post(endpoint, SPHttpClient.configurations.v1, optionsHTTP)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          if (method === "POST") {
            console.log(
              `The item was created successfully: ${itemData.CustomID}`
            );
            alert(`The item was created successfully: ${itemData.CustomID}`);
            historyLog(
              this.context.spHttpClient,
              sharepointUrl,
              nameSharepointSite,
              `The item was created successfully: ${itemData.CustomID}`
            );
            if (createdFields.length > 0) {
              `The column was created successfully: ${createdFields.join(
                ", "
              )}`;
            }
          } else if (method === "MERGE") {
            console.log(
              `The item was updated successfully: ${itemData.CustomID}`
            );
            alert(`The item was updated successfully: ${itemData.CustomID}`);
            historyLog(
              this.context.spHttpClient,
              sharepointUrl,
              nameSharepointSite,
              `The item was updated successfully: ${itemData.CustomID}`
            );
            if (updatedFields.length > 0) {
              console.log(
                `The column was updated successfully: ${updatedFields.join(
                  ", "
                )}`
              );
              alert(
                `The column was updated successfully: ${updatedFields.join(
                  ", "
                )}`
              );
            }
          }
        } else {
          response
            .json()
            .then((errorResponse) => {
              console.error("Error:", errorResponse);
            })
            .catch((jsonError) => {
              console.error("Error:", jsonError);
            });
        }
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }

  //Xóa Items ở SharePoint list
  private deleteItemFromSharePoint(listName: string, item: any): void {
    const deleteEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items(${item.Id})`;
    const optionsHTTP: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "",
        "If-Match": "*",
        "X-HTTP-Method": "DELETE",
      },
    };

    this.context.spHttpClient
      .post(deleteEndpoint, SPHttpClient.configurations.v1, optionsHTTP)
      .then((deleteResponse) => {
        if (deleteResponse.ok) {
          console.log(`The item was deleted successfully: ${item.CustomID}`);
          alert(`The item was deleted successfully: ${item.CustomID}`);
          historyLog(
            this.context.spHttpClient,
            sharepointUrl,
            nameSharepointSite,
            `The item was deleted successfully: ${item.CustomID}`
          );
        } else {
          deleteResponse
            .json()
            .then((errorResponse) => {
              console.error("Error:", errorResponse);
            })
            .catch((jsonError) => {
              console.error("Error:", jsonError);
            });
        }
      })
      .catch((error) => {
        console.error("Error", error);
      });
  }

  //Click Tạo sharepoint list, tạo cột, tạo mới, update, xóa items
  private onClickButtonCreateSharepoint(): void {
    if (!nameSharepointList) {
      alert("Please enter a name for the SharePoint list!");
      return;
    }
    //Tạo sharepoint list
    this.createSharePointList(nameSharepointList)
      .then(() => {
        nameSharepointList;
      })
      //Tạo cột từ file excel
      .then(() => this.getFileExcelFromSharePoint(excelUrl))
      .then((fileContent: ArrayBuffer) => {
        const { nameColumnSharepoint, nameItems } =
          this.readFileExcelFromSharePoint(fileContent);
        return nameColumnSharepoint
          .reduce((promise, createColumn) => {
            return promise.then(() => {
              return this.createColumnInSharePoint(
                nameSharepointList,
                createColumn
              ).catch((error) => {
                console.error("Error:", error);
              });
            });
          }, Promise.resolve())
          .then(() => ({
            nameColumnSharepoint,
            nameItems,
          }));
      })

      //Tạo items, update items
      .then(({ nameColumnSharepoint, nameItems }) => {
        console.log("Input data from Excel:", nameItems);

        return nameItems
          .reduce((promise, itemData) => {
            const itemObject = nameColumnSharepoint.reduce(
              (obj, columnName) => {
                obj[columnName] = itemData[columnName] || null;
                return obj;
              },
              {} as Record<string, any>
            );
            return promise.then(() =>
              this.createItemsInSharePointList(nameSharepointList, itemObject)
            );
          }, Promise.resolve())
          .then(() => ({
            nameColumnSharepoint,
            nameItems,
          }));
      })

      //Xóa items
      .then(({ nameItems }) => {
        const existingItemsFromExcel = new Set(
          nameItems.map((item: any) => item.CustomID)
        );
        return this.context.spHttpClient
          .get(
            `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items`,
            SPHttpClient.configurations.v1
          )
          .then((response) => {
            if (!response.ok) {
              console.log(`Error: ${response.statusText}`);
            }
            return response.json();
          })
          .then((existingItems) => {
            const itemsDelete = existingItems.value.filter(
              (item: any) => !existingItemsFromExcel.has(item.CustomID)
            );
            return Promise.all(
              itemsDelete.map((item: any) => {
                return this.deleteItemFromSharePoint(nameSharepointList, item);
              })
            );
          });
      })

      .catch((error) => {
        console.error("Error:", error);
      });
  }

  //Tạo các thư mục từ sharepoint list-------------------------------------------------------------------------------------------------------------
  //Hàm lấy data từ sharepoint list
  private getDataFromSharePointList(): Promise<
    { folderName: string; subFolderName: string; customId: string }[]
  > {
    return this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data) => {
        const folderValues = data.value
          //Tên folder con = tên cột ProjectName
          .filter(
            (item: any) => item.Nation && item.ProjectName && item.CustomID
          )
          .map((item: any) => ({
            folderName: item.Nation,
            subFolderName: item.ProjectName,
            customId: item.CustomID,
          }))
          .filter(
            (name: any, index: Number, self: any) =>
              self.findIndex(
                (i: any) => i.subFolderName === name.subFolderName
              ) === index
          );
        console.log("Input data from Sharepoint list:", folderValues);

        return folderValues;
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }

  // Hàm tạo subfolder
  private createSubfolder(subFolderName: string): Promise<any> {
    const optionsHTTP: ISPHttpClientOptions = {
      headers: {
        accept: "application/json; odata=verbose",
        "content-type": "application/json; odata=verbose",
        "odata-version": "",
      },
    };
    const subFolderUrl = `ProjectFolder/PROJECT/${subFolderName}`;
    const subFolders = ["Promotion", "Design", "Build"];
    const arrayFolderUrl: string[] = [];
    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('ProjectFolder/PROJECT/${subFolderName}')`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then(() => {
        // Tạo các thư mục con
        return subFolders.reduce((prevPromise, folder) => {
          return prevPromise.then(() => {
            const folderUrl = `${subFolderUrl}/${folder}`;
            return this.context.spHttpClient
              .post(
                `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('${folderUrl}')`,
                SPHttpClient.configurations.v1,
                optionsHTTP
              )
              .then(() => {
                // Tạo các thư mục con nhỏ hơn
                const childFolders = childSubFolders[folder];
                return childFolders.reduce((childPrevPromise, childFolder) => {
                  const childFolderName = childFolder.name;
                  const childFolderUrl = `${folderUrl}/${childFolderName}`;
                  arrayFolderUrl.push(childFolderUrl);
                  return childPrevPromise.then(() => {
                    return this.context.spHttpClient.post(
                      `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('${childFolderUrl}')`,
                      SPHttpClient.configurations.v1,
                      optionsHTTP
                    );
                  });
                }, Promise.resolve());
              });
          });
        }, Promise.resolve());
      })

      .catch((error) => {
        console.error("Error", error);
      });
  }

  //Click tạo folder
  private onCreateFolder(): Promise<any> {
    return this.getDataFromSharePointList()
      .then((folderPairs) => {
        const subFolder = folderPairs.map(({ subFolderName }) => subFolderName);
        const createSubFolder = subFolder.map((subFolderName) =>
          this.createSubfolder(subFolderName)
        );

        return Promise.all(createSubFolder);
      })
      .then(() => {
        console.log("The folders were created successfully");
        alert("The folders were created successfully");
      })
      .catch((error) => {
        console.error("Error", error);
      });
  }

  //Update cột Nation ---------------------------------------------------------------------------------------------------------------------------
  private onUpdateNationColumn(): Promise<any> {
    return this.getDataFromSharePointList()
      .then((folderValues) => {
        if (!folderValues || folderValues.length === 0) {
          return Promise.reject("No folder data found");
        }

        const updatePromises = folderValues.map(
          ({ folderName, subFolderName, customId }) => {
            return updateNationColumn(
              this.context.spHttpClient,
              sharepointUrl,
              subFolderName,
              folderName,
              customId
            );
          }
        );

        return Promise.all(updatePromises);
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }

  //Defaults-------------------------------------------------------------------------------------------------------------------------------------------------------
  private getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      //running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
