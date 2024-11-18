import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from "@microsoft/sp-lodash-subset";
import styles from "./HelloWorldWebPart.module.scss";
import * as strings from "HelloWorldWebPartStrings";
import * as XLSX from "xlsx";

//Hải add
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { __metadata } from "tslib";

//Url file Excel
const excelUrl = "/sites/QMS/Shared Documents/Book1.xlsx";
const sharepointUrl = "https://iscapevn.sharepoint.com/sites/QMS";
const nameSharepointList = "QMS";

export interface IHelloWorldWebPartProps {
  description: string;
}

//Hải add
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _environmentMessage: string = "";
  //DOM-------------------------------------------------------------------------------------------------------------------------------------------------------------
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }">

     <div class="${styles.welcome}">    
        <h2>Hello, ${escape(this.context.pageContext.user.displayName)}</h2>
       <div>${this._environmentMessage}</div>
       
        
        </div>
     </div>

     <div class=${styles.qms_btn}>
     <button class="${
       styles.qms_button
     }" id="createSharepointList">Create Sharepoint</button>

        <button class="${
          styles.qms_button
        }" id="createFolder">Create Folder</button>        

         <button class="${
           styles.qms_button
         }" id="setPermissions">Set Permissions</button></div>
     <div class="${styles.qms_actions}" id= "qms_actions">
     <p id= "qms_desc"></p>
     </div>
     </section>`;

    const clickCreateSharepoint = this.domElement.querySelector(
      "#createSharepointList"
    );
    const clickCreateFolder = this.domElement.querySelector("#createFolder");
    const setPermissions = this.domElement.querySelector("#setPermissions");
    const actionsContainer = this.domElement.querySelector("#qms_actions");

    if (!actionsContainer) {
      console.error("The actionsContainer element was not found.");
      return;
    }

    //Display----------------------------------------------------------------------------------------------------------------------------------------------------
    //Hàm Save file json vào thư mục mỗi khi click vào 1 nút
    const saveJsonSharePoint = (
      folderPath: string,
      fileName: string,
      jsonData: string
    ) => {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Files/add(url='${fileName}',overwrite=true)`;
      this.context.spHttpClient
        .post(url, SPHttpClient.configurations.v1, {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "",
          },
          body: jsonData,
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            console.log("File saved successfully");
          } else {
            response.json().then((error) => {
              console.error("Error saving file:", error);
            });
          }
        });
    };

    //Hiển thị nội dung từ file Json
    const displayJsonContent = () => {
      const folderPath = `/sites/${nameSharepointList}/Shared Documents/ActivityHistory`;
      const fileName = "activityLog.json";
      const fileUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${folderPath}/${fileName}')/$value`;

      this.context.spHttpClient
        .get(fileUrl, SPHttpClient.configurations.v1, {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "",
          },
        })
        .then((response) => {
          if (response.ok) {
            return response.text();
          } else {
            return Promise.reject(
              `Error retrieving file. Status: ${response.status}, ${response.statusText}`
            );
          }
        })
        .then((jsonContent) => JSON.parse(jsonContent))
        .then((parsedContent) => {
          if (!Array.isArray(parsedContent)) {
            return Promise.reject("JSON content is not an array.");
          }

          const contentContainer = document.getElementById("qms_actions");
          if (contentContainer) {
            contentContainer.innerHTML = "";
            parsedContent.reverse().forEach((item) => {
              const paragraph = document.createElement("p");
              paragraph.className = "qms_desc";
              paragraph.textContent = item;
              contentContainer.appendChild(paragraph);
            });
          } else {
            return Promise.reject("Container element not found!");
          }
        })
        .catch((error) => {
          console.error("Error processing JSON file:", error);
        });
    };

    //Hàm tạo nội dung khi click
    const handleClick = (buttonName: string) => {
      this.getUserName().then((userName) => {
        const getTimestamp = new Date().toLocaleString();
        const getMessage = `${getTimestamp}: ${userName} clicked the ${buttonName} button`;
        const folderPath = `/sites/${nameSharepointList}/Shared Documents/ActivityHistory`;
        const fileName = "activityLog.json";
        const fileUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${folderPath}/${fileName}')/$value`;

        this.context.spHttpClient
          .get(fileUrl, SPHttpClient.configurations.v1, {
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "odata-version": "",
            },
          })
          .then((response) => {
            if (response.ok) {
              return response.text();
            } else if (response.status === 404) {
              console.error("File not found. Creating a new one.");
              return "[]";
            } else {
              return Promise.reject(
                `Error retrieving file. Status: ${response.status}, ${response.statusText}`
              );
            }
          })
          .then((existingContent) => {
            return Promise.resolve()
              .then(() => JSON.parse(existingContent))
              .catch((error) => {
                console.error("Error parsing existing JSON content:", error);
                return [];
              })
              .then((currentData) => {
                currentData.push(getMessage);
                const updatedJson = JSON.stringify(currentData, null, 1);
                saveJsonSharePoint(folderPath, fileName, updatedJson);
                displayJsonContent();
              });
          })
          .catch((error) => {
            console.error("Error processing JSON file:", error);
          });
      });
    };

    //Event click button
    //Tạo sharepoint
    if (clickCreateSharepoint) {
      clickCreateSharepoint.addEventListener("click", () => {
        this.onClickButtonCreateSharepoint();
        handleClick("Create Sharepoint");
        displayJsonContent();
      });
    } else {
      console.warn("clickCreateSharepoint element not found.");
    }

    //Tạo folder
    if (clickCreateFolder) {
      clickCreateFolder.addEventListener("click", () => {
        this.onCreateFolder();
        handleClick("Create Folder");
      });
    } else {
      console.warn("clickCreateFolder element not found.");
    }

    //Set Permissions
    if (setPermissions) {
      setPermissions.addEventListener("click", () => {
        this.getIDGroup();

        const manageRolesValue = [
          { nameItems: "Vietnam", groupId: 25, newRoleId: 1073741826 },
          { nameItems: "Japan", groupId: 26, newRoleId: 1073741826 },
          { nameItems: "USA", groupId: 30, newRoleId: 1073741826 },
        ];
        manageRolesValue.forEach(({ nameItems, groupId, newRoleId }) => {
          this.manageRoles(nameItems, groupId, newRoleId);
        });

        handleClick("Set Permissions");
      });
    }

    this.renderListAsync();
  }

  //Lấy tên user name
  private getUserName(): Promise<any> {
    return this.context.spHttpClient
      .get(
        `${sharepointUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data) => {
        const userName = data.Title;
        return userName;
      });
  }

  //Hàm defaults--------------------------------------------------------------------------------------------------------------------------------------------------
  protected onInit(): Promise<void> {
    return this.getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
      return this.getUserName();
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

  //Tạo sharepoint list từ excel----------------------------------------------------------------------------------------------------------------------------------
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
        console.error("Error fetching file:", error);
        throw error;
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
          console.log(`List '${listName}' exists.`);
          return true;
        } else {
          console.log(`List '${listName}' does not exist.`);
          return false;
        }
      })
      .catch((error) => {
        console.error("Error checking list existence:", error);
        return false;
      });
  }

  //Tạo SharePoint list
  private async createSharePointList(listName: string): Promise<any> {
    const listNameExists = await this.checkNameSharepointList(listName);
    if (listNameExists) {
      alert(`${listName} already exists`);
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
          console.log(`Created sharepoint list: ${listName}`);
          alert(`Created sharepoint list: ${listName}`);
          return response.json();
        } else {
          return response.json().then((errorResponse) => {
            console.error("Error response:", errorResponse);
          });
        }
      })
      .catch((error) => {
        console.error("Error creating:", error);
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
      console.error("Error fetching columns.");
      return [];
    }
  }

  //Tạo cột Sharepoint list
  private async createColumnInSharePoint(
    listName: string,
    columnNames: string
  ): Promise<any> {
    const existingColumns = await this.getExistingColumns(listName);

    if (existingColumns.indexOf(columnNames) !== -1) {
      console.log(`Column "${columnNames}" already exists.`);
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
          console.log(`New column created: ${columnNames}`);
          return response.json();
        } else {
          return response.json().then((errorResponse) => {
            console.error("Error response:", errorResponse);
          });
        }
      })
      .catch((error) => {
        console.error("Error adding column:", error);
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
    //Lưu kết quả các items nếu tồn tại
    const existingItems = await checkExistingItem.json();
    const saveExistingItem = existingItems.value && existingItems.value[0];

    //Nếu item đã tồn tại thì phương thức là update, nếu ko thì tạo mới
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

    // for (const key in itemData) {
    //   if (itemData.hasOwnProperty(key)) {
    //     bodyObject[key] = String(itemData[key] || "");
    //   }
    // }

    let hasChanges = false;
    let createdFields: string[] = [];
    let updatedFields: string[] = [];

    // Loop over itemData to determine changes
    for (const key in itemData) {
      if (itemData.hasOwnProperty(key)) {
        const newValue = String(itemData[key] || "");
        const existingValue = saveExistingItem
          ? String(saveExistingItem[key] || "")
          : "";

        // If the value has changed, update or create the field
        if (newValue !== existingValue) {
          bodyObject[key] = newValue;
          hasChanges = true;

          if (saveExistingItem) {
            updatedFields.push(key); // Track updated fields
          } else {
            createdFields.push(key); // Track newly created fields
          }
        }
      }
    }

    if (!hasChanges && method === "MERGE") {
      console.log(
        `No changes detected for item with CustomID = ${itemData.CustomID}`
      );
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
            console.log(`Item created: CustomID = ${itemData.CustomID}`);
            alert(`Item created: ${itemData.CustomID}`);
            if (createdFields.length > 0) {
              console.log(`Created fields: ${createdFields.join(", ")}`);
            }
          } else if (method === "MERGE") {
            console.log(`Item updated: CustomID = ${itemData.CustomID}`);
            alert(`Item updated: ${itemData.CustomID}`);
            if (updatedFields.length > 0) {
              console.log(`Updated fields: ${updatedFields.join(", ")}`);
              alert(`Updated fields: ${updatedFields.join(", ")}`);
            }
          }
        } else {
          response
            .json()
            .then((errorResponse) => {
              console.error("Error response:", errorResponse);
            })
            .catch((jsonError) => {
              console.error("Error parsing response:", jsonError);
            });
        }
      })
      .catch((error) => {
        console.error("Error adding or updating item:", error);
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
          console.log(`Item deleted from SharePoint: ${item.CustomID}`);
          alert(`Item deleted from SharePoint: ${item.CustomID}`);
        } else {
          deleteResponse
            .json()
            .then((errorResponse) => {
              console.error(
                "Error response while deleting item from SharePoint:",
                errorResponse
              );
            })
            .catch((jsonError) => {
              console.error("Error parsing response:", jsonError);
            });
        }
      })
      .catch((error) => {
        console.error("Error deleting item from SharePoint:", error);
      });
  }

  //Click Tạo SharePoint list, tạo cột, tạo mới, update, xóa items
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
                console.error(`Error adding column "${createColumn}":`, error);
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
        console.log("Input data", nameItems);

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
              throw new Error(
                `Failed to retrieve items: ${response.statusText}`
              );
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

  //Tạo các thư mục con từ sharepoint------------------------------------------------------------------------------------------------------------------------------
  // Hàm lấy data từ SharePoint list (Lấy tên thư mục là 1 cột ở sharepoint list)
  private getFileFromSharePoint(): Promise<
    { folderName: string; subFolderName: string }[]
  > {
    const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items`;

    return this.context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: any) => {
        const folderValues = data.value
          .filter((item: any) => item.Note && item.Brand)
          .map((item: any) => ({
            folderName: item.Note,
            subFolderName: item.Brand,
          }))
          .filter(
            (name: any, index: Number, self: any) =>
              self.findIndex(
                (p: any) =>
                  p.folderName === name.folderName &&
                  p.subFolderName === name.subFolderName
              ) === index
          );

        console.log(`Folder name: ${folderValues}`);
        return folderValues;
      })
      .catch((error) => {
        console.error("Error fetching SharePoint list data:", error);
        throw error;
      });
  }

  // Hàm tạo subfolder
  private createSubfolder(
    parentFolderName: string,
    subFolderName: string
  ): Promise<any> {
    const optionsHTTP: ISPHttpClientOptions = {
      headers: {
        accept: "application/json; odata=verbose",
        "content-type": "application/json; odata=verbose",
        "odata-version": "",
      },
    };
    const subFolderPath = `Shared Documents/${parentFolderName}/${subFolderName}`;

    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('Shared Documents/${parentFolderName}/${subFolderName}')`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then(() => {
        // Create the "00_ESSENTIAL" folder within each subfolder
        return this.context.spHttpClient.post(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('${subFolderPath}/00_ESSENTIAL')`,
          SPHttpClient.configurations.v1,
          optionsHTTP
        );
      })
      .then(() => {
        console.log(`Created 00_ESSENTIAL in: ${subFolderPath}`);
      })
      .then(() => {
        // Create the "01_WORK" folder within each subfolder
        return this.context.spHttpClient.post(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('${subFolderPath}/01_WORK')`,
          SPHttpClient.configurations.v1,
          optionsHTTP
        );
      })
      .then(() => {
        console.log(`Created 01_WORK in: ${subFolderPath}`);
      })
      .then(() => {
        // Create the "02_SUBMIT" folder within each subfolder
        return this.context.spHttpClient.post(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('${subFolderPath}/02_SUBMIT')`,
          SPHttpClient.configurations.v1,
          optionsHTTP
        );
      })
      .then(() => {
        console.log(`Created 02_SUBMIT in: ${subFolderPath}`);
      })

      .catch((error) => {
        console.error("Error creating subfolder:", error);
        throw error;
      });
  }

  //Hàm tạo folder
  private createFolder(
    folderName: string,
    subFolderNames: string[]
  ): Promise<any> {
    const optionsHTTP: ISPHttpClientOptions = {
      headers: {
        accept: "application/json; odata=verbose",
        "content-type": "application/json; odata=verbose",
        "odata-version": "",
      },
    };

    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('Shared Documents/${folderName}')`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => response.json())
      .then(() => {
        console.log(`Folder created: ${folderName}`);
        alert(`Folder created: ${folderName}`);
        return Promise.all(
          subFolderNames.map((subFolderName) =>
            this.createSubfolder(folderName, subFolderName).then(() => {
              console.log(
                `Subfolder created: ${subFolderName} in ${folderName}`
              );
              alert(`Subfolder created: ${subFolderName} in ${folderName}`);
            })
          )
        );
      })
      .catch((error) => {
        console.error("Error creating folder or subfolders:", error);
        throw error;
      });
  }

  //Click Tạo folder
  private onCreateFolder(): Promise<any> {
    return this.getFileFromSharePoint()
      .then((folderSubfolderPairs) => {
        // Group subfolder names by their folder
        const folderMap = folderSubfolderPairs.reduce((acc, pair) => {
          if (!acc[pair.folderName]) {
            acc[pair.folderName] = [];
          }
          acc[pair.folderName].push(pair.subFolderName);
          return acc;
        }, {} as Record<string, string[]>);

        // Create folders with their respective subfolders
        const promises = [];
        for (const folderName in folderMap) {
          if (folderMap.hasOwnProperty(folderName)) {
            const subFolderNames = folderMap[folderName];
            promises.push(this.createFolder(folderName, subFolderNames));
          }
        }

        return Promise.all(promises);
      })
      .catch((error) => {
        console.error("Error processing folders and subfolders:", error);
      });
  }

  //Set Permissions-----------------------------------------------------------------------------------------------------------------------------------------------
  //Lấy ID group
  private getIDGroup() {
    this.context.spHttpClient
      .get(
        `${sharepointUrl}/_api/web/sitegroups`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data) => {
        const groups: { Title: string; Id: number }[] = data.value;
        groups.forEach((group: { Title: string; Id: number }) => {
          console.log(`Group Name: ${group.Title}, ID: ${group.Id}`);
        });
      })
      .catch((error) => console.error("Error fetching groups:", error));
  }

  //Lấy ID của item dựa trên tên giá trị ở cột Note
  private getItemId(nameItems: string): Promise<number[]> {
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items?$filter=Note eq '${nameItems}'&$select=ID`;
    return this.context.spHttpClient
      .get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return Promise.reject("Failed to retrieve item ID");
        }
        return response.json();
      })
      .then((data) => {
        if (data.value && data.value.length > 0) {
          const itemId = data.value.map((item: { ID: number }) => item.ID);
          return itemId;
        } else {
          return Promise.reject("Item not found");
        }
      });
  }

  //Gửi yêu cầu tới Sharepoint
  private executeRequest(url: string, method: string): Promise<void> {
    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=verbose",
          "X-RequestDigest":
            this.context.pageContext.legacyPageContext.formDigestValue,
        },
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.json().then((errorDetails) => {
            console.error("Error details:", errorDetails);
            return Promise.reject(
              "Request failed: " + errorDetails.error.message.value
            );
          });
        }
        return Promise.resolve();
      });
  }

  // Ngắt quyền kế thừa của mục
  private breakRoleInheritanceForItem(itemId: number): Promise<number> {
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/breakroleinheritance(true)`;
    return this.executeRequest(requestUrl, "POST").then(() => {
      console.log(
        `Break role inheritance for item ID: ${itemId} successfully!`
      );
      return itemId;
    });
  }

  // Xóa vai trò của nhóm hiện tại khỏi mục
  private removeCurrentRoleFromItem(
    itemId: number,
    groupId: number
  ): Promise<number> {
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/roleassignments/removeroleassignment(principalid=${groupId})`;
    return this.executeRequest(requestUrl, "POST").then(() => {
      console.log(`Deleted the current group role from item ID: ${itemId}!`);
      return itemId;
    });
  }

  // Xóa tất cả các quyền hiện có của nhóm khỏi mục
  private removeAllRolesFromItem(itemId: number): Promise<number> {
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/roleassignments`;
    return this.context.spHttpClient
      .get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return Promise.reject("Failed to retrieve role assignments");
        }
        return response.json();
      })
      .then((data) => {
        // Xóa tất cả các vai trò (nếu có)
        const removePromises = data.value.map((roleAssignment: any) =>
          this.removeCurrentRoleFromItem(itemId, roleAssignment.PrincipalId)
        );
        return Promise.all(removePromises).then(() => itemId);
      });
  }

  // Thêm vai trò mới cho nhóm vào mục
  private addNewRoleToItem(
    itemId: number,
    groupId: number,
    roleId: number
  ): Promise<void> {
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;
    return this.executeRequest(requestUrl, "POST").then(() =>
      console.log(`Updated role for item ID: ${itemId} successfully!`)
    );
  }

  // Gọi các hàm để thay đổi quyền
  public manageRoles(
    nameItems: string,
    groupId: number,
    newRoleId: number
  ): void {
    this.getItemId(nameItems)
      .then((itemIds: any) => {
        return Promise.all(
          itemIds.map((itemId: any) =>
            this.breakRoleInheritanceForItem(itemId)
              .then(() => this.removeAllRolesFromItem(itemId))
              .then(() => this.addNewRoleToItem(itemId, groupId, newRoleId))
          )
        );
      })
      .then(() =>
        alert(
          `Updated roles for all items with title '${nameItems}' successfully!`
        )
      )
      .catch((error) =>
        console.error(
          `Error updating roles for items with title '${nameItems}':`,
          error
        )
      );
  }

  //Đoạn code mặc định----------------------------------------------------------------------------------------------------------------------------------------------
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
