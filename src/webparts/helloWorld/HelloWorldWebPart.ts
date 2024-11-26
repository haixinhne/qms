import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from "@microsoft/sp-lodash-subset";
import styles from "./HelloWorldWebPart.module.scss";
import * as strings from "HelloWorldWebPartStrings";
import * as XLSX from "xlsx";
import { getIdGroup, manageRoles } from "./SetPermissions";
import { handleClick, getUserName } from "./ActivityLog";

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

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _environmentMessage: string = "";
  private childSubFolders: { [key: string]: string[] } = {
    Promotion: [
      "Basic project data",
      "Promotion activities",
      "Reports",
      "Project risk analysis",
      "Stakeholder Management",
      "Design Review 01 AS-ME",
      "Client Contract Review (CCR)",
      "Project Approval (EU-Kento)",
      "Estimate Approval (EU-Kessai)",
    ],
    Design: [
      "Drawings",
      "Funtion Checklist AS",
      "Funtion Checklist ME",
      "Designer Approval Request",
      "Design Review 23 AS-ME",
    ],
    Build: [
      "Project outline",
      "PM Policy",
      "HSE risks (Health, Safely & Env.ment)",
      "Funtion Checklist Update",
      "Quality Plan",
      "Schedule",
      "Construction Kickoff",
    ],
  };

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
         }" id="setPermissions">Set Permissions</button>

         <button class="${
           styles.qms_button
         }" id="countFiles">Count Files</button>
         </div>
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
    const clickCountFiles = this.domElement.querySelector("#countFiles");

    if (!actionsContainer) {
      console.error("The actionsContainer element was not found.");
      return;
    }

    //Event click button------------------------------------------------------------------------------------------------------------------------------------------
    //Tạo sharepoint
    if (clickCreateSharepoint) {
      clickCreateSharepoint.addEventListener("click", () => {
        this.onClickButtonCreateSharepoint();
        handleClick(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList,
          "Create Sharepoint"
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
          nameSharepointList,
          "Create Folder"
        );
      });
    }

    //Set Permissions
    if (setPermissions) {
      setPermissions.addEventListener("click", () => {
        getIdGroup(this.context.spHttpClient, sharepointUrl);

        const manageRolesValue = [
          { nameItems: "Viet Nam-VN", groupId: 25, newRoleId: 1073741826 },
          { nameItems: "Japan-JP", groupId: 26, newRoleId: 1073741826 },
          { nameItems: "USA", groupId: 30, newRoleId: 1073741826 },
        ];
        manageRolesValue.forEach(({ nameItems, groupId, newRoleId }) => {
          manageRoles(
            this.context.spHttpClient,
            sharepointUrl,
            nameSharepointList,
            nameItems,
            groupId,
            newRoleId,
            this.context.pageContext.legacyPageContext.formDigestValue
          );
        });

        handleClick(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList,
          "Set Permissions"
        );
      });
    }

    //Count files
    if (clickCountFiles) {
      clickCountFiles.addEventListener("click", () => {
        this.onCountFiles();
        this.onCountFilesDocuments();
        handleClick(
          this.context.spHttpClient,
          sharepointUrl,
          nameSharepointList,
          "Count Files"
        );
      });
    }

    this.renderListAsync();
  }

  //Hàm defaults--------------------------------------------------------------------------------------------------------------------------------------------------
  protected onInit(): Promise<any> {
    return this.getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
      return getUserName(this.context.spHttpClient, sharepointUrl);
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

  //Tạo sharepoint list từ excel, add, update, xóa items từ sharepoint--------------------------------------------------------------------------------------------
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

  //Tạo sharepoint list
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

  //Tạo cột sharepoint list
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
              console.log(`Failed to retrieve items: ${response.statusText}`);
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

  //Tạo các thư mục từ sharepoint------------------------------------------------------------------------------------------------------------------------------
  // Hàm lấy data từ sharepoint list (Lấy tên thư mục là 1 cột ở sharepoint list)
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
          .filter((item: any) => item.Branch && item.ProjectName)
          .map((item: any) => ({
            folderName: item.Branch,
            subFolderName: item.ProjectName,
          }))
          .filter(
            (name: any, index: Number, self: any) =>
              self.findIndex(
                (p: any) =>
                  p.folderName === name.folderName &&
                  p.subFolderName === name.subFolderName
              ) === index
          );

        return folderValues;
      })
      .catch((error) => {
        console.error("Error fetching SharePoint list data:", error);
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
    const subFolderUrl = `Shared Documents/PROJECT/${parentFolderName}/${subFolderName}`;
    const subFolders = ["Promotion", "Design", "Build"];
    const arrayFolderUrl: string[] = [];
    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('Shared Documents/PROJECT/${parentFolderName}/${subFolderName}')`,
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
                // Tạo các thư mục con nhỏ hơn trong từng thư mục con
                const childFolders = this.childSubFolders[folder];
                return childFolders.reduce((childPrevPromise, childFolder) => {
                  const childFolderUrl = `${folderUrl}/${childFolder}`;
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
      .then(() => {
        console.log(
          `Created subfolders ${subFolderName} in: Shared Documents/PROJECT/${parentFolderName}`
        );
        alert(
          `Created subfolders ${subFolderName} in: Shared Documents/PROJECT/${parentFolderName}`
        );
      })
      .catch((error) => {
        console.error("Error creating subfolder:", error);
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
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('Shared Documents/PROJECT/${folderName}')`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => response.json())
      .then(() => {
        console.log(
          `Created folders ${folderName} in: Shared Documents/PROJECT`
        );
        alert(`Created folders ${folderName} in: Shared Documents/PROJECT`);
        return Promise.all(
          subFolderNames.map((subFolderName) =>
            this.createSubfolder(folderName, subFolderName)
          )
        );
      })
      .catch((error) => {
        console.error("Error creating folder or subfolders:", error);
      });
  }

  //Click tạo folder
  private onCreateFolder(): Promise<any> {
    return this.getFileFromSharePoint()
      .then((folderPairs) => {
        const folderMap = folderPairs.reduce((acc, pair) => {
          if (!acc[pair.folderName]) {
            acc[pair.folderName] = [];
          }
          acc[pair.folderName].push(pair.subFolderName);
          return acc;
        }, {} as Record<string, string[]>);

        //Tạo các thư mục kèm các thư mục con tương ứng
        const loopCreateFolder = [];
        for (const folderName in folderMap) {
          if (folderMap.hasOwnProperty(folderName)) {
            const subFolderNames = folderMap[folderName];
            loopCreateFolder.push(
              this.createFolder(folderName, subFolderNames)
            );
          }
        }

        return Promise.all(loopCreateFolder);
      })
      .catch((error) => {
        console.error("Error processing folders and subfolders:", error);
      });
  }

  //Đếm số lượng folder, cập nhật lên sharepoint------------------------------------------------------------------------------------------------------------------------------------------
  //Đếm
  private countFiles(folderUrls: string[]): Promise<{
    totalFiles: number;
    approvedFiles: number;
    percentFiles: number;
  }> {
    const fetchFileCounts = (
      countFolderUrl: string
    ): Promise<{ total: number; approved: number }> => {
      return this.context.spHttpClient
        .get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`,
          SPHttpClient.configurations.v1
        )
        .then((response) => {
          if (!response.ok) {
            console.log(`HTTP error! Status: ${response.status}`);
          }
          return response.json();
        })
        .then((data) => {
          const files = data.value || [];
          const total = files.length;

          const approved = files.filter((file: any) => {
            const fileNameWithoutExtension = file.Name.split(".")
              .slice(0, -1)
              .join(".");
            return fileNameWithoutExtension.endsWith("Approved");
          }).length;

          return { total, approved };
        })
        .catch((error) => {
          console.error(`Error fetching files from ${countFolderUrl}:`, error);
          return { total: 0, approved: 0 };
        });
    };

    const loopFolders = folderUrls.map((url: string) => fetchFileCounts(url)); //Lặp qua các thư mục
    return Promise.all(loopFolders).then((results) => {
      const totalFiles = results.reduce((sum, result) => sum + result.total, 0);
      const approvedFiles = results.reduce(
        (sum, result) => sum + result.approved,
        0
      );
      const percentFiles =
        totalFiles > 0
          ? parseFloat((approvedFiles / totalFiles).toFixed(2))
          : 0;

      return { totalFiles, approvedFiles, percentFiles };
    });
  }

  //Lấy Url các thư mục
  private getUrlCountFiles(
    parentFolderName: string,
    subFolderName: string | string[]
  ): Promise<any> {
    if (typeof subFolderName === "string") {
      subFolderName = [subFolderName];
    }

    const subFolderUrl = `Shared Documents/PROJECT/${parentFolderName}/${subFolderName}`;
    const subFolders = ["Promotion", "Design", "Build"];
    const arrayFolderUrl: string[] = [];

    subFolders.forEach((folder) => {
      const folderUrl = `${subFolderUrl}/${folder}`;
      arrayFolderUrl.push(folderUrl);

      const childFolders = this.childSubFolders[folder];
      childFolders.forEach((childFolder) => {
        const childFolderUrl = `${folderUrl}/${childFolder}`;
        arrayFolderUrl.push(childFolderUrl);
      });
    });

    return this.countFiles(arrayFolderUrl)
      .then(({ totalFiles, approvedFiles, percentFiles }) => {
        console.log(`Total Files in ${subFolderName}: ${totalFiles}`);
        console.log(`Approved Files in ${subFolderName}: ${approvedFiles}`);
        console.log(`Completion rate in ${subFolderName}: ${percentFiles}`);
        return { totalFiles, approvedFiles, percentFiles };
      })
      .catch((error) => {
        console.error("Error counting files:", error);
      });
  }

  //Update Rate cho từng dự án ứng với ProjectName
  private updateRateSharepoint = (
    subFolderName: string,
    percentFiles: number
  ): Promise<any> => {
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items?$filter=ProjectName eq '${subFolderName}'&$select=ID,ProjectName`;

    return this.context.spHttpClient
      .get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return Promise.reject(`Failed to retrieve item for ${subFolderName}`);
        }
        return response.json();
      })
      .then((data) => {
        if (data.value && data.value.length > 0) {
          const item = data.value[0];
          const itemId = item.ID;
          const rateValue = percentFiles.toString();

          const body = JSON.stringify({
            __metadata: { type: `SP.Data.${nameSharepointList}ListItem` },
            Rate: rateValue,
          });

          const optionsHTTP: ISPHttpClientOptions = {
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "odata-version": "",
              "If-Match": "*",
              "X-HTTP-Method": "MERGE",
            },
            body: body,
          };

          return this.context.spHttpClient
            .post(
              `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})`,
              SPHttpClient.configurations.v1,
              optionsHTTP
            )
            .then((response) => {
              if (!response.ok) {
                return response.text().then((text) => {
                  Promise.reject(
                    `Failed to update Rate for item ${itemId}: ${response.statusText}`
                  );
                });
              }
            });
        }
        return Promise.reject(
          `No item found for ProjectName: ${subFolderName}`
        );
      })
      .catch((error) => {
        console.error("Error updating Rate value:", error);
        return Promise.reject(error);
      });
  };

  //Click đếm file
  private onCountFiles(): Promise<any> {
    return this.getFileFromSharePoint()
      .then((folderPairs) => {
        const folderMap = folderPairs.reduce(
          (acc, { folderName, subFolderName }) => {
            if (!acc[folderName]) {
              acc[folderName] = [];
            }
            acc[folderName].push(subFolderName);
            return acc;
          },
          {} as Record<string, string[]>
        );

        const updatePromises: Promise<any>[] = [];

        //Lặp qua thư mục và các thư mục con
        for (const folderName in folderMap) {
          if (folderMap.hasOwnProperty(folderName)) {
            const subFolderNames = folderMap[folderName];
            subFolderNames.forEach((subFolderName) => {
              updatePromises.push(
                this.getUrlCountFiles(folderName, subFolderName).then(
                  ({ percentFiles }) => {
                    return this.updateRateSharepoint(
                      subFolderName,
                      percentFiles
                    );
                  }
                )
              );
            });
          }
        }

        return Promise.all(updatePromises);
      })
      .catch((error) => {
        console.error("Error processing folders and subfolders:", error);
      });
  }

  //Đếm folder, update lên Document--------------------------------------------------------------------------------------------------------------------------------
  private countFilesDocument(folderUrls: string[]): Promise<{
    totalFiles: number;
    approvedFiles: number;
    percentFiles: string;
  }> {
    const fetchFileCounts = (
      countFolderUrl: string
    ): Promise<{ total: number; approved: number }> => {
      return this.context.spHttpClient
        .get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`,
          SPHttpClient.configurations.v1
        )
        .then((response) => {
          if (!response.ok) {
            console.log(`HTTP error! Status: ${response.status}`);
          }
          return response.json();
        })
        .then((data) => {
          const files = data.value || [];
          const total = files.length;

          const approved = files.filter((file: any) => {
            const fileNameWithoutExtension = file.Name.split(".")
              .slice(0, -1)
              .join(".");
            return fileNameWithoutExtension.endsWith("Approved");
          }).length;

          return { total, approved };
        })
        .catch((error) => {
          console.error(`Error fetching files from ${countFolderUrl}:`, error);
          return { total: 0, approved: 0 };
        });
    };

    const loopFolders = folderUrls.map((url: string) => fetchFileCounts(url)); //Lặp qua các thư mục
    return Promise.all(loopFolders).then((results) => {
      const totalFiles = results.reduce((sum, result) => sum + result.total, 0);
      const approvedFiles = results.reduce(
        (sum, result) => sum + result.approved,
        0
      );
      const percentFiles =
        totalFiles > 0 ? `${approvedFiles} / ${totalFiles}` : 0;

      return { totalFiles, approvedFiles, percentFiles };
    });
  }

  private getUrlCountFilesDocuments(
    parentFolderName: string,
    subFolderName: string | string[]
  ): Promise<any> {
    if (typeof subFolderName === "string") {
      subFolderName = [subFolderName];
    }

    const subFolderUrl = `Shared Documents/PROJECT/${parentFolderName}/${subFolderName}`;
    const subFolders = ["Promotion", "Design", "Build"];
    const arrayFolderUrl: string[] = [];

    subFolders.forEach((folder) => {
      const folderUrl = `${subFolderUrl}/${folder}`;
      arrayFolderUrl.push(folderUrl);

      const childFolders = this.childSubFolders[folder];
      childFolders.forEach((childFolder) => {
        const childFolderUrl = `${folderUrl}/${childFolder}`;
        arrayFolderUrl.push(childFolderUrl);
      });
    });

    return this.countFilesDocument(arrayFolderUrl)
      .then(({ totalFiles, approvedFiles, percentFiles }) => {
        console.log(`Total Files in ${subFolderName}: ${totalFiles}`);
        console.log(`Approved Files in ${subFolderName}: ${approvedFiles}`);
        console.log(`Completion rate in ${subFolderName}: ${percentFiles}`);
        return { totalFiles, approvedFiles, percentFiles };
      })
      .catch((error) => {
        console.error("Error counting files:", error);
      });
  }

  //Click để cập nhật giá trị Approved
  private onCountFilesDocuments(): Promise<any> {
    return this.getFileFromSharePoint()
      .then((folderPairs) => {
        const folderMap = folderPairs.reduce(
          (acc, { folderName, subFolderName }) => {
            if (!acc[folderName]) {
              acc[folderName] = [];
            }
            acc[folderName].push(subFolderName);
            return acc;
          },
          {} as Record<string, string[]>
        );

        const updatePromises: Promise<any>[] = [];

        //Lặp qua thư mục và các thư mục con
        for (const folderName in folderMap) {
          if (folderMap.hasOwnProperty(folderName)) {
            const subFolderNames = folderMap[folderName];
            subFolderNames.forEach((subFolderName) => {
              updatePromises.push(
                this.getUrlCountFilesDocuments(folderName, subFolderName).then(
                  ({ percentFiles }) => {
                    return this.updateFolderApprovedDocuments(percentFiles);
                  }
                )
              );
            });
          }
        }

        return Promise.all(updatePromises);
      })
      .catch((error) => {
        console.error("Error processing folders and subfolders:", error);
      });
  }

  //Update Rate cho thư mục
  private updateFolderApprovedDocuments(approvedValue: string): Promise<any> {
    const folderUrl =
      "Shared Documents/PROJECT/Viet Nam-VN/VN-QMS/Promotion/Client Contract Review (CCR)";
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields`;

    return this.context.spHttpClient
      .get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((text) => {
            console.error(`Error retrieving folder metadata: ${text}`);
            return Promise.reject(
              `Folder doesn't exist or no metadata found for folder: ${text}`
            );
          });
        }
        return response.json();
      })
      .then((data) => {
        if (!data || !data.Id) {
          const body = JSON.stringify({
            __metadata: { type: "SP.Data.DocumentsItem" },
            Approved: approvedValue,
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
              `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Documents')/items`,
              SPHttpClient.configurations.v1,
              optionsHTTP
            )
            .then((createResponse) => {
              if (!createResponse.ok) {
                return createResponse.text().then((text) => {
                  console.error(`Error creating folder item: ${text}`);
                  return Promise.reject(
                    `Failed to create item for folder: Response: ${text}`
                  );
                });
              }
            });
        } else {
          const listItemId = data.Id;

          const body = JSON.stringify({
            __metadata: { type: "SP.ListItem" },
            Approved: approvedValue,
          });

          const optionsHTTP: ISPHttpClientOptions = {
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "odata-version": "",
              "If-Match": "*",
              "X-HTTP-Method": "MERGE",
            },
            body: body,
          };

          return this.context.spHttpClient
            .post(
              `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Documents')/items(${listItemId})`,
              SPHttpClient.configurations.v1,
              optionsHTTP
            )
            .then((updateResponse) => {
              if (!updateResponse.ok) {
                return updateResponse.text().then((text) => {
                  console.error(`Error updating Approved column: ${text}`);
                  return Promise.reject(
                    `Failed to update Approved column: ${updateResponse.statusText}. Response: ${text}`
                  );
                });
              }
            });
        }
      })
      .catch((error) => {
        console.error("Error updating Approved column:", error);
        return Promise.reject(error);
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
