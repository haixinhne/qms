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

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }">

     <div class="${styles.welcome}">    
        <h2>Hello, ${escape(this.context.pageContext.user.displayName)}</h2>
       <div>${this._environmentMessage}</div>
       <input type="text" class="${
         styles.qms_input
       }" id="titleSharepointList", " placeholder="Title sharepoint list" />
        <button class="${
          styles.qms_button
        }" id="createFolderButton">Create Folder</button>
        <button class="${
          styles.qms_button
        }" id="createSharepointButton">Create Sharepoint</button>       
     </div>
     </div>     
     </section>`;

    //Click button
    const buttonClick = this.domElement.querySelector("#createFolderButton");
    if (buttonClick) {
      buttonClick.addEventListener("click", () => this.onCreateFolder());
    }

    const buttonClickCreateSharepoint = this.domElement.querySelector(
      "#createSharepointButton"
    );
    if (buttonClickCreateSharepoint) {
      buttonClickCreateSharepoint.addEventListener("click", () =>
        this.onClickButtonCreateSharepoint()
      );
    }
    this.renderListAsync();
  }

  protected onInit(): Promise<void> {
    return this.getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
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
  //Hàm lấy file excel
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

  //Hàm đọc nội dung file excel (lấy tên các cột)
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
    console.log("Items", nameItems);
    return { nameColumnSharepoint, nameItems };
  }

  // Tạo SharePoint list
  private createSharePointList(listName: string): Promise<any> {
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
  // Tạo cột Sharepoint list
  private async createColumnInSharePoint(
    listName: string,
    columnName: string
  ): Promise<any> {
    const body = JSON.stringify({
      __metadata: { type: "SP.Field" },
      Title: columnName,
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

  //Hàm caps lock chữ cái đầu
  private capsLocksFirstLetter(text: string): string {
    return text.charAt(0).toUpperCase() + text.slice(1);
  }

  //Tạo items
  private async createItemsInSharePointList(
    listName: string,
    itemData: any
  ): Promise<void> {
    const capsLockListName = this.capsLocksFirstLetter(listName);
    const bodyObject: Record<string, any> = {
      __metadata: {
        type: `SP.Data.${capsLockListName}ListItem`,
      },
      Title: itemData.Title || "QMS",
    };

    for (const key in itemData) {
      if (itemData.hasOwnProperty(key)) {
        bodyObject[key] = String(itemData[key] || "");
      }
    }
    const body = JSON.stringify(bodyObject);
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
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.json().then((errorResponse) => {
            console.error("Error response:", errorResponse);
          });
        }
      })
      .catch((error) => {
        console.error("Error adding item:", error);
      });
  }

  // Click Tạo SharePoint list, add column, add items
  private onClickButtonCreateSharepoint(): void {
    const listNameSharePoint = (
      document.getElementById("titleSharepointList") as HTMLInputElement
    ).value;

    if (!listNameSharePoint) {
      alert("Please enter a name for the SharePoint list!");

      return;
    }
    //Tạo sharepoint list
    this.createSharePointList(listNameSharePoint)
      .then(() => {
        (
          document.getElementById("titleSharepointList") as HTMLInputElement
        ).value = "";
        console.log(`Created successfully: ${listNameSharePoint}`);
      })
      //Tạo cột từ file excel
      .then(() => this.getFileExcelFromSharePoint(excelUrl))
      .then((fileContent: ArrayBuffer) => {
        const { nameColumnSharepoint, nameItems } =
          this.readFileExcelFromSharePoint(fileContent);
        return nameColumnSharepoint
          .reduce((promise, createColumn) => {
            return promise.then(() => {
              console.log(`Adding column: ${createColumn}`);
              return this.createColumnInSharePoint(
                listNameSharePoint,
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
      //Tạo items
      .then(({ nameColumnSharepoint, nameItems }) => {
        console.log("All columns created successfully.");
        console.log("Input data", nameItems);
        return nameItems.reduce((promise, itemData) => {
          const itemObject = nameColumnSharepoint.reduce((obj, columnName) => {
            obj[columnName] = itemData[columnName] || null;
            return obj;
          }, {} as Record<string, any>);
          return promise.then(() =>
            this.createItemsInSharePointList(listNameSharePoint, itemObject)
          );
        }, Promise.resolve());
      })
      .then(() => {
        console.log("All items created successfully.");
      })
      .catch((error) => {
        console.error("Click Created Error:", error);
      });
  }

  //Tạo các thư mục con từ sharepoint------------------------------------------------------------------------------------------------------------------------------
  // Hàm lấy data từ SharePoint list (Lấy tên thư mục là 1 cột ở sharepoint list)
  private getFileFromSharePoint(): Promise<string[]> {
    const listNameSharePoint = (
      document.getElementById("titleSharepointList") as HTMLInputElement
    ).value;
    const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listNameSharePoint}')/items`;

    return this.context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: any) => {
        console.log("Data", data);
        const folderNames = data.value
          .map((item: any) => item.Name) //Cột lấy tên folder
          .filter(Boolean);

        return folderNames;
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
      },
    };
    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('Shared Documents/${parentFolderName}/${subFolderName}')`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .catch((error) => {});
  }

  //Hàm tạo folder
  private createFolder(folderName: string): Promise<any> {
    const optionsHTTP: ISPHttpClientOptions = {
      headers: {
        accept: "application/json; odata=verbose",
        "content-type": "application/json; odata=verbose",
      },
    };
    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('Shared Documents/${folderName}')`,
        SPHttpClient.configurations.v1,
        optionsHTTP
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then(() => {
        return Promise.all([
          this.createSubfolder(folderName, "00_ESSENTIAL"),
          this.createSubfolder(folderName, "01_WORK"),
          this.createSubfolder(folderName, "02_SUBMIT"),
        ]);
      })
      .catch(() => {});
  }

  //Click Tạo folder
  private onCreateFolder(): void {
    this.getFileFromSharePoint()
      .then((folderNames: string[]) => {
        folderNames.forEach((folderName: string) => {
          this.createFolder(folderName)
            .then(() => {
              (
                document.getElementById(
                  "titleSharepointList"
                ) as HTMLInputElement
              ).value = "";
              console.log(`Created successfully: ${folderName}`);
            })
            .catch((error) => {
              console.error(`Error creating:: ${folderName}`, error);
            });
        });
      })
      .catch((error) => {
        console.error("Lỗi:", error);
      });
  }

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
