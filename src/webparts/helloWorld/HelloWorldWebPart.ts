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

const options: ISPHttpClientOptions = {
  headers: {
    accept: "application/json; odata=verbose",
    "content-type": "application/json; odata=verbose",
  },
};

const siteUrl = "https://iscapevn.sharepoint.com/sites/QMS";

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
        <button class="${
          styles.qms_button
        }" id="createFolderButton">Create Folder</button>
        <button class="${
          styles.qms_button
        }" id="createSharepointButton">Create Sharepoint</button>
     </div>
     </div>     
     </section>`;
    //Hành động nhấn nút
    const buttonClick = this.domElement.querySelector("#createFolderButton");
    if (buttonClick) {
      buttonClick.addEventListener("click", () => this._onClickButton());
    }

    const buttonClickCreateSharepoint = this.domElement.querySelector(
      "#createSharepointButton"
    );
    if (buttonClickCreateSharepoint) {
      buttonClickCreateSharepoint.addEventListener("click", () =>
        this._onClickButtonCreateSharepoint()
      );
    }

    this._renderListAsync();
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
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

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      })
      .catch(() => {});
  }

  //Hải add
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

  //Phần tạo sharepoint list từ excel
  //Hàm lấy file excel
  private getFileExcelFromSharePoint(ExcelUrl: string): Promise<ArrayBuffer> {
    return this.context.spHttpClient
      .get(
        `${
          this.context.pageContext.web.absoluteUrl
        }/_api/web/GetFileByServerRelativeUrl('${"/sites/QMS/Shared Documents/Book1.xlsx"}')/$value`,
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

  //Hàm đọc nội dung file excel
  private readFileExcelFromSharePoint(fileContent: ArrayBuffer): string[] {
    const data = new Uint8Array(fileContent);
    const workbook = XLSX.read(data, { type: "array" });

    //lấy sheet đầu tiên
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    //chuyển đổi sheet thành mảng JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    //lấy tên các tên thư mục nằm ở cột đầu tiên
    const folderNameSharepoint = jsonData
      .slice(1)
      .map((row: any) => row[0])
      .filter(Boolean);

    //lấy tên các tên thư mục nằm ở hàng đầu tiên
    //const folderNames = (jsonData[0] as any[]).filter(Boolean);

    return folderNameSharepoint;
  }

  // Hàm tạo SharePoint list
  private createSharePointList(listName: string): Promise<any> {
    const body = JSON.stringify({
      __metadata: { type: "SP.List" },
      Title: listName,
      AllowContentTypes: true,
      BaseTemplate: 100,
      ContentTypesEnabled: true,
      Description: "Danh sách được tạo tự động",
    });

    const optionsJson: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      body: body,
    };

    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`, // API để tạo SharePoint list
        SPHttpClient.configurations.v1,
        optionsJson
      )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          return response.json().then((errorResponse) => {
            console.error("Error response:", errorResponse);
            console.error("In ra file Json:", optionsJson);
          });
        }
      })
      .catch((error) => {
        console.error("Tạo thất bại:", error);
        throw error; // Ném lỗi để có thể xử lý ở nơi khác nếu cần
      });
  }

  //Hành động nhấn nút tạo folder
  private _onClickButtonCreateSharepoint(): void {
    //Tải file Excel từ SharePoint
    this.getFileExcelFromSharePoint(siteUrl)
      .then((fileContent: ArrayBuffer) => {
        //Đọc danh sách tên thư mục từ file Excel
        const readExcel = this.readFileExcelFromSharePoint(fileContent);
        //Tạo SharePoint list cho mỗi tên thư mục
        readExcel.forEach((listSharepoint: string) => {
          this.createSharePointList("Hải xinh nè")
            .then(() => {
              console.log(`Tạo thành công: ${listSharepoint}`);
            })
            .catch((error) => {
              console.error(`Tạo thất bại: ${listSharepoint}`, error);
              console.log(`JSON ${readExcel}`);
            });
        });
      })
      .catch((error) => {
        console.error("Error loading Excel file from SharePoint:", error);
      });
  }

  //Phần tạo các thư mục con từ sharepoint
  // Hàm lấy dữ liệu từ SharePoint list (Lấy tên thư mục cần tạo theo tên của các item (theo hàng ngang))
  private _getFileFromSharePoint(): Promise<string[]> {
    const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('TestExcel')/items`;

    return this.context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: any) => {
        const folderNames = data.value
          .map((item: any) => item.Title) // item.tên cột chứa tên thư mục
          .filter(Boolean);

        return folderNames;
      })
      .catch((error) => {
        console.error("Error fetching SharePoint list data:", error);
        throw error;
      });
  }

  // Hàm tạo subfolder
  private _createSubfolder(
    parentFolderName: string,
    subFolderName: string
  ): Promise<any> {
    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('Shared Documents/${parentFolderName}/${subFolderName}')`,
        SPHttpClient.configurations.v1,
        options
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .catch((error) => {});
  }

  //Hàm tạo folder
  private _createFolder(folderName: string): Promise<any> {
    return this.context.spHttpClient
      .post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/folders/add('Shared Documents/${folderName}')`,
        SPHttpClient.configurations.v1,
        options
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then(() => {
        // Tạo 3 folder con bên trong folder chính
        return Promise.all([
          this._createSubfolder(folderName, "00_ESSENTIAL"),
          this._createSubfolder(folderName, "01_WORK"),
          this._createSubfolder(folderName, "02_SUBMIT"),
        ]);
      })
      .catch(() => {});
  }

  // Hành động nhấn nút tạo folder
  private _onClickButton(): void {
    this._getFileFromSharePoint()
      .then((folderNames: string[]) => {
        folderNames.forEach((folderName: string) => {
          this._createFolder(folderName)
            .then(() => {
              console.log(`Tạo thành công: ${folderName}`);
            })
            .catch((error) => {
              console.error(`Tạo không thành công: ${folderName}`, error);
            });
        });
      })
      .catch((error) => {
        console.error("Lỗi:", error);
      });
  }

  private _getEnvironmentMessage(): Promise<string> {
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
