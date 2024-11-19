import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

//Lấy tên user name
const getUserName = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string
): Promise<string> => {
  return spHttpClient
    .get(
      `${sharepointUrl}/_api/web/currentuser`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => response.json())
    .then((data: any) => data.Title);
};

//Hàm tạo nội dung khi click
export const handleClick = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  listname: string,
  buttonName: string
) => {
  getUserName(spHttpClient, sharepointUrl).then((userName) => {
    const getTimestamp = new Date().toLocaleString();
    const getMessage = `${getTimestamp}: ${userName} clicked the ${buttonName} button`;
    const folderUrl = `/sites/${listname}/Shared Documents/ActivityHistory`;
    const fileName = "activityLog.json";
    const fileUrl = `${spHttpClient}/_api/web/GetFileByServerRelativeUrl('${folderUrl}/${fileName}')/$value`;

    spHttpClient
      .get(fileUrl, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "odata-version": "",
        },
      })
      .then((response: any) => {
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
      .then((existingContent: any) => {
        return Promise.resolve()
          .then(() => JSON.parse(existingContent))
          .catch((error) => {
            console.error("Error parsing existing JSON content:", error);
            return [];
          })
          .then((currentData) => {
            currentData.push(getMessage);
            const updatedJson = JSON.stringify(currentData, null, 1);
            saveJsonSharePoint(context, listName, fileName, updatedJson);
            displayJsonContent(context, nameSharepointList);
          });
      })
      .catch((error: any) =>
        console.error("Error processing JSON file:", error)
      );
  });
};

//Hàm Save file json vào thư mục mỗi khi click vào 1 nút
const saveJsonSharePoint = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  listName: string,
  fileName: string,
  jsonData: string
) => {
  const url = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${listName}')/Files/add(url='${fileName}',overwrite=true)`;
  spHttpClient
    .post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "",
      },
      body: jsonData,
    })
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        response
          .json()
          .then((error) => console.error("Error saving file:", error));
      }
    });
};

//Hiển thị nội dung từ file Json
const displayJsonContent = (spHttpClient: SPHttpClient, listName: string) => {
  const folderUrl = `/sites/${listName}/Shared Documents/ActivityHistory`;
  const fileName = "activityLog.json";
  const fileUrl = `${spHttpClient}/_api/web/GetFileByServerRelativeUrl('${folderUrl}/${fileName}')/$value`;

  spHttpClient
    .get(fileUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "",
      },
    })
    .then((response: any) => {
      if (response.ok) {
        return response.text();
      } else {
        return Promise.reject(
          `Error retrieving file. Status: ${response.status}, ${response.statusText}`
        );
      }
    })
    .then((jsonContent: any) => JSON.parse(jsonContent))
    .then((parsedContent: any) => {
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
    .catch((error: any) => console.error("Error processing JSON file:", error));
};
