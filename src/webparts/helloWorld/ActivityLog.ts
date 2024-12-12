import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

//Lấy tên user name
export const getUserName = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string
): Promise<string> => {
  return spHttpClient
    .get(
      `${sharepointUrl}/_api/web/currentuser`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => response.json())
    .then((data) => data.Title);
};

//Hàm tạo nội dung khi click
export const handleClick = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointSite: string,
  buttonName: string
) => {
  getUserName(spHttpClient, sharepointUrl).then((userName) => {
    const getTimestamp = new Date().toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: true,
    });
    const getMessage = `${getTimestamp}: ${userName} Run ${buttonName}`;
    const folderUrl = `/sites/${nameSharepointSite}/ProjectFolder/ADMIN/Activity log`;
    const fileName = "Activity_log.json";

    return spHttpClient
      .get(
        `${sharepointUrl}/_api/web/GetFileByServerRelativeUrl('${folderUrl}/${fileName}')/$value`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      )
      .then((response) => {
        if (response.ok) {
          return response.text();
        } else if (response.status === 404) {
          return "[]";
        } else {
          return Promise.reject(
            `Error: ${response.status}, ${response.statusText}`
          );
        }
      })
      .then((existingContent) => {
        return Promise.resolve()
          .then(() => JSON.parse(existingContent))
          .catch((error) => {
            console.error("Error:", error);
            return [];
          })
          .then((currentData) => {
            currentData.push(getMessage);
            const updatedJson = JSON.stringify(currentData, null, 1);
            saveJsonSharePoint(
              spHttpClient,
              sharepointUrl,
              folderUrl,
              fileName,
              updatedJson
            );
            setTimeout(() => {
              displayJsonContent(
                spHttpClient,
                sharepointUrl,
                nameSharepointSite
              );
            }, 1000);
          });
      })
      .catch((error) => console.error("Error:", error));
  });
};

//Hiển thị nội dung từ file Json
export const displayJsonContent = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointSite: string
) => {
  const folderUrl = `/sites/${nameSharepointSite}/ProjectFolder/ADMIN/Activity log`;
  const fileName = "Activity_log.json";

  spHttpClient
    .get(
      `${sharepointUrl}/_api/web/GetFileByServerRelativeUrl('${folderUrl}/${fileName}')/$value`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "odata-version": "",
        },
      }
    )
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
    .catch((error) => console.error("Error:", error));
};

//Hàm Save file json vào thư mục mỗi khi click vào 1 nút
const saveJsonSharePoint = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  fileName: string,
  jsonData: string
) => {
  const url = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files/add(url='${fileName}',overwrite=true)`;
  return spHttpClient
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
        response.json().then((error) => console.error("Error:", error));
      }
    });
};

//--------------------------------------------------------------------------------------------------------------------------------------------
// Hàm tạo nội dung khi console.log
export const historyLog = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointSite: string,
  buttonName: string
) => {
  getUserName(spHttpClient, sharepointUrl).then((userName) => {
    const getTimestamp = new Date().toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: true,
    });
    const getMessage = `${getTimestamp}: ${userName} ${buttonName}`;
    const folderUrl = `/sites/${nameSharepointSite}/ProjectFolder/ADMIN/History log`;
    const fileName = "History_log.json";

    return spHttpClient
      .get(
        `${sharepointUrl}/_api/web/GetFileByServerRelativeUrl('${folderUrl}/${fileName}')/$value`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      )
      .then((response) => {
        if (response.ok) {
          return response.text();
        } else if (response.status === 404) {
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
            console.error("Error:", error);
            return [];
          })
          .then((currentData) => {
            currentData.push(getMessage);
            const updatedJson = JSON.stringify(currentData, null, 1);

            return saveJsonSharePoint(
              spHttpClient,
              sharepointUrl,
              folderUrl,
              fileName,
              updatedJson
            );
          });
      })
      .catch((error) => console.error("Error:", error));
  });
};
