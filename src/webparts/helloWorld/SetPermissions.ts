import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

//Hàm lấy ID của các nhóm
export const getIdGroup = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string
): Promise<void> => {
  return spHttpClient
    .get(`${sharepointUrl}/_api/web/sitegroups`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    })
    .then((data) => {
      const groups: { Title: string; Id: number }[] = data.value;
      groups.forEach((group) => {
        console.log(`Group Name: ${group.Title}, ID: ${group.Id}`);
      });
    })
    .catch((error) => console.error("Error fetching groups:", error));
};

//Set permissions cho item-sharepoint list-------------------------------------------------------------------------------------------------------
//Lấy ID của item dựa vào giá trị ở cột Nation
const getItemId = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  nameItems: string
): Promise<number[]> => {
  const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items?$filter=Nation eq '${nameItems}'&$select=ID`;
  return spHttpClient
    .get(requestUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        return Promise.reject("Failed to retrieve item ID");
      }
      return response.json();
    })
    .then((data) => {
      if (data.value && data.value.length > 0) {
        return data.value.map((item: { ID: number }) => item.ID);
      } else {
        return Promise.reject("Item not found");
      }
    });
};

//Hàm quản lý quyền
export const manageRoles = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  nameItems: string,
  groupId: number,
  newRoleId: number,
  formDigestValue: string
): Promise<void> => {
  return getItemId(spHttpClient, sharepointUrl, nameSharepointList, nameItems)
    .then((itemIds: number[]) => {
      return Promise.all(
        itemIds.map((itemId) =>
          breakRoleInheritanceItem(
            spHttpClient,
            sharepointUrl,
            nameSharepointList,
            itemId,
            formDigestValue
          )
            .then(() =>
              removeAllRolesFromItem(
                spHttpClient,
                sharepointUrl,
                nameSharepointList,
                itemId,
                formDigestValue
              )
            )
            .then(() =>
              addNewRoleItem(
                spHttpClient,
                sharepointUrl,
                nameSharepointList,
                itemId,
                groupId,
                newRoleId,
                formDigestValue
              )
            )
        )
      );
    })
    .then(() => {
      console.log(`Updated roles for all items with branch '${nameItems}'`);
      alert(`Updated roles for all items with branch '${nameItems}'`);
    })
    .catch((error) =>
      console.error(
        `Error updating roles for items with branch '${nameItems}':`,
        error
      )
    );
};

//Hàm gửi yêu cầu tới SharePoint
const executeRequest = (
  spHttpClient: SPHttpClient,
  url: string,
  method: string,
  formDigestValue: string
): Promise<void> => {
  return spHttpClient
    .post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: "application/json;odata=verbose",
        "X-RequestDigest": formDigestValue,
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
};

// Ngắt quyền kế thừa của mục
const breakRoleInheritanceItem = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  itemId: number,
  formDigestValue: string
): Promise<number> => {
  const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/breakroleinheritance(true)`;
  return executeRequest(spHttpClient, requestUrl, "POST", formDigestValue).then(
    () => {
      console.log(`Break role inheritance for item ID: ${itemId}`);
      return itemId;
    }
  );
};

//Xóa quyền của nhóm hiện tại khỏi mục
const removeCurrentRoleFromItem = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  itemId: number,
  groupId: number,
  formDigestValue: string
): Promise<number> => {
  const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/roleassignments/removeroleassignment(principalid=${groupId})`;
  return executeRequest(spHttpClient, requestUrl, "POST", formDigestValue).then(
    () => {
      console.log(`Remove the current group role from item ID: ${itemId}`);
      return itemId;
    }
  );
};

//Xóa tất cả các quyền hiện có của các nhóm khỏi mục
const removeAllRolesFromItem = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  itemId: number,
  formDigestValue: string
): Promise<number> => {
  const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/roleassignments`;
  return spHttpClient
    .get(requestUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        return Promise.reject("Failed to retrieve role assignments");
      }
      return response.json();
    })
    .then((data) => {
      // Xóa tất cả các quyền (nếu có)
      const removePromises = data.value.map((roleAssignment: any) =>
        removeCurrentRoleFromItem(
          spHttpClient,
          sharepointUrl,
          nameSharepointList,
          itemId,
          roleAssignment.PrincipalId,
          formDigestValue
        )
      );
      return Promise.all(removePromises).then(() => itemId);
    });
};

//Thêm quyền mới cho nhóm
const addNewRoleItem = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  itemId: number,
  groupId: number,
  roleId: number,
  formDigestValue: string
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;
  return executeRequest(spHttpClient, requestUrl, "POST", formDigestValue);
};

//Set permissions cho Folders--------------------------------------------------------------------------------------------------------------------
const getFolderServerRelativeUrl = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  libraryName: string,
  folderPath: string
): Promise<string> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryName}/${folderPath}')`;
  return spHttpClient
    .get(requestUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        return Promise.reject("Failed to retrieve folder information");
      }
      return response.json();
    })
    .then((data) => {
      console.log(`Folder URL: ${data.ServerRelativeUrl}`);
      return data.ServerRelativeUrl;
    });
};

// Ngắt quyền kế thừa của mục
const breakRoleInheritanceFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  formDigestValue: string
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/listItemAllFields/breakroleinheritance(true)`;
  return executeRequest(spHttpClient, requestUrl, "POST", formDigestValue).then(
    () => {
      console.log(`Break role inheritance for folder: ${folderUrl}`);
    }
  );
};

//Thêm quyền mới cho nhóm
const addNewRoleFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  groupId: number,
  roleId: number,
  formDigestValue: string
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/listItemAllFields/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;
  return executeRequest(spHttpClient, requestUrl, "POST", formDigestValue);
};

//Hàm quản lý quyền
export const manageRolesFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  libraryName: string,
  folderPath: string,
  groupId: number,
  newRoleId: number,
  formDigestValue: string
): Promise<void> => {
  return getFolderServerRelativeUrl(
    spHttpClient,
    sharepointUrl,
    libraryName,
    folderPath
  )
    .then((folderUrl) =>
      breakRoleInheritanceFolder(
        spHttpClient,
        sharepointUrl,
        folderUrl,
        formDigestValue
      ).then(() =>
        addNewRoleFolder(
          spHttpClient,
          sharepointUrl,
          folderUrl,
          groupId,
          newRoleId,
          formDigestValue
        )
      )
    )
    .then(() => {
      console.log(`Updated roles for folder: ${folderPath}`);
      alert(`Updated roles for folder: ${folderPath}`);
    })
    .catch((error) => {
      console.error(`Error updating roles for folder: ${folderPath}`, error);
    });
};
