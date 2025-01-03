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
    .catch((error) => console.error("Error", error));
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
      console.log("The list permissions were set successfully");
    })
    .catch((error) => console.error(`Error ${nameItems}`, error));
};

//Hàm gửi yêu cầu tới SharePoint
const executeRequest = (
  spHttpClient: SPHttpClient,
  url: string,
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
          console.error("Error", errorDetails);
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
  return executeRequest(spHttpClient, requestUrl, formDigestValue).then(() => {
    return itemId;
  });
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
  return executeRequest(spHttpClient, requestUrl, formDigestValue).then(() => {
    return itemId;
  });
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
  return executeRequest(spHttpClient, requestUrl, formDigestValue);
};

//Set permissions cho Folders--------------------------------------------------------------------------------------------------------------------
export const manageRolesFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  groupId: number,
  newRoleId: number,
  formDigestValue: string
): Promise<void> => {
  return breakRoleInheritanceFolder(
    spHttpClient,
    sharepointUrl,
    folderUrl,
    formDigestValue
  )
    .then(() =>
      removeAllRolesFromFolder(
        spHttpClient,
        sharepointUrl,
        folderUrl,
        formDigestValue
      )
    )
    .then(() =>
      addNewRoleToFolder(
        spHttpClient,
        sharepointUrl,
        folderUrl,
        groupId,
        newRoleId,
        formDigestValue
      )
    )
    .then(() => {
      console.log("The folders permissions were set successfully");
    })
    .catch((error) => console.error(`Error ${folderUrl}`, error));
};

const breakRoleInheritanceFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  formDigestValue: string
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=true, clearSubscopes=true)`;
  return executeRequest(spHttpClient, requestUrl, formDigestValue);
};

const removeAllRolesFromFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  formDigestValue: string
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments`;

  return spHttpClient
    .get(requestUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        return Promise.reject("Failed to retrieve role assignments for folder");
      }
      return response.json();
    })
    .then((data) => {
      const removePromises = data.value.map((roleAssignment: any) =>
        removeCurrentRoleFromFolder(
          spHttpClient,
          sharepointUrl,
          folderUrl,
          roleAssignment.PrincipalId,
          formDigestValue
        )
      );

      return Promise.all(removePromises).then(() => {});
    });
};

const removeCurrentRoleFromFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  groupId: number,
  formDigestValue: string
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/removeroleassignment(principalid=${groupId})`;

  return executeRequest(spHttpClient, requestUrl, formDigestValue);
};

const addNewRoleToFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  groupId: number,
  roleId: number,
  formDigestValue: string
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;

  return executeRequest(spHttpClient, requestUrl, formDigestValue);
};
