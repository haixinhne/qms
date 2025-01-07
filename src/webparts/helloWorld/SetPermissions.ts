import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { DigestCache } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

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

//Hàm gửi yêu cầu tới Sharepoint
const executeRequest = (
  spHttpClient: SPHttpClient,
  url: string,
  contextDigestCache: WebPartContext
): Promise<void> => {
  const digestCache = contextDigestCache.serviceScope.consume(
    DigestCache.serviceKey
  );

  return digestCache
    .fetchDigest(contextDigestCache.pageContext.web.serverRelativeUrl)
    .then((digest: string) => {
      return spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=verbose",
          "X-RequestDigest": digest,
        },
      });
    })
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        return response.json().then((error) => {
          console.error("Error", error);
          return Promise.reject("Request failed: " + error);
        });
      }
      return Promise.resolve();
    });
};

//Set permissions cho item-sharepoint list-------------------------------------------------------------------------------------------------------------------------
//Lấy ID của item dựa vào giá trị ở cột Nation
const getIdItem = (
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

//Ngắt quyền kế thừa của item
const breakRoleInheritanceItem = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  itemId: number,
  contextDigestCache: WebPartContext
): Promise<number> => {
  const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/breakroleinheritance(true)`;
  return executeRequest(spHttpClient, requestUrl, contextDigestCache).then(
    () => {
      return itemId;
    }
  );
};

//Xóa quyền của 1 nhóm khỏi item
const removeCurrentRoleItem = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  itemId: number,
  groupId: number,
  contextDigestCache: WebPartContext
): Promise<number> => {
  const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/roleassignments/removeroleassignment(principalid=${groupId})`;
  return executeRequest(spHttpClient, requestUrl, contextDigestCache).then(
    () => {
      return itemId;
    }
  );
};

//Xóa quyền của tất cả nhóm khỏi item
const removeAllRoleItem = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  itemId: number,
  contextDigestCache: WebPartContext
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
        removeCurrentRoleItem(
          spHttpClient,
          sharepointUrl,
          nameSharepointList,
          itemId,
          roleAssignment.PrincipalId,
          contextDigestCache
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
  contextDigestCache: WebPartContext
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;
  return executeRequest(spHttpClient, requestUrl, contextDigestCache);
};

//Event
export const manageRoles = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  nameItems: string,
  groupId: number,
  newRoleId: number,
  contextDigestCache: WebPartContext
): Promise<void> => {
  return getIdItem(spHttpClient, sharepointUrl, nameSharepointList, nameItems)
    .then((itemIds: number[]) => {
      return Promise.all(
        itemIds.map((itemId) =>
          breakRoleInheritanceItem(
            spHttpClient,
            sharepointUrl,
            nameSharepointList,
            itemId,
            contextDigestCache
          )
            .then(() =>
              removeAllRoleItem(
                spHttpClient,
                sharepointUrl,
                nameSharepointList,
                itemId,
                contextDigestCache
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
                contextDigestCache
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

//Set permissions cho Folder-------------------------------------------------------------------------------------------------------------------------------------
//Ngắt quyền kế thừa của folder
const breakRoleInheritanceFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  contextDigestCache: WebPartContext
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=true, clearSubscopes=true)`;
  return executeRequest(spHttpClient, requestUrl, contextDigestCache);
};

//Xóa quyền của 1 nhóm khỏi folder
const removeCurrentRoleFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  groupId: number,
  contextDigestCache: WebPartContext
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/removeroleassignment(principalid=${groupId})`;

  return executeRequest(spHttpClient, requestUrl, contextDigestCache);
};

//Xóa quyền của tất cả nhóm khỏi folder
const removeAllRoleFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  contextDigestCache: WebPartContext
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
        removeCurrentRoleFolder(
          spHttpClient,
          sharepointUrl,
          folderUrl,
          roleAssignment.PrincipalId,
          contextDigestCache
        )
      );

      return Promise.all(removePromises).then(() => {});
    });
};

//Thêm quyền mới cho nhóm
const addNewRoleFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  groupId: number,
  roleId: number,
  contextDigestCache: WebPartContext
): Promise<void> => {
  const requestUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;

  return executeRequest(spHttpClient, requestUrl, contextDigestCache);
};

//Event
export const manageRolesFolder = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrl: string,
  groupId: number,
  newRoleId: number,
  formDigestValue: any
): Promise<void> => {
  return breakRoleInheritanceFolder(
    spHttpClient,
    sharepointUrl,
    folderUrl,
    formDigestValue
  )
    .then(() =>
      removeAllRoleFolder(
        spHttpClient,
        sharepointUrl,
        folderUrl,
        formDigestValue
      )
    )
    .then(() =>
      addNewRoleFolder(
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
