import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
export interface IHelloWorldWebPartProps {
  description: string;
}

const sharepointUrl = "https://iscapevn.sharepoint.com/sites/QMS";

export default class SharePointRoleManager extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  public render(): void {
    //Click button
    const getIdGroup = this.domElement.querySelector("#setPermissions");
    if (getIdGroup) {
      getIdGroup.addEventListener("click", () => this.getIDGroup());
    }

    const setPermissions = this.domElement.querySelector("#setPermissions");
    if (setPermissions) {
      setPermissions.addEventListener("click", () => {
        const manageRolesValue = [
          { nameItems: "Vietnam", groupId: 25, newRoleId: 1073741826 },
          { nameItems: "Japan", groupId: 26, newRoleId: 1073741826 },
          { nameItems: "USA", groupId: 30, newRoleId: 1073741826 },
        ];
        manageRolesValue.forEach(({ nameItems, groupId, newRoleId }) => {
          this.manageRoles(nameItems, groupId, newRoleId);
        });
      });
    }
  }

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
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('QMS')/items?$filter=Note eq '${nameItems}'&$select=ID`;
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
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('QMS')/items(${itemId})/breakroleinheritance(true)`;
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
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('QMS')/items(${itemId})/roleassignments/removeroleassignment(principalid=${groupId})`;
    return this.executeRequest(requestUrl, "POST").then(() => {
      console.log(`Deleted the current group role from item ID: ${itemId}!`);
      return itemId;
    });
  }

  // Xóa tất cả các quyền hiện có của nhóm khỏi mục
  private removeAllRolesFromItem(itemId: number): Promise<number> {
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('QMS')/items(${itemId})/roleassignments`;
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
    const requestUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('QMS')/items(${itemId})/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;
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
}
