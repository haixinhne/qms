import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export default class SharePointRoleManager {
  private siteUrl: string;
  private listName: string;
  private spHttpClient: SPHttpClient;

  constructor(siteUrl: string, listName: string, spHttpClient: SPHttpClient) {
    this.siteUrl = siteUrl;
    this.listName = listName;
    this.spHttpClient = spHttpClient;
  }

  public manageRoles(groupId: number, newRoleId: number): void {
    this.breakRoleInheritance()
      .then(() => this.removeCurrentRole(groupId))
      .then(() => this.addNewRole(groupId, newRoleId))
      .then(() => console.log("Cập nhật vai trò thành công"))
      .catch((error) => console.error("Lỗi khi cập nhật vai trò: ", error));
  }

  private breakRoleInheritance(): Promise<void> {
    const requestUrl = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/breakroleinheritance(true)`;
    return this.executeRequest(requestUrl, "POST").then(() =>
      console.log("Đã ngắt kế thừa vai trò")
    );
  }

  private removeCurrentRole(groupId: number): Promise<void> {
    const requestUrl = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/roleassignments/removeroleassignment(${groupId})`;
    return this.executeRequest(requestUrl, "POST").then(() =>
      console.log("Đã xóa vai trò hiện tại của nhóm")
    );
  }

  private addNewRole(groupId: number, roleId: number): Promise<void> {
    const requestUrl = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;
    return this.executeRequest(requestUrl, "POST").then(() =>
      console.log("Đã thêm vai trò mới cho nhóm")
    );
  }

  private executeRequest(url: string, method: string): Promise<void> {
    return this.spHttpClient
      .post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=verbose",
          "X-RequestDigest": (
            document.getElementById("__REQUESTDIGEST") as HTMLInputElement
          ).value,
        },
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return Promise.reject("Request failed");
        }
        return Promise.resolve();
      });
  }
}
