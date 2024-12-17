import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export const updateNationColumn = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  subFolderName: string,
  folderName: string,
  customId: string
): Promise<void> => {
  const body = JSON.stringify({
    __metadata: { type: "SP.ListItem" },
    Nation: folderName,
    CustomID: customId,
  });

  return spHttpClient
    .post(
      `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('ProjectFolder/PROJECT/${subFolderName}')/ListItemAllFields`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "odata-version": "",
          "If-Match": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: body,
      }
    )
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        return response.json().then((error) => {
          return Promise.reject(`Error ${response.status}: ${error}`);
        });
      }
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};
