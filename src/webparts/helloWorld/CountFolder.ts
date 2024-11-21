import { SPHttpClient } from "@microsoft/sp-http";

export const countFiles = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  countFolderUrl: string
) => {
  const fileUrl = `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`;

  return spHttpClient
    .get(fileUrl, SPHttpClient.configurations.v1)
    .then((response) => {
      if (!response.ok) {
        console.log(`HTTP error! Status: ${response.status}`);
      }
      return response.json();
    })
    .then((data) => {
      const files = data.value || [];
      const countTotal = files.length;

      const countApproved = files.filter((file: any) => {
        const fileNameWithoutExtension = file.Name.split(".")
          .slice(0, -1)
          .join(".");
        return fileNameWithoutExtension.endsWith("Approved");
      }).length;

      console.log(`Count Total in ${countFolderUrl}: ${countTotal}`);
      console.log(`Count Approved in ${countFolderUrl}: ${countApproved}`);

      return { countTotal, countApproved };
    })
    .catch((error) => {
      console.error("Error fetching files:", error);
    });
};
