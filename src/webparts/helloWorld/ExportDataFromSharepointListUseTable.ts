import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export const exportDataFromSharepointListUseTable = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<
  {
    CustomID: string;
    ProjectName: string;
    Nation: string;
    Phase01Date: string;
    Phase01Progress: number;
    Phase02Date: string;
    Phase02Progress: number;
    Phase03Date: string;
    Phase03Progress: number;
  }[]
> => {
  return spHttpClient
    .get(
      `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => {
      return response.json();
    })
    .then((data) => {
      const exportData = data.value
        .filter((item: any) => item.CustomID && item.ProjectName && item.Nation)
        .map((item: any) => ({
          CustomID: item.CustomID,
          ProjectName: item.ProjectName,
          Nation: item.Nation,
          Phase01Date:
            typeof item.Phase01Date === "string" ? item.Phase01Date : "",
          Phase01Progress:
            typeof item.Phase01Progress === "number"
              ? `${item.Phase01Progress * 100}`
              : 0,
          Phase02Date:
            typeof item.Phase02Date === "string" ? item.Phase02Date : "",
          Phase02Progress:
            typeof item.Phase02Progress === "number"
              ? `${item.Phase02Progress * 100}`
              : 0,
          Phase03Date:
            typeof item.Phase03Date === "string" ? item.Phase03Date : "",
          Phase03Progress:
            typeof item.Phase03Progress === "number"
              ? `${item.Phase03Progress * 100}`
              : 0,
        }));

      console.log(exportData);
      return exportData;
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};
