import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export const exportDataFromSharepointListUseAGGridTable = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<
  {
    CustomID: string;
    ProjectName: string;
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
        .filter(
          (item: any) =>
            item.CustomID &&
            item.ProjectName &&
            item.Phase01Progress &&
            item.Phase01Date &&
            item.Phase01Review &&
            item.Phase02Progress &&
            item.Phase02Date &&
            item.Phase02Review &&
            item.Phase03Progress &&
            item.Phase03Date &&
            item.Phase03Review
        )
        .map((item: any) => ({
          CustomID: item.CustomID,
          ProjectName: item.Title,
        }));
      console.log(data.value);
      return exportData;
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};
