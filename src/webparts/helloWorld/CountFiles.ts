import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";

//Subfolders
export const childSubFolders: {
  [key: string]: { name: string; DocumentId: string }[];
} = {
  Promotion: [
    { name: "Basic project data", DocumentId: "0100" },
    { name: "Promotion activities", DocumentId: "0101" },
    // { name: "Project risk analysis", DocumentId: "0105" },
    // { name: "Stakeholder Management", DocumentId: "0107" },
    // { name: "Design Review 01 AS-ME", DocumentId: "0109" },
    // { name: "Client Contract Review (CCR)", DocumentId: "0111" },
    // { name: "Project Approval (EU-Kento)", DocumentId: "0113" },
    // { name: "Estimate Approval (EU-Kessai)", DocumentId: "0113" },
  ],
  Design: [
    { name: "Drawings", DocumentId: "2300" },
    { name: "Funtion Checklist AS", DocumentId: "2301" },
    // { name: "Funtion Checklist ME", DocumentId: "2303" },
    // { name: "Designer Approval Request", DocumentId: "2305" },
    // { name: "Design Review 23 AS-ME", DocumentId: "2307" },
  ],
  Build: [
    { name: "Project outline", DocumentId: "4500" },
    { name: "PM Policy", DocumentId: "4501" },
    // { name: "HSE risks (Health, Safely & Env.ment)", DocumentId: "4503" },
    // { name: "Funtion Checklist Update", DocumentId: "4504" },
    // { name: "Quality Plan", DocumentId: "4505" },
    // { name: "Schedule", DocumentId: "4506" },
    // { name: "Construction Kickoff", DocumentId: "4507" },
  ],
};

//Sharepoint-Update progress, progress phase và phase lên Sharepoint list-------------------------------------------------------------------------------------------------
//Update progress----------------------------------------------------------
//Hàm đọc sharepoint list
const getDataFromSharepointList = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<{ subFolderName: string; customId: string }[]> => {
  const listUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items`;

  return spHttpClient
    .get(listUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    })
    .then((data) => {
      const folderValues = data.value

        .filter((item: any) => item.Nation && item.ProjectName)
        .map((item: any) => ({
          subFolderName: item.ProjectName, //ProjectName
          customId: item.CustomID, //CustomId
        }))
        .filter(
          (name: any, index: Number, self: any) =>
            self.findIndex(
              (i: any) => i.subFolderName === name.subFolderName
            ) === index
        );
      return folderValues;
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};

//Hàm đếm file
// const progressFiles = (
//   spHttpClient: SPHttpClient,
//   sharepointUrl: string,
//   folderUrls: string[]
// ): Promise<{
//   totalFiles: number;
//   approvedFiles: number;
//   percentFiles: number;
// }> => {
//   const fetchFileCounts = (
//     countFolderUrl: string
//   ): Promise<{ total: number; approved: number }> => {
//     return spHttpClient
//       .get(
//         `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`,
//         SPHttpClient.configurations.v1
//       )
//       .then((response) => {
//         if (!response.ok) {
//           console.log(`Error: ${response.status}`);
//         }
//         return response.json();
//       })
//       .then((data) => {
//         const files = data.value || [];
//         const total = files.length;

//         const approved = files.filter((file: any) => {
//           const fileNameWithoutExtension = file.Name.split(".")
//             .slice(0, -1)
//             .join(".");
//           return fileNameWithoutExtension.endsWith("Approved");
//         }).length;

//         return { total, approved };
//       })
//       .catch((error) => {
//         console.error(`Error: ${countFolderUrl}`, error);
//         return { total: 0, approved: 0 };
//       });
//   };

//   const loopFolders = folderUrls.map((url: string) => fetchFileCounts(url)); //Lặp qua các thư mục
//   return Promise.all(loopFolders).then((results) => {
//     const totalFiles = results.reduce((sum, result) => sum + result.total, 0);
//     const approvedFiles = results.reduce(
//       (sum, result) => sum + result.approved,
//       0
//     );
//     const percentFiles =
//       totalFiles > 0 ? parseFloat((approvedFiles / totalFiles).toFixed(4)) : 0;
//     console.log(percentFiles);

//     return { totalFiles, approvedFiles, percentFiles };
//   });
// };

//op2
const progressFiles = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrls: string[]
): Promise<{
  totalFiles: number;
  approvedFiles: number;
  percentFiles: number;
}> => {
  const fetchFileCounts = (
    countFolderUrl: string
  ): Promise<{ total: number; approved: number }> => {
    return spHttpClient
      .get(
        `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files?$expand=ListItemAllFields`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        if (!response.ok) {
          console.log(`Error: ${response.status}`);
        }
        return response.json();
      })
      .then((data) => {
        const files = data.value || [];
        const total = files.length;

        const approved = files.filter(
          (file: any) => file.ListItemAllFields?.Status === "Approved"
        ).length;

        return { total, approved };
      })
      .catch((error) => {
        console.error(`Error: ${countFolderUrl}`, error);
        return { total: 0, approved: 0 };
      });
  };

  const loopFolders = folderUrls.map((url: string) => fetchFileCounts(url));
  return Promise.all(loopFolders).then((results) => {
    const totalFiles = results.reduce((sum, result) => sum + result.total, 0);
    const approvedFiles = results.reduce(
      (sum, result) => sum + result.approved,
      0
    );
    const percentFiles =
      totalFiles > 0 ? parseFloat((approvedFiles / totalFiles).toFixed(4)) : 0;

    return { totalFiles, approvedFiles, percentFiles };
  });
};

//Hàm lấy url các thư mục
const getUrlProgressFiles = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  subFolderName: string | string[]
): Promise<any> => {
  if (typeof subFolderName === "string") {
    subFolderName = [subFolderName];
  }

  const subFolderUrl = `ProjectFolder/PROJECT/${subFolderName}`;
  const subFolders = Object.keys(childSubFolders);
  const arrayFolderUrl: string[] = [];

  subFolders.forEach((folder) => {
    const folderUrl = `${subFolderUrl}/${folder}`;
    arrayFolderUrl.push(folderUrl);
    const childFolders = childSubFolders[folder];
    childFolders.forEach((childFolder) => {
      const childFolderName = childFolder.name;
      const childFolderUrl = `${folderUrl}/${childFolderName}`;
      arrayFolderUrl.push(childFolderUrl);
    });
  });

  return progressFiles(spHttpClient, sharepointUrl, arrayFolderUrl)
    .then(({ totalFiles, approvedFiles, percentFiles }) => {
      console.log(`Approved files in  ${subFolderName}  :  ${approvedFiles}`);
      console.log(`Total files in     ${subFolderName}  :  ${totalFiles}`);
      console.log(`Progress in        ${subFolderName}  :  ${percentFiles}`);
      return { totalFiles, approvedFiles, percentFiles };
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};

//Hàm update Progress lên sharepoint list
const updateProgressSharepointList = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  subFolderName: string,
  percentFiles: number
): Promise<any> => {
  return spHttpClient
    .get(
      `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items?$filter=ProjectName eq '${subFolderName}'&$select=ID,ProjectName`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        return Promise.reject(`Failed to retrieve item for ${subFolderName}`);
      }
      return response.json();
    })
    .then((data) => {
      if (data.value && data.value.length > 0) {
        const item = data.value[0];
        const itemId = item.ID;
        const rateValue = percentFiles;

        const body = JSON.stringify({
          __metadata: {
            type: `SP.Data.${nameSharepointList}ListItem`,
          },
          Progress: rateValue,
        });

        const optionsHTTP: ISPHttpClientOptions = {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "",
            "If-Match": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: body,
        };

        return spHttpClient
          .post(
            `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})`,
            SPHttpClient.configurations.v1,
            optionsHTTP
          )
          .then((response) => {
            if (!response.ok) {
              return Promise.reject(`Error ${subFolderName}`);
            }
          });
      } else {
        return Promise.reject(
          `No item found for ${nameSharepointList}: ${subFolderName}`
        );
      }
    })
    .catch((error) => {
      return Promise.reject(error);
    });
};

//Event đếm, update file
export const onProgressSharepointList = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<any> => {
  return getDataFromSharepointList(
    spHttpClient,
    sharepointUrl,
    nameSharepointList
  )
    .then((folderPairs) => {
      const subFolder = folderPairs.map(({ subFolderName }) => subFolderName);
      const updatePromises: Promise<any>[] = [];

      subFolder.forEach((subFolderName) => {
        updatePromises.push(
          getUrlProgressFiles(spHttpClient, sharepointUrl, subFolderName).then(
            ({ percentFiles }) => {
              return updateProgressSharepointList(
                spHttpClient,
                sharepointUrl,
                nameSharepointList,
                subFolderName,
                percentFiles
              );
            }
          )
        );
      });

      return Promise.all(updatePromises);
    })
    .then(() => {})
    .catch((error) => {
      console.error("Error:", error);
    });
};

//Update phase progress Promotion, Design và Build-------------------------------------------
//Hàm đếm file
// const progressFilesPhase = (
//   spHttpClient: SPHttpClient,
//   sharepointUrl: string,
//   folderUrls: string[]
// ): Promise<{
//   totalFiles: number;
//   approvedFiles: number;
//   percentFiles: number;
// }> => {
//   const fetchFileCounts = (
//     folderUrl: string
//   ): Promise<{ total: number; approved: number }> => {
//     return spHttpClient
//       .get(
//         `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files`,
//         SPHttpClient.configurations.v1
//       )
//       .then((response) => {
//         if (!response.ok) {
//           console.log(`Error: ${response.status}`);
//         }
//         return response.json();
//       })
//       .then((data) => {
//         const files = data.value || [];
//         const total = files.length;
//         const approved = files.filter((file: any) => {
//           const fileNameWithoutExtension = file.Name.split(".")
//             .slice(0, -1)
//             .join(".");
//           return fileNameWithoutExtension.endsWith("Approved");
//         }).length;

//         return { total, approved };
//       })
//       .catch((error) => {
//         console.error(`Error: ${folderUrl}`, error);
//         return { total: 0, approved: 0 };
//       });
//   };

//   const folderPromises = folderUrls.map((url) => fetchFileCounts(url));
//   return Promise.all(folderPromises).then((results) => {
//     const totalFiles = results.reduce((sum, result) => sum + result.total, 0);
//     const approvedFiles = results.reduce(
//       (sum, result) => sum + result.approved,
//       0
//     );
//     const percentFiles =
//       totalFiles > 0 ? parseFloat((approvedFiles / totalFiles).toFixed(4)) : 0;
//     return { totalFiles, approvedFiles, percentFiles };
//   });
// };

//op2
const progressFilesPhase = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrls: string[]
): Promise<{
  totalFiles: number;
  approvedFiles: number;
  percentFiles: number;
}> => {
  const fetchFileCounts = (
    folderUrl: string
  ): Promise<{ total: number; approved: number }> => {
    return spHttpClient
      .get(
        `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files?$expand=ListItemAllFields`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        if (!response.ok) {
          console.log(`Error: ${response.status}`);
        }
        return response.json();
      })
      .then((data) => {
        const files = data.value || [];
        const total = files.length;

        const approved = files.filter(
          (file: any) => file.ListItemAllFields?.Status === "Approved"
        ).length;

        return { total, approved };
      })
      .catch((error) => {
        console.error(`Error: ${folderUrl}`, error);
        return { total: 0, approved: 0 };
      });
  };

  const folderPromises = folderUrls.map((url) => fetchFileCounts(url));
  return Promise.all(folderPromises).then((results) => {
    const totalFiles = results.reduce((sum, result) => sum + result.total, 0);
    const approvedFiles = results.reduce(
      (sum, result) => sum + result.approved,
      0
    );
    const percentFiles =
      totalFiles > 0 ? parseFloat((approvedFiles / totalFiles).toFixed(4)) : 0;
    return { totalFiles, approvedFiles, percentFiles };
  });
};

//Hàm lấy url các thư mục
const getUrlProgressPhaseFiles = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  subFolderName: string
): Promise<{
  Promotion: {
    totalFiles: number;
    approvedFiles: number;
    percentFiles: number;
  };
  Design: {
    totalFiles: number;
    approvedFiles: number;
    percentFiles: number;
  };
  Build: {
    totalFiles: number;
    approvedFiles: number;
    percentFiles: number;
  };
}> => {
  const subFolders = Object.keys(childSubFolders);
  const folderPromises = subFolders.map((parentFolder) => {
    const childFolders = childSubFolders[parentFolder];
    const folderUrls = childFolders.map(
      (child) =>
        `ProjectFolder/PROJECT/${subFolderName}/${parentFolder}/${child.name}`
    );

    return progressFilesPhase(spHttpClient, sharepointUrl, folderUrls).then(
      ({ totalFiles, approvedFiles }) => ({
        parentFolder,
        totalFiles,
        approvedFiles,
      })
    );
  });

  return Promise.all(folderPromises).then((results) => {
    const progressMap: {
      Promotion: {
        totalFiles: number;
        approvedFiles: number;
        percentFiles: number;
      };
      Design: {
        totalFiles: number;
        approvedFiles: number;
        percentFiles: number;
      };
      Build: {
        totalFiles: number;
        approvedFiles: number;
        percentFiles: number;
      };
    } = {
      Promotion: { totalFiles: 0, approvedFiles: 0, percentFiles: 0 },
      Design: { totalFiles: 0, approvedFiles: 0, percentFiles: 0 },
      Build: { totalFiles: 0, approvedFiles: 0, percentFiles: 0 },
    };

    results.forEach(({ parentFolder, totalFiles, approvedFiles }) => {
      //Số file và file đã Approved cho mỗi thư mục cha
      progressMap[parentFolder as keyof typeof progressMap].totalFiles +=
        totalFiles;
      progressMap[parentFolder as keyof typeof progressMap].approvedFiles +=
        approvedFiles;
    });

    //Tỷ lệ Approved cho mỗi thư mục cha
    Object.keys(progressMap).forEach((key) => {
      const folderKey = key as keyof typeof progressMap;
      const total = progressMap[folderKey].totalFiles;
      const approved = progressMap[folderKey].approvedFiles;
      progressMap[folderKey].percentFiles =
        total > 0 ? parseFloat((approved / total).toFixed(4)) : 0;
    });
    console.log(`${subFolderName}:${JSON.stringify(progressMap, null, 2)}`);
    return progressMap;
  });
};

//Hàm update Progress Promotion, Design và Build lên sharepoint list Phase01, Phase02, Phase03
const updateProgressPhaseSharepointList = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  subFolderName: string,
  progressData: {
    Promotion: { percentFiles: number };
    Design: { percentFiles: number };
    Build: { percentFiles: number };
  }
): Promise<any> => {
  return spHttpClient
    .get(
      `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items?$filter=ProjectName eq '${subFolderName}'&$select=ID,ProjectName`,
      SPHttpClient.configurations.v1
    )
    .then((response) => response.json())
    .then((data) => {
      if (data.value && data.value.length > 0) {
        const item = data.value[0];
        const itemId = item.ID;

        const body = JSON.stringify({
          __metadata: {
            type: `SP.Data.${nameSharepointList}ListItem`,
          },
          Phase01Progress: progressData.Promotion.percentFiles,
          Phase02Progress: progressData.Design.percentFiles,
          Phase03Progress: progressData.Build.percentFiles,
        });

        const optionsHTTP: ISPHttpClientOptions = {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "",
            "If-Match": "*",
            "X-HTTP-Method": "MERGE",
          },
          body,
        };
        return spHttpClient
          .post(
            `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})`,
            SPHttpClient.configurations.v1,
            optionsHTTP
          )
          .then((response) => {
            if (!response.ok) {
              return Promise.reject(`Error updating ${subFolderName}`);
            }
          });
      } else {
        return Promise.reject(`No item found for ${subFolderName}`);
      }
    });
};

//Event đếm, update file
export const onProgressPhaseSharepointList = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<any> => {
  return getDataFromSharepointList(
    spHttpClient,
    sharepointUrl,
    nameSharepointList
  )
    .then((folderPairs) => {
      const subFolder = folderPairs.map(({ subFolderName }) => subFolderName);
      const updatePromises: Promise<any>[] = [];

      subFolder.forEach((subFolderName) => {
        updatePromises.push(
          getUrlProgressPhaseFiles(
            spHttpClient,
            sharepointUrl,
            subFolderName
          ).then((progress) => {
            return updateProgressPhaseSharepointList(
              spHttpClient,
              sharepointUrl,
              nameSharepointList,
              subFolderName,
              progress
            );
          })
        );
      });
      return Promise.all(updatePromises);
    })
    .then(() => {})
    .catch((error) => {
      console.error("Error:", error);
    });
};

//Update Phase Sharepoint list-------------------------------------------
//Hàm update
const updateSubPhaseProgressSharepointList = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string,
  customId: string,
  documentId: string,
  percentFiles: number
): Promise<void> => {
  return spHttpClient
    .get(
      `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items?$filter=CustomID eq '${customId}'&$select=ID,CustomID`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        return Promise.reject("Failed to retrieve item for CustomID");
      }
      return response.json();
    })
    .then((data) => {
      if (data.value && data.value.length > 0) {
        const item = data.value[0];
        if (!item.ID) {
          return Promise.reject("Item does not have a valid ID");
        }
        const itemId = item.ID;

        const body = JSON.stringify({
          __metadata: {
            type: `SP.Data.${nameSharepointList}ListItem`,
          },
          [`Phase${documentId}`]: percentFiles,
        });

        const optionsHTTP: ISPHttpClientOptions = {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "",
            "If-Match": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: body,
        };

        return spHttpClient
          .post(
            `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items(${itemId})`,
            SPHttpClient.configurations.v1,
            optionsHTTP
          )
          .then((response) => {
            if (!response.ok) {
              return Promise.reject("Error");
            }
          });
      } else {
        return Promise.reject("No item found for CustomID");
      }
    })
    .catch((error) => {
      console.error("Error:", error);
      return Promise.reject(error);
    });
};

//Hàm lấy url
const getUrlSubPhaseProgress = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  subFolderName: string | string[],
  customId: string
): Promise<
  {
    subFolderNames: string[];
    childFolderDocumentId: string;
    percentFiles: number;
  }[]
> => {
  const subFolderNames = Array.isArray(subFolderName)
    ? subFolderName
    : [subFolderName];
  const subFolders = Object.keys(childSubFolders);
  const updatePromises: Promise<{
    subFolderNames: string[];
    childFolderDocumentId: string;
    percentFiles: number;
  }>[] = [];

  subFolders.forEach((folder) => {
    const baseFolderUrl = `ProjectFolder/PROJECT/${subFolderNames}/${folder}`;
    const childFolders = childSubFolders[folder];

    childFolders.forEach((childFolder) => {
      const childFolderName = childFolder.name;
      const childFolderDocumentId = childFolder.DocumentId;
      const childFolderUrl = `${baseFolderUrl}/${childFolderName}`;
      const documentId = childFolder.DocumentId;

      const countAndUpdate = progressFileFoldersOption2(
        spHttpClient,
        sharepointUrl,
        [childFolderUrl]
      )
        .then(({ percentFiles }) => {
          return updateProgressFileFoldersOption2(
            spHttpClient,
            sharepointUrl,
            percentFiles,
            childFolderUrl,
            documentId
          ).then(() => ({
            subFolderNames,
            childFolderDocumentId,
            percentFiles,
          }));
        })
        .catch((error) => {
          console.error(`Error ${childFolderUrl}:`, error);

          return {
            subFolderNames,
            childFolderDocumentId,
            percentFiles: 0,
          };
        });

      updatePromises.push(countAndUpdate);
    });
  });

  return Promise.all(updatePromises);
};

//Event update
export const onSubPhaseProgressSharepointList = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<void> => {
  return getDataFromSharepointList(
    spHttpClient,
    sharepointUrl,
    nameSharepointList
  )
    .then((folderPairs) => {
      const updatePromises: Promise<void>[] = [];

      folderPairs.forEach(({ subFolderName, customId }) => {
        updatePromises.push(
          getUrlSubPhaseProgress(
            spHttpClient,
            sharepointUrl,
            subFolderName,
            customId
          ).then((results) => {
            results.forEach(({ childFolderDocumentId, percentFiles }) => {
              console.log(
                `CustomID:${customId}, SubFolderName:${subFolderName}, DocumentId:${childFolderDocumentId}, percentFiles:${percentFiles}`
              );

              return updateSubPhaseProgressSharepointList(
                spHttpClient,
                sharepointUrl,
                nameSharepointList,
                customId,
                childFolderDocumentId,
                percentFiles
              );
            });
          })
        );
      });

      return Promise.all(updatePromises);
    })
    .then(() => {})
    .catch((error) => {
      console.error("Error:", error);
      return Promise.reject(error);
    });
};

//Project Folder-Update progress, documentID lên Project folder-------------------------------------------------------------------------------------------------------------
//Hàm đếm files
//Option1
// const progressFileFolders = (
//   spHttpClient: SPHttpClient,
//   sharepointUrl: string,
//   folderUrls: string[]
// ): Promise<{
//   totalFiles: string;
//   approvedFiles: string;
//   percentFiles: string;
// }> => {
//   const fetchFileCounts = (
//     countFolderUrl: string
//   ): Promise<{ total: number; approved: number }> => {
//     return spHttpClient
//       .get(
//         `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`,
//         SPHttpClient.configurations.v1
//       )
//       .then((response) => {
//         if (!response.ok) {
//           console.warn(`Error: ${response.status} for ${countFolderUrl}`);
//           return { total: 0, approved: 0 };
//         }
//         return response.json();
//       })
//       .then((data) => {
//         const files = data.value || [];
//         const approved = files.filter((file: any) =>
//           file.Name.split(".").slice(0, -1).join(".").endsWith("Approved")
//         ).length;
//         return { total: files.length, approved };
//       })
//       .catch((error) => {
//         console.error(`Error: ${countFolderUrl}`, error);
//         return { total: 0, approved: 0 };
//       });
//   };

//   return Promise.all(folderUrls.map(fetchFileCounts))
//     .then((results) => {
//       const { totalFiles, approvedFiles } = results.reduce(
//         (acc, result) => ({
//           totalFiles: acc.totalFiles + result.total,
//           approvedFiles: acc.approvedFiles + result.approved,
//         }),
//         { totalFiles: 0, approvedFiles: 0 }
//       );

//       return {
//         totalFiles: totalFiles.toString(),
//         approvedFiles: approvedFiles.toString(),
//         percentFiles: totalFiles > 0 ? `${approvedFiles}/${totalFiles}` : "0/0",
//       };
//     })
//     .catch((error) => {
//       console.error("Error:", error);
//       return {
//         totalFiles: "0",
//         approvedFiles: "0",
//         percentFiles: "0/0",
//       };
//     });
// };

//op1
const progressFileFolders = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrls: string[]
): Promise<{
  totalFiles: string;
  approvedFiles: string;
  percentFiles: string;
}> => {
  const fetchFileCounts = (
    countFolderUrl: string
  ): Promise<{ total: number; approved: number }> => {
    return spHttpClient
      .get(
        `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files?$expand=ListItemAllFields`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        if (!response.ok) {
          console.warn(`Error: ${response.status} for ${countFolderUrl}`);
          return { total: 0, approved: 0 };
        }
        return response.json();
      })
      .then((data) => {
        const files = data.value || [];
        const approved = files.filter(
          (file: any) => file.ListItemAllFields?.Status === "Approved"
        ).length;

        return { total: files.length, approved };
      })
      .catch((error) => {
        console.error(`Error: ${countFolderUrl}`, error);
        return { total: 0, approved: 0 };
      });
  };

  return Promise.all(folderUrls.map(fetchFileCounts))
    .then((results) => {
      const { totalFiles, approvedFiles } = results.reduce(
        (acc, result) => ({
          totalFiles: acc.totalFiles + result.total,
          approvedFiles: acc.approvedFiles + result.approved,
        }),
        { totalFiles: 0, approvedFiles: 0 }
      );

      return {
        totalFiles: totalFiles.toString(),
        approvedFiles: approvedFiles.toString(),
        percentFiles: totalFiles > 0 ? `${approvedFiles}/${totalFiles}` : "0/0",
      };
    })
    .catch((error) => {
      console.error("Error:", error);
      return {
        totalFiles: "0",
        approvedFiles: "0",
        percentFiles: "0/0",
      };
    });
};

//Option2
// const progressFileFoldersOption2 = (
//   spHttpClient: SPHttpClient,
//   sharepointUrl: string,
//   folderUrls: string[]
// ): Promise<{
//   totalFiles: number;
//   approvedFiles: number;
//   percentFiles: number;
// }> => {
//   const fetchFileCounts = (
//     countFolderUrl: string
//   ): Promise<{ total: number; approved: number }> => {
//     return spHttpClient
//       .get(
//         `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`,
//         SPHttpClient.configurations.v1
//       )
//       .then((response) => {
//         if (!response.ok) {
//           console.warn(`Error: ${response.status} for ${countFolderUrl}`);
//           return { total: 0, approved: 0 };
//         }
//         return response.json();
//       })
//       .then((data) => {
//         const files = data.value || [];
//         const approved = files.filter((file: any) =>
//           file.Name.split(".").slice(0, -1).join(".").endsWith("Approved")
//         ).length;
//         return { total: files.length, approved };
//       })
//       .catch((error) => {
//         console.error(`Error: ${countFolderUrl}`, error);
//         return { total: 0, approved: 0 };
//       });
//   };

//   return Promise.all(folderUrls.map(fetchFileCounts))
//     .then((results) => {
//       const { totalFiles, approvedFiles } = results.reduce(
//         (acc, result) => ({
//           totalFiles: acc.totalFiles + result.total,
//           approvedFiles: acc.approvedFiles + result.approved,
//         }),
//         { totalFiles: 0, approvedFiles: 0 }
//       );

//       return {
//         totalFiles: totalFiles,
//         approvedFiles: approvedFiles,
//         percentFiles:
//           totalFiles > 0
//             ? parseFloat((approvedFiles / totalFiles).toFixed(4))
//             : 0,
//       };
//     })
//     .catch((error) => {
//       console.error("Error:", error);
//       return {
//         totalFiles: 0,
//         approvedFiles: 0,
//         percentFiles: 0,
//       };
//     });
// };

//op2
const progressFileFoldersOption2 = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  folderUrls: string[]
): Promise<{
  totalFiles: number;
  approvedFiles: number;
  percentFiles: number;
}> => {
  const fetchFileCounts = (
    countFolderUrl: string
  ): Promise<{ total: number; approved: number }> => {
    return spHttpClient
      .get(
        `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files?$expand=ListItemAllFields`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        if (!response.ok) {
          console.warn(`Error: ${response.status} for ${countFolderUrl}`);
          return { total: 0, approved: 0 };
        }
        return response.json();
      })
      .then((data) => {
        const files = data.value || [];
        const approved = files.filter(
          (file: any) => file.ListItemAllFields?.Status === "Approved"
        ).length;

        return { total: files.length, approved };
      })
      .catch((error) => {
        console.error(`Error: ${countFolderUrl}`, error);
        return { total: 0, approved: 0 };
      });
  };

  return Promise.all(folderUrls.map(fetchFileCounts))
    .then((results) => {
      const { totalFiles, approvedFiles } = results.reduce(
        (acc, result) => ({
          totalFiles: acc.totalFiles + result.total,
          approvedFiles: acc.approvedFiles + result.approved,
        }),
        { totalFiles: 0, approvedFiles: 0 }
      );

      return {
        totalFiles: totalFiles,
        approvedFiles: approvedFiles,
        percentFiles:
          totalFiles > 0
            ? parseFloat((approvedFiles / totalFiles).toFixed(4))
            : 0,
      };
    })
    .catch((error) => {
      console.error("Error:", error);
      return {
        totalFiles: 0,
        approvedFiles: 0,
        percentFiles: 0,
      };
    });
};

//Hàm update Progress và DocumentID cho thư mục
//Option1
const updateProgressFileFolders = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  percentFiles: string,
  folderUrl: string,
  documentId: string
): Promise<any> => {
  return spHttpClient
    .get(
      `${sharepointUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields`,
      SPHttpClient.configurations.v1
    )
    .then((response) =>
      response.ok
        ? response.json()
        : Promise.reject("Folder metadata not found")
    )
    .then((data) => {
      if (!data.Id) {
        return Promise.reject("Item ID not found");
      }

      const body = JSON.stringify({
        __metadata: { type: "SP.ListItem" },
        ProgressOp1: percentFiles,
        DocumentID: documentId,
      });

      const optionsHTTP: ISPHttpClientOptions = {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "odata-version": "",
          "If-Match": "*",
          "X-HTTP-Method": data.Id ? "MERGE" : "POST",
        },
        body,
      };

      const url = data.Id
        ? `${sharepointUrl}/_api/web/lists/getByTitle('ProjectFolder')/items(${data.Id})`
        : `${sharepointUrl}/_api/web/lists/getByTitle('ProjectFolder')/items`;

      return spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        optionsHTTP
      );
    })
    .catch((error) => console.error("Error:", error));
};

//Option2
const updateProgressFileFoldersOption2 = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  percentFiles: number,
  folderUrl: string,
  documentId: string
): Promise<any> => {
  const requestUrl = `${sharepointUrl}/_api/web/getFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields`;

  return spHttpClient
    .get(requestUrl, SPHttpClient.configurations.v1)
    .then((response) =>
      response.ok
        ? response.json()
        : Promise.reject("Folder metadata not found")
    )
    .then((data) => {
      if (!data.Id) {
        return Promise.reject("Item ID not found");
      }
      const body = JSON.stringify({
        __metadata: { type: "SP.ListItem" },
        ProgressOp2: percentFiles,
        DocumentID: documentId,
      });

      const optionsHTTP: ISPHttpClientOptions = {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "odata-version": "",
          "If-Match": "*",
          "X-HTTP-Method": data.Id ? "MERGE" : "POST",
        },
        body,
      };

      const url = data.Id
        ? `${sharepointUrl}/_api/web/lists/getByTitle('ProjectFolder')/items(${data.Id})`
        : `${sharepointUrl}/_api/web/lists/getByTitle('ProjectFolder')/items`;

      return spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        optionsHTTP
      );
    })
    .catch((error) => console.error("Error:", error));
};

//Hàm lấy url các thư mục
//Option1
const getUrlProgressFolders = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  subFolderName: string | string[]
): Promise<void> => {
  const subFolderNames = Array.isArray(subFolderName)
    ? subFolderName
    : [subFolderName];
  const subFolders = Object.keys(childSubFolders);
  const updatePromises: Promise<void>[] = [];

  subFolders.forEach((folder) => {
    const baseFolderUrl = `ProjectFolder/PROJECT/${subFolderNames}/${folder}`;
    const childFolders = childSubFolders[folder];

    childFolders.forEach((childFolder) => {
      const childFolderName = childFolder.name;
      const childFolderUrl = `${baseFolderUrl}/${childFolderName}`;
      const documentId = childFolder.DocumentId;
      const countAndUpdate = progressFileFolders(spHttpClient, sharepointUrl, [
        childFolderUrl,
      ])
        .then(({ percentFiles }) => {
          //console.log(childFolderUrl, percentFiles);
          return updateProgressFileFolders(
            spHttpClient,
            sharepointUrl,
            percentFiles,
            childFolderUrl,
            documentId
          );
        })
        .catch((error) => {
          console.error(`Error ${childFolderUrl}:`, error);
        });

      updatePromises.push(countAndUpdate);
    });
  });
  return Promise.all(updatePromises).then(() => {});
};

//Option2
// const getUrlProgressFoldersOption2 = (
//   spHttpClient: SPHttpClient,
//   sharepointUrl: string,
//   subFolderName: string | string[]
// ): Promise<void> => {
//   const subFolderNames = Array.isArray(subFolderName)
//     ? subFolderName
//     : [subFolderName];
//   const subFolders = Object.keys(childSubFolders);
//   const updatePromises: Promise<void>[] = [];

//   subFolders.forEach((folder) => {
//     const baseFolderUrl = `ProjectFolder/PROJECT/${subFolderNames}/${folder}`;
//     const childFolders = childSubFolders[folder];

//     childFolders.forEach((childFolder) => {
//       const childFolderName = childFolder.name;
//       const childFolderDocumentId = childFolder.DocumentId;
//       const childFolderUrl = `${baseFolderUrl}/${childFolderName}`;
//       const documentId = childFolder.DocumentId;
//       const countAndUpdate = progressFileFoldersOption2(
//         spHttpClient,
//         sharepointUrl,
//         [childFolderUrl]
//       )
//         .then(({ percentFiles }) => {
//           //console.log(childFolderUrl, percentFiles);
//           console.log(subFolderNames, childFolderDocumentId, percentFiles);
//           return updateProgressFileFoldersOption2(
//             spHttpClient,
//             sharepointUrl,
//             percentFiles,
//             childFolderUrl,
//             documentId
//           );
//         })
//         .catch((error) => {
//           console.error(`Error ${childFolderUrl}:`, error);
//         });

//       updatePromises.push(countAndUpdate);
//     });
//   });
//   return Promise.all(updatePromises).then(() => {});
// };

//thay đổi
const getUrlProgressFoldersOption2 = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  subFolderName: string | string[]
): Promise<
  {
    subFolderNames: string[];
    childFolderDocumentId: string;
    percentFiles: number;
  }[]
> => {
  const subFolderNames = Array.isArray(subFolderName)
    ? subFolderName
    : [subFolderName];
  const subFolders = Object.keys(childSubFolders);
  const updatePromises: Promise<{
    subFolderNames: string[];
    childFolderDocumentId: string;
    percentFiles: number;
  }>[] = [];

  subFolders.forEach((folder) => {
    const baseFolderUrl = `ProjectFolder/PROJECT/${subFolderNames}/${folder}`;
    const childFolders = childSubFolders[folder];

    childFolders.forEach((childFolder) => {
      const childFolderName = childFolder.name;
      const childFolderDocumentId = childFolder.DocumentId;
      const childFolderUrl = `${baseFolderUrl}/${childFolderName}`;
      const documentId = childFolder.DocumentId;

      const countAndUpdate = progressFileFoldersOption2(
        spHttpClient,
        sharepointUrl,
        [childFolderUrl]
      )
        .then(({ percentFiles }) => {
          return updateProgressFileFoldersOption2(
            spHttpClient,
            sharepointUrl,
            percentFiles,
            childFolderUrl,
            documentId
          ).then(() => ({
            subFolderNames,
            childFolderDocumentId,
            percentFiles,
          }));
        })
        .catch((error) => {
          console.error(`Error ${childFolderUrl}:`, error);

          return {
            subFolderNames,
            childFolderDocumentId,
            percentFiles: 0,
          };
        });

      updatePromises.push(countAndUpdate);
    });
  });

  return Promise.all(updatePromises);
};

//Event đếm file và update giá trị cột Progress và DocumentID lên ProjectFolder
//Option1
export const onProgressFolders = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<void> => {
  return getDataFromSharepointList(
    spHttpClient,
    sharepointUrl,
    nameSharepointList
  )
    .then((folderPairs) => {
      const subFolder = folderPairs.map(({ subFolderName }) => subFolderName);
      const updatePromises: Promise<void>[] = [];

      subFolder.forEach((subFolderName) => {
        updatePromises.push(
          getUrlProgressFolders(spHttpClient, sharepointUrl, subFolderName)
        );
      });

      return Promise.all(updatePromises).then(() => {});
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};

//Option2
// export const onProgressFoldersOption2 = (
//   spHttpClient: SPHttpClient,
//   sharepointUrl: string,
//   nameSharepointList: string
// ): Promise<void> => {
//   return getDataFromSharepointList(
//     spHttpClient,
//     sharepointUrl,
//     nameSharepointList
//   )
//     .then((folderPairs) => {
//       const subFolder = folderPairs.map(({ subFolderName }) => subFolderName);
//       const updatePromises: Promise<void>[] = [];

//       subFolder.forEach((subFolderName) => {
//         updatePromises.push(
//           getUrlProgressFoldersOption2(
//             spHttpClient,
//             sharepointUrl,
//             subFolderName
//           )
//         );
//       });
//       return Promise.all(updatePromises).then(() => {
//         console.log(
//           `The Progress column was updated successfully in ProjectFolder Option2`
//         );
//       });
//     })
//     .catch((error) => {
//       console.error("Error:", error);
//     });
// };

//thay đổi
export const onProgressFoldersOption2 = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<void> => {
  return getDataFromSharepointList(
    spHttpClient,
    sharepointUrl,
    nameSharepointList
  )
    .then((folderPairs) => {
      const subFolder = folderPairs.map(({ subFolderName }) => subFolderName);
      const updatePromises: Promise<
        {
          subFolderNames: string[];
          childFolderDocumentId: string;
          percentFiles: number;
        }[]
      >[] = [];

      subFolder.forEach((subFolderName) => {
        updatePromises.push(
          getUrlProgressFoldersOption2(
            spHttpClient,
            sharepointUrl,
            subFolderName
          )
        );
      });

      return Promise.all(updatePromises)
        .then((results) => {
          results.forEach((resultArray) => {
            resultArray.forEach(
              ({ subFolderNames, childFolderDocumentId, percentFiles }) => {
                console.log(
                  `SubFolderName: ${subFolderNames}, DocumentId: ${childFolderDocumentId}, PercentFiles: ${percentFiles}`
                );
              }
            );
          });
        })
        .catch((error) => {
          console.error("Error:", error);
        });
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};
