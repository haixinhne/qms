import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";

//Tên subfolders
export const childSubFolders: { [key: string]: string[] } = {
  Promotion: [
    "Basic project data",
    "Promotion activities",
    "Reports",
    "Project risk analysis",
    "Stakeholder Management",
    "Design Review 01 AS-ME",
    "Client Contract Review (CCR)",
    "Project Approval (EU-Kento)",
    "Estimate Approval (EU-Kessai)",
  ],
  Design: [
    "Drawings",
    "Funtion Checklist AS",
    "Funtion Checklist ME",
    "Designer Approval Request",
    "Design Review 23 AS-ME",
  ],
  Build: [
    "Project outline",
    "PM Policy",
    "HSE risks (Health, Safely & Env.ment)",
    "Funtion Checklist Update",
    "Quality Plan",
    "Schedule",
    "Construction Kickoff",
  ],
};

//Đọc sharepoint list
const getDataFromSharepointList = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): Promise<{ subFolderName: string }[]> => {
  const listUrl = `${sharepointUrl}/_api/web/lists/GetByTitle('${nameSharepointList}')/items`;

  return spHttpClient
    .get(listUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    })
    .then((data) => {
      const folderValues = data.value
        //Tên folder con = tên cột ProjectName
        .filter((item: any) => item.Nation && item.ProjectName)
        .map((item: any) => ({
          subFolderName: item.ProjectName,
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
      console.error("Error fetching SharePoint list data:", error);
    });
};

//Sharepoint-------------------------------------------------------------------------------------------------------------------------------------
//Đếm file
const countFiles = (
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
        `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        if (!response.ok) {
          console.log(`HTTP error! Status: ${response.status}`);
        }
        return response.json();
      })
      .then((data) => {
        const files = data.value || [];
        const total = files.length;

        const approved = files.filter((file: any) => {
          const fileNameWithoutExtension = file.Name.split(".")
            .slice(0, -1)
            .join(".");
          return fileNameWithoutExtension.endsWith("Approved");
        }).length;

        return { total, approved };
      })
      .catch((error) => {
        console.error(`Error fetching files from ${countFolderUrl}:`, error);
        return { total: 0, approved: 0 };
      });
  };

  const loopFolders = folderUrls.map((url: string) => fetchFileCounts(url)); //Lặp qua các thư mục
  return Promise.all(loopFolders).then((results) => {
    const totalFiles = results.reduce((sum, result) => sum + result.total, 0);
    const approvedFiles = results.reduce(
      (sum, result) => sum + result.approved,
      0
    );
    const percentFiles =
      totalFiles > 0 ? parseFloat((approvedFiles / totalFiles).toFixed(2)) : 0;

    return { totalFiles, approvedFiles, percentFiles };
  });
};

//Lấy Url các thư mục
const getUrlCountFiles = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  subFolderName: string | string[]
): Promise<any> => {
  if (typeof subFolderName === "string") {
    subFolderName = [subFolderName];
  }

  const subFolderUrl = `ProjectFolder/PROJECT/${subFolderName}`;
  const subFolders = ["Promotion", "Design", "Build"];
  const arrayFolderUrl: string[] = [];

  subFolders.forEach((folder) => {
    const folderUrl = `${subFolderUrl}/${folder}`;
    arrayFolderUrl.push(folderUrl);
    ``;
    const childFolders = childSubFolders[folder];
    childFolders.forEach((childFolder) => {
      const childFolderUrl = `${folderUrl}/${childFolder}`;
      arrayFolderUrl.push(childFolderUrl);
    });
  });

  return countFiles(spHttpClient, sharepointUrl, arrayFolderUrl)
    .then(({ totalFiles, approvedFiles, percentFiles }) => {
      console.log(`Total Files in ${subFolderName}: ${totalFiles}`);
      console.log(`Approved Files in ${subFolderName}: ${approvedFiles}`);
      console.log(`Completion rate in ${subFolderName}: ${percentFiles}`);
      return { totalFiles, approvedFiles, percentFiles };
    })
    .catch((error) => {
      console.error("Error counting files:", error);
    });
};

//Update Rate lên sharepoint list
const updateRateSharepoint = (
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
          Rate: rateValue,
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
              return response.text().then(() => {
                Promise.reject(
                  `Failed to update Rate for item ${itemId}: ${response.statusText}`
                );
              });
            }
            console.log(
              `All items updated completion rate in ${nameSharepointList}`
            );
          });
      }
      return Promise.reject(`No item found for ProjectName: ${subFolderName}`);
    })
    .catch((error) => {
      console.error("Error updating Rate value:", error);
      return Promise.reject(error);
    });
};

//Click đếm, update file
export const onCountFiles = (
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
          getUrlCountFiles(spHttpClient, sharepointUrl, subFolderName).then(
            ({ percentFiles }) => {
              return updateRateSharepoint(
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
    .catch((error) => {
      console.error("Error processing folders and subfolders:", error);
    });
};

//Folder-----------------------------------------------------------------------------------------------------------------------------------------
//Đếm files
//Option1
const countFilesFolders = (
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
        `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        if (!response.ok) {
          console.warn(
            `HTTP error! Status: ${response.status} for ${countFolderUrl}`
          );
          return { total: 0, approved: 0 };
        }
        return response.json();
      })
      .then((data) => {
        const files = data.value || [];
        const approved = files.filter((file: any) =>
          file.Name.split(".").slice(0, -1).join(".").endsWith("Approved")
        ).length;
        return { total: files.length, approved };
      })
      .catch((error) => {
        console.error(`Error fetching files from ${countFolderUrl}:`, error);
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
      console.error("Error processing folders:", error);
      return {
        totalFiles: "0",
        approvedFiles: "0",
        percentFiles: "0/0",
      };
    });
};

//Option2
const countFilesFoldersOption2 = (
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
        `${sharepointUrl}/_api/web/GetFolderByServerRelativeUrl('${countFolderUrl}')/Files`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        if (!response.ok) {
          console.warn(
            `HTTP error! Status: ${response.status} for ${countFolderUrl}`
          );
          return { total: 0, approved: 0 };
        }
        return response.json();
      })
      .then((data) => {
        const files = data.value || [];
        const approved = files.filter((file: any) =>
          file.Name.split(".").slice(0, -1).join(".").endsWith("Approved")
        ).length;
        return { total: files.length, approved };
      })
      .catch((error) => {
        console.error(`Error fetching files from ${countFolderUrl}:`, error);
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
            ? parseFloat((approvedFiles / totalFiles).toFixed(2))
            : 0,
      };
    })
    .catch((error) => {
      console.error("Error processing folders:", error);
      return {
        totalFiles: 0,
        approvedFiles: 0,
        percentFiles: 0,
      };
    });
};

//Lấy Url các thư mục
//Option1
const getUrlCountFilesFolders = (
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

    childFolders.forEach((child) => {
      const childFolderUrl = `${baseFolderUrl}/${child}`;
      //Đếm file trong thư mục này và cập nhật Approved
      const countAndUpdate = countFilesFolders(spHttpClient, sharepointUrl, [
        childFolderUrl,
      ])
        .then(({ percentFiles }) => {
          console.log(`Updating folder: ${childFolderUrl}: ${percentFiles}`);
          return updateFolderApprovedFolders(
            spHttpClient,
            sharepointUrl,
            percentFiles,
            childFolderUrl
          );
        })
        .catch((error) => {
          console.error(`Error updating folder ${childFolderUrl}:`, error);
        });

      updatePromises.push(countAndUpdate);
    });
  });

  return Promise.all(updatePromises).then(() => {
    console.log(`All updates for ${subFolderNames} completed`);
  });
};

//Option2
const getUrlCountFilesFoldersOption2 = (
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

    childFolders.forEach((child) => {
      const childFolderUrl = `${baseFolderUrl}/${child}`;
      //Đếm file trong thư mục này và cập nhật Approved
      const countAndUpdate = countFilesFoldersOption2(
        spHttpClient,
        sharepointUrl,
        [childFolderUrl]
      )
        .then(({ percentFiles }) => {
          console.log(`Updating folder: ${childFolderUrl}: ${percentFiles}`);
          return updateFolderApprovedFoldersOption2(
            spHttpClient,
            sharepointUrl,
            percentFiles,
            childFolderUrl
          );
        })
        .catch((error) => {
          console.error(`Error updating folder ${childFolderUrl}:`, error);
        });

      updatePromises.push(countAndUpdate);
    });
  });

  return Promise.all(updatePromises).then(() => {
    console.log(`All updates for ${subFolderNames} completed`);
  });
};

//Update Rate cho thư mục
//Option1
const updateFolderApprovedFolders = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  approvedValue: string,
  folderUrl: string
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
        console.error("Item ID not found");
        return Promise.reject("Item ID not found");
      }
      const body = JSON.stringify({
        __metadata: { type: "SP.ListItem" },
        ProgressOp1: approvedValue,
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
    .catch((error) => console.error("Error updating Approved column:", error));
};

//Option2
const updateFolderApprovedFoldersOption2 = (
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  approvedValue: number,
  folderUrl: string
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
        console.error("Item ID not found");
        return Promise.reject("Item ID not found");
      }
      const body = JSON.stringify({
        __metadata: { type: "SP.ListItem" },
        ProgressOp2: approvedValue,
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
    .catch((error) => console.error("Error updating Approved column:", error));
};

//Click đếm file và update giá trị cột Approved vào Folder
//Option1
export const onCountFilesFolders = (
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
          getUrlCountFilesFolders(spHttpClient, sharepointUrl, subFolderName)
        );
      });

      return Promise.all(updatePromises).then(() => {
        console.log("The number of files updated in Op1");
        alert("The number of files updated in Op1");
      });
    })
    .catch((error) => {
      console.error("Error processing folders and subfolders:", error);
    });
};

//Option2
export const onCountFilesFoldersOption2 = (
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
          getUrlCountFilesFoldersOption2(
            spHttpClient,
            sharepointUrl,
            subFolderName
          )
        );
      });
      return Promise.all(updatePromises).then(() => {
        console.log("The number of files updated in Op2");
        alert("The number of files updated in Op2");
      });
    })

    .catch((error) => {
      console.error("Error processing folders and subfolders:", error);
    });
};
