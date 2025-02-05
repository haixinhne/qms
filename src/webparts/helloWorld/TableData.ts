import "datatables.net";
import * as $ from "jquery";
import { SPHttpClient } from "@microsoft/sp-http";
import { exportDataFromSharepointListUseTable } from "./ExportDataFromSharepointListUseTable";

export interface dataTable {
  CustomID: string;
  ProjectName: string;
  Nation: string;
  Phase01Date: string;
  Phase01Progress: number;
  Phase02Date: string;
  Phase02Progress: number;
  Phase03Date: string;
  Phase03Progress: number;
}

//Hàm định dạng ngày
const formatDate = (dateString: string): string =>
  new Date(dateString).toLocaleDateString("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });

//Hàm tạo bảng, cập nhật, xóa hàng
export function initializeTableFromSharepoint(
  spHttpClient: SPHttpClient,
  sharepointUrl: string,
  nameSharepointList: string
): void {
  exportDataFromSharepointListUseTable(
    spHttpClient,
    sharepointUrl,
    nameSharepointList
  )
    .then((data) => {
      if (!data || data.length === 0) {
        console.warn("No data from SharePoint");
        return;
      }

      const tableElement = $("#tableDataSharepointList");

      if (!$.fn.DataTable.isDataTable(tableElement)) {
        tableElement.DataTable({
          info: false,
          paging: false,
          searching: false,
          lengthChange: false,
          autoWidth: false,
          data: [],
          columns: [
            { title: "CustomID" },
            { title: "ProjectName" },
            { title: "Nation" },
            { title: "Phase01 Date" },
            { title: "Phase01 Progress" },
            { title: "Phase02 Date" },
            { title: "Phase02 Progress" },
            { title: "Phase03 Date" },
            { title: "Phase03 Progress" },
          ],
        });
      }

      const table = tableElement.DataTable();

      //Lấy CustomID của dữ liệu mới
      const newCustomID: { [key: string]: boolean } = {};
      for (let i = 0; i < data.length; i++) {
        newCustomID[data[i].CustomID] = true;
      }

      //Lưu index của các hàng hiện tại để kiểm tra hàng cần xóa
      const rowDelete: number[] = [];
      table.rows().every(function () {
        const rowData = this.data();
        const customID = rowData[0];

        if (newCustomID[customID]) {
          //Cập nhật dữ liệu nếu CustomID đã tồn tại
          for (let i = 0; i < data.length; i++) {
            if (data[i].CustomID === customID) {
              this.data([
                data[i].CustomID,
                data[i].ProjectName,
                data[i].Nation,
                formatDate(data[i].Phase01Date),
                `${data[i].Phase01Progress}%`,
                formatDate(data[i].Phase02Date),
                `${data[i].Phase02Progress}%`,
                formatDate(data[i].Phase03Date),
                `${data[i].Phase03Progress}%`,
              ]);
              break;
            }
          }
        } else {
          //Nếu không có trong danh sách mới, đánh dấu để xóa
          rowDelete.push(this.index());
        }
      });

      //Xóa các hàng không có trong danh sách mới
      for (let i = rowDelete.length - 1; i >= 0; i--) {
        table.row(rowDelete[i]).remove();
      }

      //Thêm dữ liệu mới nếu chưa tồn tại
      for (let i = 0; i < data.length; i++) {
        let exists = false;
        for (let j = 0; j < table.rows().count(); j++) {
          if (table.cell(j, 0).data() === data[i].CustomID) {
            exists = true;
            break;
          }
        }
        if (!exists) {
          table.row.add([
            data[i].CustomID,
            data[i].ProjectName,
            data[i].Nation,
            formatDate(data[i].Phase01Date),
            `${data[i].Phase01Progress}%`,
            formatDate(data[i].Phase02Date),
            `${data[i].Phase02Progress}%`,
            formatDate(data[i].Phase03Date),
            `${data[i].Phase03Progress}%`,
          ]);
        }
      }

      table.draw(false);
    })
    .catch((error) => {
      console.error("Error", error);
    });
}
