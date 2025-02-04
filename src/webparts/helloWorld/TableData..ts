import "datatables.net";
import * as $ from "jquery";

export interface dataFromSharepointList {
  CustomID: string;
  ProjectName: string;
  Nation: string;
  Phase01Date: string;
  Phase02Date: string;
  Phase03Date: string;
  Phase01Progress: number;
  Phase02Progress: number;
  Phase03Progress: number;
}

export const dataFromSharepointList: dataFromSharepointList[] = [
  {
    CustomID: "01",
    ProjectName: "VN-QMS",
    Nation: "Viet Nam-VN",
    Phase01Date: "2025-01-23T08:00:00Z",
    Phase02Date: "2025-01-22T08:00:00Z",
    Phase03Date: "2025-01-08T08:00:00Z",
    Phase01Progress: 1,
    Phase02Progress: 2,
    Phase03Progress: 3,
  },
  {
    CustomID: "02",
    ProjectName: "Japan-CT5",
    Nation: "Japan",
    Phase01Date: "2025-01-23T08:00:00Z",
    Phase02Date: "2025-01-22T08:00:00Z",
    Phase03Date: "2025-01-08T08:00:00Z",
    Phase01Progress: 4,
    Phase02Progress: 5,
    Phase03Progress: 6,
  },
  {
    CustomID: "03",
    ProjectName: "VN-TH",
    Nation: "Viet Nam-VN",
    Phase01Date: "2025-01-23T08:00:00Z",
    Phase02Date: "2025-01-22T08:00:00Z",
    Phase03Date: "2025-01-08T08:00:00Z",
    Phase01Progress: 7,
    Phase02Progress: 8,
    Phase03Progress: 9,
  },
];

//Hàm tạo bảng
export function initializeTable(data: dataFromSharepointList[]): void {
  $(() => {
    const table = $("#myTable");
    if ($.fn.DataTable.isDataTable(table)) {
      table.DataTable().destroy();
    }

    data.forEach((row) => addRow(row));

    table.DataTable({
      info: false,
      paging: false,
      searching: false,
      lengthChange: false,
      autoWidth: false,
      columnDefs: [
        { width: "50px", targets: 0 },
        { width: "50px", targets: 1 },
        { width: "50px", targets: 2 },
        { width: "50px", targets: 3 },
        { width: "50px", targets: 4 },
        { width: "50px", targets: 5 },
        { width: "50px", targets: 6 },
        { width: "50px", targets: 7 },
        { width: "50px", targets: 8 },
      ],
    });
  });
}

//Hàm thêm hàng vào bảng
function addRow(data: dataFromSharepointList): void {
  const table = document
    .getElementById("myTable")
    ?.getElementsByTagName("tbody")[0];
  if (!table) return;

  const newRow = table.insertRow();
  newRow.insertCell(0).textContent = data.CustomID;
  newRow.insertCell(1).textContent = data.ProjectName;
  newRow.insertCell(2).textContent = data.Nation;
  newRow.insertCell(3).textContent = formatDate(data.Phase01Date);
  newRow.insertCell(4).textContent = formatDate(data.Phase02Date);
  newRow.insertCell(5).textContent = formatDate(data.Phase03Date);
  newRow.insertCell(6).textContent = `${data.Phase01Progress}%`;
  newRow.insertCell(7).textContent = `${data.Phase02Progress}%`;
  newRow.insertCell(8).textContent = `${data.Phase03Progress}%`;
}

//Hàm định dạng ngày
const formatDate = (dateString: string): string =>
  new Date(dateString).toLocaleDateString("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
