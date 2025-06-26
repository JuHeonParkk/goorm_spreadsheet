const spreadsheetContainer = document.querySelector("#sheet_containter");
const exportBtn = document.querySelector("#export_btn");

// 10X10 테이블 정의
const ROWS = 10;
const COLUMNS = 10;
const spreadsheet = [];

// cell 클래스 정의
class Cell {
  constructor(
    isHeader,
    disabled,
    data,
    row,
    column,
    rowName,
    columnName,
    active
  ) {
    this.isHeader = isHeader; // 헤더 여부
    this.disabled = disabled; // 비활성화 여부
    this.data = data; // 셀에 입력될 값
    this.row = row; // 행 인덱스
    this.column = column; // 열 인덱스
    this.rowName = rowName; // 행 이름(예: 1, 2, 3 등)
    this.columnName = columnName; // 열 이름(예: A, B, C 등)
    this.active = active; // 활성화 여부
  }
}

// export 버튼 클릭 시 스프레드시트 데이터를 CSV로 변환
exportBtn.onclick = () => handleClickExportBtn();

function handleClickExportBtn(e) {
  let csv = "";
  for (let i = 0; i < spreadsheet.length; i++) {
    if (i === 0) continue; // 첫 번째 행은 헤더이므로 건너뜀
    csv +=
      spreadsheet[i]
        .filter((cell) => !cell.isHeader) // 헤더 셀 제외
        .map((cell) => cell.data) // 셀 데이터만 추출
        .join(",") + "\r\n"; // 쉼표로 구분하고 줄바꿈 추가
  }

  const blob = new Blob([csv]);
  //console.log("blob", blob);

  const csvURL = URL.createObjectURL(blob);

  const downloadLink = document.createElement("a");
  downloadLink.href = csvURL; // 생성된 Blob URL 설정
  downloadLink.download = "spreadsheet.csv"; // 다운로드할 파일 이름 설정
  downloadLink.click(); // 다운로드 링크 클릭하여 파일 다운로드
  URL.revokeObjectURL(csvURL); // Blob URL 해제
}

initSpreadsheet(); // 스프레드시트 초기화

// spreadsheet 초기화
function initSpreadsheet() {
  for (let i = 0; i < ROWS; i++) {
    let row = []; // 행 배열 초기화
    for (let j = 0; j < COLUMNS; j++) {
      let data = ""; // 초기 데이터는 빈 문자열
      let isHeader = false;
      let disabled = false;
      let active = false; // 활성화 여부

      if (j === 0) {
        data = i;
        isHeader = true; // 첫 번째 열은 행 헤더
        disabled = true; // 행 헤더는 비활성화
      }
      if (i === 0) {
        data = String.fromCharCode(65 + (j - 1)); // A, B, C 등 열 헤더
        isHeader = true; // 첫 번째 행은 열 헤더
        disabled = true; // 열 헤더는 비활성화
      }

      if (i === 0 && j === 0) {
        data = ""; // 첫 번째 셀은 비워둠
      }

      const rowName = i;
      const columnName = String.fromCharCode(65 + (j - 1)); // A, B, C 등 열 이름

      const cell = new Cell(
        isHeader,
        disabled,
        data,
        i,
        j,
        rowName,
        columnName,
        active
      );
      row.push(cell); // 행에 셀 추가
    }
    spreadsheet.push(row); // 스프레드시트에 행 추가
  }
  drawSpreadsheet(); // 스프레드시트 그리기
}

// spreadsheet 그리기
function createCellElement(cell) {
  const cellElement = document.createElement("input");
  cellElement.className = "cell";
  cellElement.id = `cell_${cell.row}${cell.column}`; // 셀 ID 설정
  cellElement.value = cell.data; // 셀 데이터 설정
  cellElement.disabled = cell.disabled; // 셀 비활성화 설정

  if (cell.isHeader) {
    cellElement.classList.add("header_cell"); // 헤더  셀 클래스 추가
  }

  // 셀 클릭 시 이벤트 핸들러 설정
  cellElement.onclick = () => handleCellClick(cell);
  // 셀 값 변경 시 데이터 저장
  cellElement.onchange = (e) => handleChangeData(e.target.value, cell);

  return cellElement;
}

function drawSpreadsheet() {
  for (let i = 0; i < spreadsheet.length; i++) {
    const rowContainerElement = document.createElement("div");
    rowContainerElement.className = "row_container";
    for (let j = 0; j < spreadsheet[i].length; j++) {
      const cell = spreadsheet[i][j];
      rowContainerElement.append(createCellElement(cell)); // 셀 요소 생성 및 추가
    }
    spreadsheetContainer.append(rowContainerElement); // 행 컨테이너를 스프레드시트 컨테이너에 추가
  }
}

// 사용자가 입력한 값 셀 객체에 저장
function handleChangeData(data, cell) {
  cell.data = data;
}

// spreadsheet 클릭 시 header 셀 활성화
function handleCellClick(cell) {
  clearHeaderActiveStates(); // 이전 활성화 상태 제거
  const columnHeader = spreadsheet[0][cell.column];
  const rowHeader = spreadsheet[cell.row][0];
  const colHeaderElement = getElementROWCOL(
    columnHeader.row,
    columnHeader.column
  );
  const rowHeaderElement = getElementROWCOL(rowHeader.row, rowHeader.column);

  colHeaderElement.classList.add("active"); // 열 헤더 활성화
  rowHeaderElement.classList.add("active"); // 행 헤더 활성화

  // 셀 상태 표시
  document.querySelector(
    "#cell_state"
  ).innerHTML = `Cell : ${cell.rowName}${cell.columnName} `;
}

function clearHeaderActiveStates() {
  const headerCells = document.querySelectorAll(".header_cell");

  headerCells.forEach((header_cell) => {
    header_cell.classList.remove("active");
  });
}

function getElementROWCOL(row, column) {
  return document.querySelector(`#cell_${row}${column}`);
}
