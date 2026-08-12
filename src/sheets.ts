/**
 * Google Sheets 자동화 모듈
 */

/**
 * 신규 스프레드시트를 생성하고 데이터를 기입합니다.
 */
export function createSpreadsheetWithData(title: string, headers: string[][], rows: (string | number)[][]): GoogleAppsScript.Spreadsheet.Spreadsheet {
  const ss = SpreadsheetApp.create(title);
  const sheet = ss.getActiveSheet();

  // 헤더 작성 및 스타일 지정
  if (headers.length > 0) {
    const headerRange = sheet.getRange(1, 1, headers.length, headers[0].length);
    headerRange.setValues(headers);
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
  }

  // 본문 행 작성
  if (rows.length > 0) {
    const dataRange = sheet.getRange(2, 1, rows.length, rows[0].length);
    dataRange.setValues(rows);
  }

  sheet.autoResizeColumns(1, headers[0]?.length || 1);
  Logger.log(`[Sheets] 스프레드시트 생성 완료: ${ss.getUrl()}`);
  return ss;
}
