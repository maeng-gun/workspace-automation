/**
 * 구글드라이브 파일 관리 모듈 (Drive Manager)
 * 스프레드시트 ID: 1-MKH0ug-IfN4HdpH5sSrwuq94eXB6pQumfTwFfpWZeo
 * 
 * 탭 구성:
 * 1) [현재 파일] 탭: 전체 파일 목록 생성 대상 (11개 열)
 * 2) [변경 대상] 탭: 일괄 삭제 및 일괄 이동 대상 (17개 열)
 * 3) [폴더일괄변경] 탭: 폴더 계층 내용물 일괄 이동 및 빈 원본 폴더 삭제 (10개 열)
 */

const DRIVE_MANAGEMENT_SPREADSHEET_ID = '1-MKH0ug-IfN4HdpH5sSrwuq94eXB6pQumfTwFfpWZeo';

/** 탐색 제외 폴더 목록 */
const EXCLUDED_FOLDERS: string[] = ['하영영어'];

/**
 * 스프레드시트 또는 활성 시트를 가져오는 헬퍼 함수
 */
function getTargetSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  try {
    return SpreadsheetApp.openById(DRIVE_MANAGEMENT_SPREADSHEET_ID);
  } catch (e) {
    Logger.log(`[DriveManager] openById 실패, getActiveSpreadsheet 시도: ${e}`);
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * 1. 전체 파일 목록 생성 (100개 청크 단위 기록)
 * 대상 탭: [현재 파일]
 */
function generateDriveFileList() {
  const ss = getTargetSpreadsheet();
  let sheet = ss.getSheetByName('현재 파일');

  if (!sheet) {
    sheet = ss.insertSheet('현재 파일');
    Logger.log('[DriveManager] [현재 파일] 탭이 없어 새로 생성했습니다.');
  }

  // 1. 헤더 설정 확인/작성 (11개 열)
  const headers = [
    ['파일명', '파일 URL', '파일 종류', '용량(MB)', '최종 수정일', '파일ID', '레벨1', '레벨2', '레벨3', '레벨4', '레벨5']
  ];
  sheet.getRange(1, 1, 1, 11).setValues(headers);
  sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#e6f2ff');

  // 2. 기존 데이터 영역 초기화 (2행부터)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 11).clearContent();
  }

  Logger.log('[DriveManager] 전체 파일 목록 탐색 시작...');
  const toastMessage = (msg: string) => ss.toast(msg, '드라이브 동기화', 5);
  toastMessage('구글 드라이브 전체 파일 목록 탐색을 시작합니다...');

  let chunk: (string | number)[][] = [];
  let currentRow = 2;
  let totalProcessed = 0;
  const CHUNK_SIZE = 100;

  const flushChunk = () => {
    if (chunk.length === 0) return;
    sheet.getRange(currentRow, 1, chunk.length, 11).setValues(chunk);
    currentRow += chunk.length;
    totalProcessed += chunk.length;
    chunk = [];
    SpreadsheetApp.flush();
    toastMessage(`현재 ${totalProcessed}개 파일 기록 중...`);
    Logger.log(`[DriveManager] ${totalProcessed}개 파일 시트 작성 완료.`);
  };

  // 재귀 탐색 함수
  const processFile = (file: GoogleAppsScript.Drive.File, pathLevels: string[]) => {
    const sizeInMB = (file.getSize() / (1024 * 1024)).toFixed(2);
    const lastUpdated = Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // 5단계 레벨 구성
    const level1 = pathLevels[0] || '';
    const level2 = pathLevels[1] || '';
    const level3 = pathLevels[2] || '';
    const level4 = pathLevels[3] || '';
    const level5 = pathLevels[4] || '';

    const row = [
      file.getName(),
      file.getUrl(),
      file.getMimeType(),
      parseFloat(sizeInMB),
      lastUpdated,
      file.getId(),
      level1,
      level2,
      level3,
      level4,
      level5
    ];

    chunk.push(row);
    if (chunk.length >= CHUNK_SIZE) {
      flushChunk();
    }
  };

  const traverseFolder = (folder: GoogleAppsScript.Drive.Folder, currentPath: string[]) => {
    const folderName = folder.getName();
    if (EXCLUDED_FOLDERS.includes(folderName)) {
      Logger.log(`[DriveManager] 제외된 폴더 탐색 스킵: ${folderName}`);
      return;
    }

    // 현재 폴더 직하위 파일 탐색
    const files = folder.getFiles();
    while (files.hasNext()) {
      processFile(files.next(), currentPath);
    }

    // 하위 폴더 재귀 탐색 (최대 5단계 깊이 제한)
    if (currentPath.length < 5) {
      const subFolders = folder.getFolders();
      while (subFolders.hasNext()) {
        const subFolder = subFolders.next();
        const subFolderName = subFolder.getName();
        if (EXCLUDED_FOLDERS.includes(subFolderName)) {
          Logger.log(`[DriveManager] 제외된 폴더 스킵: ${subFolderName}`);
          continue;
        }
        traverseFolder(subFolder, [...currentPath, subFolderName]);
      }
    }
  };

  // 드라이브 루트부터 시작
  const rootFolder = DriveApp.getRootFolder();
  traverseFolder(rootFolder, []);

  // 남은 청크 플러시
  flushChunk();

  Logger.log(`[DriveManager] 파일 목록 생성 완료! 총 ${totalProcessed}개 파일.`);
  toastMessage(`✅ 탐색 완료! 총 ${totalProcessed}개의 파일 목록을 생성했습니다.`);
}

/**
 * 2. 일괄 삭제 기능
 * 대상 탭: [변경 대상] ('삭제필요' 열 == 1 인 항목 휴지통으로 이동)
 */
function deleteDriveFiles() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('변경 대상');

  if (!sheet) {
    ss.toast('⚠️ [변경 대상] 탭을 찾을 수 없습니다.', '일괄 삭제 오류');
    Logger.log('[DriveManager] [변경 대상] 탭이 존재하지 않습니다.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ss.toast('ℹ️ [변경 대상] 탭에 처리할 데이터가 없습니다.', '일괄 삭제');
    return;
  }

  // 17개 열 데이터를 가져옴 (12번째 열: 삭제필요, 6번째 열: 파일ID)
  const range = sheet.getRange(2, 1, lastRow - 1, 17);
  const values = range.getValues();

  let deletedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const fileId = String(row[5] || '').trim(); // Col 6: 파일ID
    const deleteFlag = String(row[11] || '').trim(); // Col 12: 삭제필요

    if (deleteFlag === '1' && fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        file.setTrashed(true);
        deletedCount++;
        // 시트 삭제필요 열 상태 표시 변경
        sheet.getRange(i + 2, 12).setValue('삭제완료');
      } catch (e) {
        Logger.log(`[DriveManager] 파일 삭제 실패 (ID: ${fileId}): ${e}`);
        sheet.getRange(i + 2, 12).setValue(`오류: ${e}`);
      }
    }
  }

  SpreadsheetApp.flush();
  ss.toast(`✅ 총 ${deletedCount}개 파일이 휴지통으로 이동되었습니다.`, '일괄 삭제 완료');
  Logger.log(`[DriveManager] 일괄 삭제 완료: ${deletedCount}건`);
}

/**
 * 목표 폴더 경로(LV1~LV5)를 가져오거나 없으면 생성하는 헬퍼 함수
 */
function getOrCreateFolderPath(folderNames: string[]): GoogleAppsScript.Drive.Folder {
  let currentFolder = DriveApp.getRootFolder();
  for (const name of folderNames) {
    if (!name || String(name).trim() === '') break;
    const cleanName = String(name).trim();
    const existingFolders = currentFolder.getFoldersByName(cleanName);
    if (existingFolders.hasNext()) {
      currentFolder = existingFolders.next();
    } else {
      currentFolder = currentFolder.createFolder(cleanName);
      Logger.log(`[DriveManager] 새 폴더 생성: ${cleanName}`);
    }
  }
  return currentFolder;
}

/**
 * 지정된 폴더 경로(레벨1~레벨5)를 루트부터 순차 탐색하여 폴더 객체를 찾는 헬퍼 함수
 * 경로 상에 폴더가 존재하지 않으면 null을 반환합니다.
 */
function findFolderByPath(folderNames: string[]): GoogleAppsScript.Drive.Folder | null {
  const validNames = folderNames.map(n => String(n || '').trim()).filter(n => n !== '');
  if (validNames.length === 0) return null;

  let currentFolder = DriveApp.getRootFolder();
  for (const name of validNames) {
    const subFolders = currentFolder.getFoldersByName(name);
    if (!subFolders.hasNext()) {
      Logger.log(`[DriveManager] 폴더를 찾을 수 없음: ${name} (경로: ${validNames.join('/')})`);
      return null;
    }
    currentFolder = subFolders.next();
  }
  return currentFolder;
}

/**
 * 3. 일괄 이동 기능
 * 대상 탭: [변경 대상] (LV1~LV5 지정 폴더로 이동)
 */
function moveDriveFiles() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('변경 대상');

  if (!sheet) {
    ss.toast('⚠️ [변경 대상] 탭을 찾을 수 없습니다.', '일괄 이동 오류');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ss.toast('ℹ️ [변경 대상] 탭에 처리할 데이터가 없습니다.', '일괄 이동');
    return;
  }

  // 17개 열 가져오기
  const range = sheet.getRange(2, 1, lastRow - 1, 17);
  const values = range.getValues();

  let movedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const fileId = String(row[5] || '').trim(); // Col 6: 파일ID

    // 현재 레벨 1~5 (Col 7~11)
    const currentLevels = [row[6], row[7], row[8], row[9], row[10]].map(v => String(v || '').trim());
    // 목표 레벨 LV1~LV5 (Col 13~17)
    const targetLevels = [row[12], row[13], row[14], row[15], row[16]].map(v => String(v || '').trim());

    // 목표 레벨 중 최소 LV1이 지정되어 있고 현재 레벨과 다른지 확인
    const hasTarget = targetLevels.some(l => l !== '');
    const isDifferent = targetLevels.join('/') !== currentLevels.join('/');

    if (fileId && hasTarget && isDifferent) {
      try {
        const file = DriveApp.getFileById(fileId);
        const targetFolder = getOrCreateFolderPath(targetLevels);

        file.moveTo(targetFolder);
        movedCount++;

        // 이동 완료 후 시트에 신규 레벨 기록 및 목표 레벨 비우기
        const validTargets = targetLevels.filter(l => l !== '');
        const newLevels = [
          validTargets[0] || '',
          validTargets[1] || '',
          validTargets[2] || '',
          validTargets[3] || '',
          validTargets[4] || ''
        ];

        // 레벨1~레벨5 (Col 7~11) 갱신
        sheet.getRange(i + 2, 7, 1, 5).setValues([newLevels]);
        // LV1~LV5 (Col 13~17) 초기화
        sheet.getRange(i + 2, 13, 1, 5).clearContent();

        Logger.log(`[DriveManager] 파일 이동 성공: ${file.getName()} -> ${targetLevels.join('/')}`);
      } catch (e) {
        Logger.log(`[DriveManager] 파일 이동 실패 (ID: ${fileId}): ${e}`);
        sheet.getRange(i + 2, 13).setValue(`오류: ${e}`);
      }
    }
  }

  SpreadsheetApp.flush();
  ss.toast(`✅ 총 ${movedCount}개 파일이 지정된 폴더로 이동되었습니다.`, '일괄 이동 완료');
  Logger.log(`[DriveManager] 일괄 이동 완료: ${movedCount}건`);
}

/**
 * 4. 폴더 일괄 변경 (삭제/이동/생성) 기능
 * 대상 탭: [폴더일괄변경]
 * - Col 1~5: 레벨1 ~ 레벨5 (현재 폴더 계층)
 * - Col 6~10: LV1 ~ LV5 (이동/생성할 폴더 계층)
 * - 삭제/생성 열 (Col 11): 
 *    - '1': 해당 폴더 계층 및 하위 파일/폴더 전체 삭제 (휴지통 이동)
 *    - '2': '생성' 모드. LV1~LV5 경로의 대상 폴더를 생성 (이미 존재하는 경우 스킵, 레벨1~레벨5 열 무시)
 *    - 그 외 ('1' 또는 '2'가 아닌 경우): '이동' 모드. 레벨1~레벨5 원본 폴더 내용물을 LV1~LV5 대상 폴더로 이동 후 원본 폴더 삭제
 */
function batchMoveFolders() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('폴더일괄변경');

  if (!sheet) {
    ss.toast('⚠️ [폴더일괄변경] 탭을 찾을 수 없습니다.', '폴더 일괄 변경 오류');
    Logger.log('[DriveManager] [폴더일괄변경] 탭이 존재하지 않습니다.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ss.toast('ℹ️ [폴더일괄변경] 탭에 처리할 데이터가 없습니다.', '폴더 일괄 변경');
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), 12);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

  // '삭제/생성' (또는 기존 '삭제여부') 열 인덱스 확인 (0-based)
  let flagColIdx = headers.indexOf('삭제/생성');
  if (flagColIdx === -1) {
    flagColIdx = headers.indexOf('삭제여부');
  }
  if (flagColIdx === -1) {
    flagColIdx = 10; // 기본 11번째 열 (Col 11)
  }

  // '처리 결과' 열 인덱스 확인 및 헤더 생성
  let resultColIdx = headers.findIndex(h => h.includes('결과') || h.includes('상태'));
  if (resultColIdx === -1) {
    resultColIdx = Math.max(flagColIdx + 1, 11);
    sheet.getRange(1, resultColIdx + 1).setValue('처리 결과').setFontWeight('bold').setBackground('#e6f2ff');
  }

  const totalCols = Math.max(lastCol, resultColIdx + 1);
  const range = sheet.getRange(2, 1, lastRow - 1, totalCols);
  const values = range.getValues();

  let movedFolderCount = 0;
  let createdFolderCount = 0;
  let deletedFolderCount = 0;
  let totalMovedFiles = 0;
  let totalMovedFolders = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    // 현재 레벨 1~5 (Col 1~5)
    const sourceLevels = [row[0], row[1], row[2], row[3], row[4]]
      .map(v => String(v || '').trim())
      .filter(v => v !== '');

    // 목표 레벨 LV1~LV5 (Col 6~10)
    const targetLevels = [row[5], row[6], row[7], row[8], row[9]]
      .map(v => String(v || '').trim())
      .filter(v => v !== '');

    // 삭제/생성 값 확인
    const flagVal = String(row[flagColIdx] || '').trim();

    // ==========================================
    // 1. [폴더 전체 삭제 모드] (삭제/생성 == '1')
    // ==========================================
    if (flagVal === '1') {
      if (sourceLevels.length === 0) {
        continue;
      }
      try {
        const sourceFolder = findFolderByPath(sourceLevels);
        if (!sourceFolder) {
          sheet.getRange(i + 2, resultColIdx + 1).setValue('오류: 원본 폴더를 찾을 수 없음');
          Logger.log(`[DriveManager] 삭제 대상 폴더 없음: ${sourceLevels.join('/')}`);
          continue;
        }

        // 폴더 자체를 휴지통으로 이동 (폴더 자신 및 하위 파일/서브폴더 전체 삭제)
        sourceFolder.setTrashed(true);
        deletedFolderCount++;

        // 시트 상태 갱신
        sheet.getRange(i + 2, flagColIdx + 1).setValue('삭제완료');
        sheet.getRange(i + 2, resultColIdx + 1).setValue('✅ 폴더 및 하위 항목 전체 삭제 완료');

        Logger.log(`[DriveManager] 폴더 일괄 삭제 성공: ${sourceLevels.join('/')}`);
      } catch (e) {
        Logger.log(`[DriveManager] 폴더 삭제 실패 (${sourceLevels.join('/')}): ${e}`);
        sheet.getRange(i + 2, resultColIdx + 1).setValue(`오류: ${e}`);
      }
      continue;
    }

    // ==========================================
    // 2. [폴더 생성 모드] (삭제/생성 == '2')
    // 레벨1~레벨5 열의 내용은 무시하고 LV1~LV5 대상 폴더 생성 (이미 존재 시 건너뛰기)
    // ==========================================
    if (flagVal === '2') {
      if (targetLevels.length === 0) {
        sheet.getRange(i + 2, resultColIdx + 1).setValue('오류: 생성 대상 경로(LV1~LV5) 없음');
        continue;
      }

      try {
        const existingFolder = findFolderByPath(targetLevels);
        if (existingFolder) {
          sheet.getRange(i + 2, resultColIdx + 1).setValue('이미 존재함 (스킵)');
          Logger.log(`[DriveManager] 폴더 생성 스킵 (이미 존재): ${targetLevels.join('/')}`);
          continue;
        }

        getOrCreateFolderPath(targetLevels);
        createdFolderCount++;

        sheet.getRange(i + 2, flagColIdx + 1).setValue('생성완료');
        sheet.getRange(i + 2, resultColIdx + 1).setValue('✅ 폴더 생성 완료');

        Logger.log(`[DriveManager] 폴더 생성 성공: ${targetLevels.join('/')}`);
      } catch (e) {
        Logger.log(`[DriveManager] 폴더 생성 실패 (${targetLevels.join('/')}): ${e}`);
        sheet.getRange(i + 2, resultColIdx + 1).setValue(`오류: ${e}`);
      }
      continue;
    }

    // ==========================================
    // 3. [폴더 내용물 이동 모드] (삭제/생성 != '1' && != '2')
    // ==========================================
    if (sourceLevels.length === 0 || targetLevels.length === 0) {
      continue;
    }

    // 동일 경로인 경우 스킵
    if (sourceLevels.join('/') === targetLevels.join('/')) {
      sheet.getRange(i + 2, resultColIdx + 1).setValue('동일 경로 (스킵)');
      continue;
    }

    try {
      // 1. 원본 폴더 탐색
      const sourceFolder = findFolderByPath(sourceLevels);
      if (!sourceFolder) {
        sheet.getRange(i + 2, resultColIdx + 1).setValue('오류: 원본 폴더를 찾을 수 없음');
        Logger.log(`[DriveManager] 원본 폴더 없음: ${sourceLevels.join('/')}`);
        continue;
      }

      // 2. 대상 폴더 탐색 또는 생성
      const targetFolder = getOrCreateFolderPath(targetLevels);

      // 원본 폴더와 대상 폴더가 동일한 경우 스킵
      if (sourceFolder.getId() === targetFolder.getId()) {
        sheet.getRange(i + 2, resultColIdx + 1).setValue('동일 폴더 ID (스킵)');
        continue;
      }

      // 3. 원본 폴더 내 직하위 파일들을 대상 폴더로 이동
      let fileCount = 0;
      const files = sourceFolder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        file.moveTo(targetFolder);
        fileCount++;
      }

      // 4. 원본 폴더 내 직하위 폴더들을 대상 폴더로 이동
      let folderCount = 0;
      const subFolders = sourceFolder.getFolders();
      while (subFolders.hasNext()) {
        const subFolder = subFolders.next();
        // 대상 폴더 자체나 동일 폴더 이동 방지
        if (subFolder.getId() === targetFolder.getId()) continue;
        subFolder.moveTo(targetFolder);
        folderCount++;
      }

      // 5. 내용물이 모두 이동된 원본 폴더 휴지통으로 이동 (삭제)
      sourceFolder.setTrashed(true);

      // 6. 시트 행 갱신
      const newLevels = [
        targetLevels[0] || '',
        targetLevels[1] || '',
        targetLevels[2] || '',
        targetLevels[3] || '',
        targetLevels[4] || ''
      ];
      // 레벨1~5 갱신
      sheet.getRange(i + 2, 1, 1, 5).setValues([newLevels]);
      // LV1~5 초기화
      sheet.getRange(i + 2, 6, 1, 5).clearContent();
      // 처리 결과 표기
      sheet.getRange(i + 2, resultColIdx + 1).setValue(`✅ 완료 (파일 ${fileCount}개, 폴더 ${folderCount}개 이동 후 원본 삭제)`);

      movedFolderCount++;
      totalMovedFiles += fileCount;
      totalMovedFolders += folderCount;

      Logger.log(`[DriveManager] 폴더 이동 성공: ${sourceLevels.join('/')} -> ${targetLevels.join('/')} (파일: ${fileCount}개, 폴더: ${folderCount}개)`);
    } catch (e) {
      Logger.log(`[DriveManager] 폴더 이동 실패 (${sourceLevels.join('/')} -> ${targetLevels.join('/')}): ${e}`);
      sheet.getRange(i + 2, resultColIdx + 1).setValue(`오류: ${e}`);
    }
  }

  SpreadsheetApp.flush();
  const summaryParts: string[] = [];
  if (createdFolderCount > 0) summaryParts.push(`폴더 생성 ${createdFolderCount}건`);
  if (movedFolderCount > 0) summaryParts.push(`폴더 이동 ${movedFolderCount}건`);
  if (deletedFolderCount > 0) summaryParts.push(`폴더 삭제 ${deletedFolderCount}건`);

  let resultMsg = '';
  if (summaryParts.length > 0) {
    resultMsg = `✅ ${summaryParts.join(', ')} 완료`;
  } else {
    resultMsg = `✅ 처리 완료 (변경 항목 없음 또는 스킵)`;
  }

  ss.toast(resultMsg, '폴더 일괄 변경 완료');
  Logger.log(`[DriveManager] ${resultMsg}`);
}

