"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const calendar_1 = require("./calendar");
const tasks_1 = require("./tasks");
const drive_1 = require("./drive");
const sheets_1 = require("./sheets");
const docs_1 = require("./docs");
/**
 * 1. 구글 캘린더 테스트 함수
 * GAS 에디터의 실행 메뉴에서 직접 실행할 수 있습니다.
 */
function testCalendar() {
    Logger.log('=== Google Calendar 테스트 시작 ===');
    (0, calendar_1.getTodayEvents)();
    const now = new Date();
    const startTime = new Date(now.getTime() + 60 * 60 * 1000); // 1시간 뒤
    const endTime = new Date(now.getTime() + 2 * 60 * 60 * 1000); // 2시간 뒤
    (0, calendar_1.createCalendarEvent)({
        title: '[자동화 샘플] 워크스페이스 동기화 회의',
        startTime,
        endTime,
        description: 'GAS를 이용해 자동 생성된 일정입니다.',
        location: 'Google Meet',
    });
}
/**
 * 2. Google Tasks 테스트 함수
 */
function testTasks() {
    Logger.log('=== Google Tasks 테스트 시작 ===');
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    (0, tasks_1.createNewTask)({
        title: '[자동화 샘플] 워크스페이스 자동화 모듈 점검',
        notes: 'Calendar, Drive, Sheets 연동 상태 확인',
        dueDate: tomorrow,
    });
}
/**
 * 3. 구글 드라이브 테스트 함수
 */
function testDrive() {
    Logger.log('=== Google Drive 테스트 시작 ===');
    const folder = (0, drive_1.getOrCreateFolder)('Workspace_Automation_Logs');
    (0, drive_1.listFilesByFolder)('Workspace_Automation_Logs');
}
/**
 * 4. 구글 시트 테스트 함수
 */
function testSheets() {
    Logger.log('=== Google Sheets 테스트 시작 ===');
    const headers = [['일자', '카테고리', '작업 내용', '상태']];
    const rows = [
        [new Date().toLocaleDateString(), 'Calendar', '일정 등록 자동화', '완료'],
        [new Date().toLocaleDateString(), 'Drive', '자동화 로그 폴더 생성', '진행중'],
        [new Date().toLocaleDateString(), 'Sheets', '데이터 리포트 생성', '대기중'],
    ];
    (0, sheets_1.createSpreadsheetWithData)('Workspace 자동화 리포트', headers, rows);
}
/**
 * 5. 구글 문서 테스트 함수
 */
function testDocs() {
    Logger.log('=== Google Docs 테스트 시작 ===');
    const sections = [
        {
            heading: '1. 개요',
            body: '본 문서는 Google Apps Script를 활용하여 생성된 일일 자동화 리포트입니다.',
        },
        {
            heading: '2. 실행 항목',
            body: '캘린더 일정 조회, Tasks 할 일 등록, 드라이브 폴더 생성, 시트 데이터 저장이 정상 처리되었습니다.',
        },
    ];
    (0, docs_1.createDocumentReport)('Workspace 자동화 일일보고서', sections);
}
/**
 * 6. 전체 워크스페이스 종합 실행 함수
 */
function runDailyAutomation() {
    Logger.log('🚀 [Workspace Automation] 전체 자동화 시작');
    testCalendar();
    testTasks();
    testDrive();
    testSheets();
    testDocs();
    Logger.log('✅ [Workspace Automation] 전체 자동화 완료');
}
