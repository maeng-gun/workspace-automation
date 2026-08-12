"use strict";
/**
 * Google Drive 자동화 모듈
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.getOrCreateFolder = getOrCreateFolder;
exports.listFilesByFolder = listFilesByFolder;
/**
 * 구글 드라이브 루트 또는 특정 경로에 폴더를 생성하거나 조회합니다.
 */
function getOrCreateFolder(folderName) {
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
        const existingFolder = folders.next();
        Logger.log(`[Drive] 기존 폴더 사용: ${existingFolder.getName()} (${existingFolder.getId()})`);
        return existingFolder;
    }
    const newFolder = DriveApp.createFolder(folderName);
    Logger.log(`[Drive] 새 폴더 생성: ${newFolder.getName()} (${newFolder.getId()})`);
    return newFolder;
}
/**
 * 특정 폴더 안의 파일 목록을 로깅합니다.
 */
function listFilesByFolder(folderName) {
    const folder = getOrCreateFolder(folderName);
    const files = folder.getFiles();
    const fileList = [];
    while (files.hasNext()) {
        const file = files.next();
        fileList.push(file.getName());
        Logger.log(`[Drive] 파일 발견: ${file.getName()} (${file.getSize()} bytes)`);
    }
    return fileList;
}
