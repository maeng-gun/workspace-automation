"use strict";
/**
 * Google Tasks 자동화 모듈
 * (참고: Tasks 고급 서비스를 사용합니다)
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.createNewTask = createNewTask;
/**
 * 기본 할 일 목록에 새 Task를 추가합니다.
 */
function createNewTask(params) {
    var _a, _b, _c;
    // Tasks 고급 서비스가 활성화되어 있어야 합니다 (appsscript.json 확인)
    if (typeof Tasks === 'undefined') {
        Logger.log('[Tasks] Tasks 고급 서비스가 활성화되어 있지 않습니다.');
        return null;
    }
    const taskLists = (_a = Tasks.Tasklists) === null || _a === void 0 ? void 0 : _a.list();
    const defaultList = (_b = taskLists === null || taskLists === void 0 ? void 0 : taskLists.items) === null || _b === void 0 ? void 0 : _b[0];
    if (!defaultList || !defaultList.id) {
        Logger.log('[Tasks] 기본 할 일 목록을 찾을 수 없습니다.');
        return null;
    }
    const newTask = {
        title: params.title,
        notes: params.notes || '',
    };
    if (params.dueDate) {
        newTask.due = params.dueDate.toISOString();
    }
    const createdTask = (_c = Tasks.Tasks) === null || _c === void 0 ? void 0 : _c.insert(newTask, defaultList.id);
    Logger.log(`[Tasks] 새 할 일 추가 완료: ${createdTask === null || createdTask === void 0 ? void 0 : createdTask.title}`);
    return createdTask;
}
