/**
 * Google Tasks 자동화 모듈
 * (참고: Tasks 고급 서비스를 사용합니다)
 */

interface TaskParams {
  title: string;
  notes?: string;
  dueDate?: Date;
}

/**
 * 기본 할 일 목록에 새 Task를 추가합니다.
 */
function createNewTask(params: TaskParams) {
  // Tasks 고급 서비스가 활성화되어 있어야 합니다 (appsscript.json 확인)
  if (typeof Tasks === 'undefined') {
    Logger.log('[Tasks] Tasks 고급 서비스가 활성화되어 있지 않습니다.');
    return null;
  }

  const taskLists = Tasks.Tasklists?.list();
  const defaultList = taskLists?.items?.[0];

  if (!defaultList || !defaultList.id) {
    Logger.log('[Tasks] 기본 할 일 목록을 찾을 수 없습니다.');
    return null;
  }

  const newTask: GoogleAppsScript.Tasks.Schema.Task = {
    title: params.title,
    notes: params.notes || '',
  };

  if (params.dueDate) {
    newTask.due = params.dueDate.toISOString();
  }

  const createdTask = Tasks.Tasks?.insert(newTask, defaultList.id);
  Logger.log(`[Tasks] 새 할 일 추가 완료: ${createdTask?.title}`);
  return createdTask;
}
