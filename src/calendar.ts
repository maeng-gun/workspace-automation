/**
 * Google Calendar 자동화 모듈
 */

interface CalendarEventParams {
  title: string;
  startTime: Date;
  endTime: Date;
  description?: string;
  location?: string;
}

/**
 * 캘린더에 새로운 일정을 등록합니다.
 */
function createCalendarEvent(params: CalendarEventParams): GoogleAppsScript.Calendar.CalendarEvent {
  const calendar = CalendarApp.getDefaultCalendar();
  const event = calendar.createEvent(params.title, params.startTime, params.endTime, {
    description: params.description || '',
    location: params.location || '',
  });
  Logger.log(`[Calendar] 일정 생성 완료: ${event.getTitle()} (${event.getId()})`);
  return event;
}

/**
 * 오늘 예정된 캘린더 일정 목록을 조회합니다.
 */
function getTodayEvents(): GoogleAppsScript.Calendar.CalendarEvent[] {
  const calendar = CalendarApp.getDefaultCalendar();
  const today = new Date();
  const events = calendar.getEventsForDay(today);
  
  Logger.log(`[Calendar] 오늘 일정 개수: ${events.length}개`);
  events.forEach((evt, idx) => {
    Logger.log(`  ${idx + 1}. [${evt.getStartTime().toLocaleTimeString()}] ${evt.getTitle()}`);
  });
  
  return events;
}
