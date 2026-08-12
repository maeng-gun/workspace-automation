/**
 * Google Docs 자동화 모듈
 */

/**
 * 신규 구글 문서를 생성하고 서식 있는 텍스트를 작성합니다.
 */
function createDocumentReport(title: string, sections: { heading: string; body: string }[]): GoogleAppsScript.Document.Document {
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  // 제목 추가
  const titlePara = body.appendParagraph(title);
  titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);

  // 섹션별 내용 추가
  sections.forEach((sec) => {
    const headingPara = body.appendParagraph(sec.heading);
    headingPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    body.appendParagraph(sec.body);
    body.appendParagraph(''); // 빈 줄
  });

  doc.saveAndClose();
  Logger.log(`[Docs] 문서 생성 완료: ${doc.getUrl()}`);
  return doc;
}
