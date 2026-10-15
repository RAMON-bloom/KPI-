/**
 * 面接振り返りフォーム（候補者向け）を作成する Google Apps Script。
 * script.google.com で新規プロジェクトに貼り付けて createInterviewFeedbackForm を1回実行する。
 * 実行ログに、フォームURL・回答スプレッドシートID・事前入力用のエントリーIDが出るので、
 * kpi-mgr の services/interviewFeedbackForm.ts の定数に転記する。
 */
function createInterviewFeedbackForm() {
  var form = FormApp.create('面接振り返りアンケート（bloom）');
  form.setDescription(
    '面接お疲れさまでした。今後の選考対策のため、面接の内容を教えてください（5分ほどで回答できます）。\n' +
    'いただいた回答は、担当エージェントがあなたの次の選考対策に使うほか、個人が特定できない形に' +
    '加工したうえで、bloom社内の面接対策ナレッジとして活用させていただきます。'
  );
  // 社外（候補者）が回答するため、メール収集は求めない。setRequireLogin は現在サポートされて
  // いないため、ログイン不要（回答者:リンクを知っている全員）はフォームの設定画面で行う。
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);
  form.setConfirmationMessage('ご回答ありがとうございました。担当エージェントより改めてご連絡いたします。');

  // --- 事前入力される選考情報（担当者が送るリンクに最初から入っている） ---
  form.addSectionHeaderItem()
    .setTitle('選考情報')
    .setHelpText('担当者が入力済みです。変更しないでください。');
  var appId = form.addTextItem().setTitle('受付番号').setRequired(true);
  var company = form.addTextItem().setTitle('企業名').setRequired(true);
  var phase = form.addTextItem().setTitle('選考フェーズ').setRequired(true);
  form.addDateItem().setTitle('面接日').setRequired(false);

  // --- 面接の内容 ---
  form.addPageBreakItem().setTitle('面接の内容');
  form.addMultipleChoiceItem().setTitle('面接の形式')
    .setChoiceValues(['対面', 'オンライン', '電話']).showOtherOption(true).setRequired(false);
  form.addMultipleChoiceItem().setTitle('所要時間')
    .setChoiceValues(['30分未満', '30〜45分', '45〜60分', '60〜90分', '90分以上']).setRequired(false);
  form.addTextItem().setTitle('面接官の人数・役職')
    .setHelpText('例：事業部長1名、人事1名（お名前は不要です）').setRequired(false);
  form.addParagraphTextItem().setTitle('聞かれた質問')
    .setHelpText('覚えている範囲で、箇条書きで構いません。').setRequired(true);
  form.addParagraphTextItem().setTitle('うまく答えられた点').setRequired(false);
  form.addParagraphTextItem().setTitle('答えにくかった点・準備不足だと感じた点').setRequired(false);
  form.addParagraphTextItem().setTitle('面接官の反応・雰囲気').setRequired(false);
  form.addParagraphTextItem().setTitle('逆質問とその回答').setRequired(false);

  // --- 任意：手応えや志望度（担当者が今後の進め方を考える参考にする） ---
  form.addPageBreakItem().setTitle('差し支えなければ教えてください（任意）')
    .setHelpText('今後のサポートの参考にします。答えにくい項目は空欄のままで構いません。');
  form.addScaleItem().setTitle('面接の手応え').setBounds(1, 5).setLabels('手応えなし', '手応えあり').setRequired(false);
  form.addScaleItem().setTitle('この企業への志望度').setBounds(1, 5).setLabels('低い', '高い').setRequired(false);
  form.addMultipleChoiceItem().setTitle('面接を受けて、志望度は変わりましたか')
    .setChoiceValues(['上がった', '変わらない', '下がった']).setRequired(false);
  form.addParagraphTextItem().setTitle('他社の選考状況（内定・最終面接の予定など）').setRequired(false);
  form.addParagraphTextItem().setTitle('担当者に伝えたいこと・不安な点').setRequired(false);

  // 回答の保存先スプレッドシート
  var ss = SpreadsheetApp.create('面接振り返りアンケート（回答）');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // 事前入力URLからエントリーIDを取り出す
  var prefilled = form.createResponse()
    .withItemResponse(appId.createResponse('APPID'))
    .withItemResponse(company.createResponse('COMPANY'))
    .withItemResponse(phase.createResponse('PHASE'))
    .toPrefilledUrl();
  var ids = {};
  prefilled.replace(/entry\.(\d+)=([A-Z]+)/g, function (_, id, key) { ids[key] = id; return _; });

  Logger.log('FORM_EDIT_URL=' + form.getEditUrl());
  Logger.log('FORM_RESPONDER_URL=' + form.getPublishedUrl());
  Logger.log('RESPONSE_SPREADSHEET_ID=' + ss.getId());
  Logger.log('ENTRY_APPLICATION_ID=' + ids.APPID);
  Logger.log('ENTRY_COMPANY=' + ids.COMPANY);
  Logger.log('ENTRY_PHASE=' + ids.PHASE);
  Logger.log('PREFILLED_SAMPLE=' + prefilled);
}
