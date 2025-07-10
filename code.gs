let output_message = ""

function createAndConfigureGroupsV2() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Groups");
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("API作成ログ");
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const lastRow = logSheet.getLastRow() + 1;
    const data = sheet.getDataRange().getValues();
    const successRows = []; // 成功した行番号（1ベース）を記録
    sheet.getRange("H:H").clearContent();
  
    for (let i = 1; i < data.length; i++) {
      const [account, email, name, description, manager, directoryVisible] = data[i];
      const logRow = 7;
      const logCol = 7;

      if (!email || !email.includes("@")) {
        continue;
      }
  
      const visibleFlag = directoryVisible !== "〇";
      let visibleresult = "Visible";
      if (!visibleFlag) {
        visibleresult = "Not Visible";
      }
  
      try {
        // グループ作成
        AdminDirectory.Groups.insert({
          email: email,
          name: name,
          description: description
        });
        output_(`グループ作成成功: ${email}`);

        logSheet.appendRow([email, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")]);
  
        // 設定反映
        GroupsSettings.Groups.patch({
          whoCanJoin: "INVITED_CAN_JOIN",
          whoCanViewGroup: "ALL_MEMBERS_CAN_VIEW",
          whoCanViewMembership: "ALL_IN_DOMAIN_CAN_VIEW",
          allowExternalMembers: "false",
          whoCanPostMessage: "ANYONE_CAN_POST",
          isArchived: true,
          archiveOnly: false,
          membersCanPostAsTheGroup: true,
          messageModerationLevel: "MODERATE_NONE",
          spamModerationLevel: "ALLOW",
          sendMessageDenyNotification: false,
          showInGroupDirectory: visibleFlag,
          includeInGlobalAddressList: visibleFlag,
          enableCollaborativeInbox: false,
          primaryLanguage: "ja"
        }, email);
        output_(`設定反映成功: ${email}`);
        output_(`アドレス帳: ${visibleresult}`)
  
        // マネージャー追加
        if (manager && manager.includes("@")) {
          AdminDirectory.Members.insert({
            email: manager,
            role: "MANAGER"
          }, email);
          output_(`マネージャー追加成功: ${manager}`);
        }
  
        // 削除予定として記録（1ベース）
        successRows.push(i + 1);
  
      } catch (e) {
        output_(`エラー（${email}）: ${e.message}`, "red");
      }
    }
  
    // 成功した行だけを、後ろから順に削除（上から削除すると行ズレが起きるため）
    successRows.forEach(row => {
      // A列（1）、C列（3）〜E列（5）、F列（6）のセル内容を削除
      sheet.getRange(row, 1).clearContent(); // A列
      sheet.getRange(row, 3, 1, 3).clearContent(); // C〜E列（3列）
      sheet.getRange(row, 6).clearContent(); // F列
    });

  }

function appendLogToCell(sheet, row, column, newLog) {
  const cell = sheet.getRange(row, column);
  const existingLog = cell.getValue();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const entry = `[${timestamp}] ${newLog}`;

  const updatedLog = existingLog ? `${entry}\n${existingLog}` : entry;
  cell.setValue(updatedLog);
}

function output_(message, bgcolor="white") {
    const inputsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Groups");
    const lastRow = inputsheet.getRange("H:H").getValues().filter(String).length + 1;

    Logger.log(message);

    output_message = (Utilities.formatDate( new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")+" "+message);
    inputsheet.getRange(`H${lastRow}`).setBackground(bgcolor).setValue(output_message);
}
  
