/* exported Group */
class Group {
  constructor(settings) {
    this.settings = settings;
  }

  get groupMembers() {
    const group = GroupsApp.getGroupByEmail(this.settings.APP_GROUP_EMAIL);
    return (group ? group.getUsers() : []).map((member) => member.toString());
  }

  get bookGroupData() {
    if (!this._bookGroupData) {
      this._bookGroupData = SpreadsheetApp.openById(this.settings.APP_STORAGE_GROUP_DATA_ID);
    }
    return this._bookGroupData;
  }

  get storageGroupMembers() {
    return this.bookGroupData
      .getRange('Участники!A2:A')
      .getValues()
      .map((row) => row[0]);
  }

  checkNewMembers() {
    const membersFromGroup = this.groupMembers;
    const membersFromStorage = this.storageGroupMembers;
    const newMembers = membersFromGroup.filter((member) => !membersFromStorage.includes(member));
    if (newMembers.length) {
      const sheet = this.bookGroupData.getSheetByName('Участники');
      if (sheet) {
        const lastRow = sheet.getLastRow();
        const maxRow = sheet.getMaxRows();
        if (lastRow === maxRow) {
          sheet.appendRow(['']);
        }
        const range = sheet.getRange(lastRow + 1, 1, newMembers.length, 2);
        range.setValues(newMembers.map((member) => [member, new Date()]));
      }
    }
  }
}
