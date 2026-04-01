/* global Office */

Office.onReady(() => {
  Office.actions.associate("ShowTaskpane", showTaskpane);
});

function showTaskpane(): void {
  // 由 manifest 的 ShowTaskpane Action 控制显示任务窗格。
}

