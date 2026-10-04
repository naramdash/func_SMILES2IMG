Office.onReady(() => {
  // Initialize Excel ribbon commands here when needed.
});

/** Entry point for a ribbon ExecuteFunction command. */
function action(event: Office.AddinCommands.Event): void {
  // Add the command's Excel operations here.
  event.completed();
}

Office.actions.associate("action", action);
