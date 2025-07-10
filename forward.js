Office.initialize = () => {
  Office.actions.associate("forwardToLog", forwardToLog);
};

function forwardToLog(event) {
  const item = Office.context.mailbox.item;

  item.forwardAsync(
    {
      toRecipients: ["log@teachforaustralia.org"],
    },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Email forwarded.");
      } else {
        console.error("Error forwarding: ", result.error);
      }
      event.completed();
    }
  );
}
