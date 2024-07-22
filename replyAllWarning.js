Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      // Add any initialization logic to be run once the host is ready
    }
  });
  
  function showReplyAllWarning(event) {
    const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      if (item.toRecipients.length > 1) {
        Office.context.mailbox.displayNewMessageForm({
          toRecipients: [],
          subject: "Warning: Reply All",
          body: "Are you sure you want to reply all?"
        });
      }
    }
    event.completed();
  }
  