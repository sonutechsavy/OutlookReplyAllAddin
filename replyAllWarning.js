Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      // Add any initialization logic to be run once the host is ready
    }
  });
  
  async function onMessageSendHandler(event) {
    const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      const subject = item.subject;
      if (subject.startsWith('RE:') || subject.startsWith('Re:')) {
        if (item.toRecipients.length > 1 || item.ccRecipients.length > 1 || item.bccRecipients.length > 1) {
          const userResponse = confirm("Are you sure you want to reply all?");
          if (!userResponse) {
            event.completed({ allowEvent: false });
            return;
          }
        }
      }
    }
    event.completed({ allowEvent: true });
  }
  