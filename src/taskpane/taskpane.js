Office.initialize = function (reason) {
  $(document).ready(function () {
      if (Office.context.mailbox.item) {
          addReplyAllWarning();
      }
  });
};

function addReplyAllWarning() {
  Office.context.mailbox.item.addHandlerAsync(
      Office.EventType.ItemSend,
      function (eventArgs) {
          var recipients = Office.context.mailbox.item.to.concat(
              Office.context.mailbox.item.cc
          );

          if (recipients.length > 2) {
              Office.context.mailbox.item.notificationMessages.addAsync(
                  'replyAllWarning',
                  {
                      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                      message: 'Warning: You are about to reply all to more than 2 recipients.',
                      icon: 'icon16',
                      persistent: true
                  }
              );

              eventArgs.completed({ allowEvent: false });
          } else {
              eventArgs.completed({ allowEvent: true });
          }
      }
  );
}
