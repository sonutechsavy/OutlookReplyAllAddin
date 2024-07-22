Office.initialize = function (reason) {
    $(document).ready(function () {
        Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
    });
};

function itemChanged(eventArgs) {
    if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
        var message = Office.context.mailbox.item;
        if (message.conversationId && message.conversationId.replyAllRecipients.length > 1) {
            Office.context.mailbox.displayNewMessageForm({
                toRecipients: [],
                subject: "Warning: Reply All",
                body: "Are you sure you want to reply all?"
            });
        }
    }
}
