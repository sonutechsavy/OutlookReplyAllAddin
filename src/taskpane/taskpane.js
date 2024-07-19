/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    // Add the ItemSend event handler
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, onItemSend);
  }
});

async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}

function onItemSend(eventArgs) {
  let item = Office.context.mailbox.item;
  if (item) {
    // Check if there are multiple recipients in the To or CC fields
    if (item.to.length > 1 || item.cc.length > 0) {
      let message = "You are replying to all. Do you want to proceed?";
      Office.context.ui.displayDialogAsync('https://sonutechsavy.github.io/OutlookReplyAllAddin/src/dialog.html', { height: 30, width: 20 }, function (asyncResult) {
        let dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
      });
    } else {
      eventArgs.completed({ allowEvent: true });
    }
  } else {
    eventArgs.completed({ allowEvent: true });
  }
}

function messageHandler(arg) {
  if (arg.message === 'ok') {
    // Continue with the send operation
    eventArgs.completed({ allowEvent: true });
  } else {
    // Cancel the send operation
    eventArgs.completed({ allowEvent: false });
  }
}
