/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, console */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("open-dialog").onclick = openDialog;
    document.getElementById("message-dialog").onclick = messageDialog;
  }
});

let dialog: Office.Dialog;
function openDialog() {
  Office.context.ui.displayDialogAsync('https://localhost:3000/dialog.html', {
    width: 30,
    height: 80
  }, asyncResult => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(asyncResult.error);
    } else {
      dialog = asyncResult.value;
      console.log(dialog);

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        dialog?.close();
        dialog = null;
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        dialog?.close();
        dialog = null;
      });
    }
  });
}

function messageDialog() {
  //@ts-ignore
  dialog?.messageChild('Wakka, wakka!');
}
