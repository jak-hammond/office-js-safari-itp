import getGlobal from '../getGlobal';
/* eslint-disable */
/* global , Office */

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

let sendEvent: Office.AddinCommands.Event;
let dialog: Office.Dialog;

function processOnSendEvent(event: Office.AddinCommands.Event) {
    console.info('Pausing send event');
    sendEvent = event;

    setTimeout(() => {
        console.info('Setting [jh-test] localStorage value.');
        localStorage.setItem('jh-test', 'Hello, world!');

        displayDialog();
    }, 5000);
}

function displayDialog() {
    Office.context.ui.displayDialogAsync('https://localhost:3000/dialog.html', {
      width: 30,
      height: 80
    }, asyncResult => {
       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
           console.error(asyncResult.error);
           sendEvent.completed({ allowEvent: false });
       } else {
           dialog = asyncResult.value;

           dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
               sendEvent.completed({ allowEvent: false });
           });

           dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
               sendEvent.completed({ allowEvent: false });
           });
       }
    });
}

let g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.processOnSendEvent = processOnSendEvent;
