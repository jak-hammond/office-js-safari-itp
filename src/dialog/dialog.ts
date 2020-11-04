// eslint-disable-next-line
/* global document, Office, console */

Office.onReady(() => {
   console.log(Office.context);
   // @ts-ignore
   Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, messageHandler);
});

const button = document.getElementById('cancel-send');
// @ts-ignore
button.addEventListener('click', () => {
   Office.context.ui.messageParent('cancel-send');
});

function messageHandler(payload: string) {
   console.info('Message received from dialog parent:', payload);
   //@ts-ignore
   document.getElementById('dialog-messages')?.innerHTML += `<br/>Message:[${payload.message}]. Type:[${payload.type}].`;
}
