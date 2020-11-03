// eslint-disable-next-line
/* global document, Office, localStorage */

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

const value = localStorage.getItem('jh-test');
const span = document.getElementById('storage-value');
span.innerText = value;


const button = document.getElementById('cancel-send');
// @ts-ignore
button.addEventListener('click', () => {
   Office.context.ui.messageParent('cancel-send');
});
