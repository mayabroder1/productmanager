/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    document.getElementById("run").onclick = run;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    
  }
});

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const range = context.document.getSelection();

    // change the paragraph font to bold.
    range.font.bold = true;

    range.load("text");

    await context.sync();
    localStorage.setItem('userSelectionText', 'range.text');
    //console.log(`the selected text was "${range.text}".`);
    Office.context.ui.displayDialogAsync(
      `https://localhost:3000/src/submitForm/submitForm.html?userText=${range.text}`, 
      { height: 200, width: 30, displayInIframe: true }
      );

    document.querySelector('h3').textContent =`the selected1 text was "${range.text}".`;

    
  });
}





