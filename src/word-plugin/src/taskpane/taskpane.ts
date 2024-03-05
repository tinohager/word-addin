/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run1").onclick = run1;
    document.getElementById("run2").onclick = run2;
  }
});

export async function run1() {
  return await Word.run(async (context) => {
    let start = performance.now();

    document.getElementById("progressbox").style.display = "block";
    document.getElementById("progressbar").style.width = "0%";
    
    const paragraphs = context.document.body.paragraphs;

    // load text
    //paragraphs.load("$all");
    paragraphs.load("text");
    await context.sync();

    document.getElementById("progressbar").style.width = "10%";

    const end2 = performance.now();
    console.log(`Execution time1: ${end2 - start} ms`);
    start = performance.now();

    const wordChars = [];

    for (let i = 0; i < paragraphs.items.length; i++) {
      const paragraph = paragraphs.items[i];

      // only process if text available
      if (paragraph.text) {
        const wordsRangeCollection = paragraph.getRange().split([" "]);

        wordsRangeCollection.load("$none");

        try {
          await wordsRangeCollection.context.sync();
        } catch (error) {
          console.error("error on sync1 - " + error);
          continue;
        }

        for (let j = 0; j < wordsRangeCollection.items.length; j++) {
          const wordRange = wordsRangeCollection.items[j];

          const wordChar = wordRange.getRange().split([""]);

          //wordChar.load("font");
          wordChar.load("$none");

          wordChars.push(wordChar);

          wordChar.untrack();
        }
      }
    }

    document.getElementById("progressbar").style.width = "30%";

    const end1 = performance.now();
    console.log(`Execution time2: ${end1 - start} ms`);
    start = performance.now();

    document.getElementById("progressbar").style.width = "40%";

    try {
      await context.sync();
    } catch (error) {
      console.error("error on sync2 - " + error);
    }

    document.getElementById("progressbar").style.width = "50%";

    for (let j = 0; j < wordChars.length; j++) {
      const wordChar = wordChars[j];
      try {
        if (wordChar && !wordChar.isNullObject) {
          if (wordChar.items.length > 2) {
            wordChar.items[1].font.bold = true;
            wordChar.items[2].font.bold = true;
          }
        }
      } catch (error) {
        console.error("error on process - " + error);
      }
    }

    // Synchronize the document state.

    document.getElementById("progressbar").style.width = "60%";

    const end = performance.now();
    console.log(`Execution time3: ${end - start} ms`);
    start = performance.now();

    await context.sync();

    const end4 = performance.now();
    console.log(`Execution time4: ${end4 - start} ms`);

    document.getElementById("progressbar").style.width = "100%";
    document.getElementById("progressbox").style.display = "none";
  });
}

export async function run2() {
  return await Word.run(async (context) => {
    const start = performance.now();

    document.getElementById("progressbox").style.display = "block";
    document.getElementById("progressbar").style.width = "0%";

    const paragraphs = context.document.body.paragraphs;

    // load text
    //paragraphs.load("$all");
    paragraphs.load("text");
    await context.sync();

    document.getElementById("progressbar").style.width = "10%";

    for (let i = 0; i < paragraphs.items.length; i++) {
      const paragraph = paragraphs.items[i];

      const percentage = (100 * i) / paragraphs.items.length;
      document.getElementById("progressbar").style.width = percentage + "%";

      // only process if text available
      if (paragraph.text) {
        const wordsRangeCollection = paragraph.getRange().split([" "]);
        wordsRangeCollection.load("$none");

        await wordsRangeCollection.context.sync();

        const wordChars = [];
        
        for (let j = 0; j < wordsRangeCollection.items.length; j++) {
          const wordRange = wordsRangeCollection.items[j];

          const wordChar = wordRange.getRange().split([""]);

          //wordChar.load("font");
          wordChar.load("$none");

          wordChars.push(wordChar);
        }

        try {
          await context.sync();
        } catch (error) {
          console.error("error on sync - " + error);
          continue;
        }

        for (let j = 0; j < wordChars.length; j++) {
          const wordChar = wordChars[j];

          if (wordChar && !wordChar.isNullObject) {
            if (wordChar.items.length > 2) {
              wordChar.items[1].font.bold = true;
              wordChar.items[2].font.bold = true;
            }
          }
        }
      }
    }

    // Synchronize the document state.
    await context.sync();

    const end = performance.now();
    console.log(`Execution time: ${end - start} ms`);

    document.getElementById("progressbar").style.width = "100%";
    document.getElementById("progressbox").style.display = "none";
  });
}
