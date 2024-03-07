/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const addinVersion = "1.0";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log(`AddIn - V${addinVersion}`);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run1").onclick = run1;
    document.getElementById("run2").onclick = run2;
  }
});

export async function run1() {
  console.log(`Run V1 - (V${addinVersion})`);
  return await Word.run(async (context) => {
    let start = performance.now();

    document.getElementById("progressbox").style.display = "block";
    document.getElementById("progressbar").style.width = "0%";

    const paragraphs = context.document.body.paragraphs;

    // load text
    //paragraphs.load("$all");
    paragraphs.load(["text", "items"]);
    await context.sync();

    document.getElementById("progressbar").style.width = "10%";

    let end = performance.now();
    console.log(`Execution time load paragraphs: ${end - start} ms`);
    start = performance.now();

    const wordsRangeCollections = [];

    for (let i = 0; i < paragraphs.items.length; i++) {
      const paragraph = paragraphs.items[i];

      // only process if text available
      if (paragraph.text) {
        const wordsRangeCollection = paragraph.getRange().split([" "]);

        wordsRangeCollection.load("$none");
        wordsRangeCollections.push(wordsRangeCollection);
      }
    }

    try {
      await context.sync();
    } catch (error) {
      console.error("error on sync2 - " + error);
    }

    end = performance.now();
    console.log(`Execution time load wordsRangeCollections: ${end - start} ms`);
    start = performance.now();

    const wordChars = [];

    for (let i = 0; i < wordsRangeCollections.length; i++) {
      const wordsRangeCollection = wordsRangeCollections[i];

      for (let j = 0; j < wordsRangeCollection.items.length; j++) {
        const wordRange = wordsRangeCollection.items[j];

        const wordChar = wordRange.getRange().split([""]);

        wordChar.load("$none");
        wordChars.push(wordChar);
        //wordChar.untrack();
      }
    }

    document.getElementById("progressbar").style.width = "30%";

    end = performance.now();
    console.log(`Execution time load wordChars: ${end - start} ms`);
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
        if (wordChar) {
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

    end = performance.now();
    console.log(`Execution time update word formatting: ${end - start} ms`);
    start = performance.now();

    await context.sync();

    end = performance.now();
    console.log(`Execution time context sync: ${end - start} ms`);

    document.getElementById("progressbar").style.width = "100%";
    document.getElementById("progressbox").style.display = "none";
  });
}

export async function run2() {
  console.log(`Run V2 - (V${addinVersion})`);
  return await Word.run(async (context) => {
    const start = performance.now();

    document.getElementById("progressbox").style.display = "block";
    document.getElementById("progressbar").style.width = "0%";

    const paragraphs = context.document.body.paragraphs;

    // load text
    //paragraphs.load("$all");
    paragraphs.load(["text", "items"]);
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

          if (wordChar) {
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
