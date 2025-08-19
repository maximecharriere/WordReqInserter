/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, Word */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Inserts a new requirement with incremental number format [REQ_XXXX]
 * @param event {Office.AddinCommands.Event}
 */
async function InsertRequirement(event) {
  console.log("=== Starting requirement insertion ===");

  try {
    await Word.run(async (context) => {
      // Get the entire document text to search for existing requirements
      console.log("Loading document content...");
      const documentBody = context.document.body;
      context.load(documentBody, "text");

      await context.sync();

      // Find all existing requirements using regex
      const regex = /\[REQ_(\d{4})\]/g;
      const matches = documentBody.text.match(regex);

      let nextReqNumber = 1; // Default value if no requirements found

      if (matches) {
        console.log(`Found ${matches.length} existing requirements`);
        // Extract numbers and find the highest one
        const reqNumbers = matches.map(match => {
          const numberMatch = match.match(/\[REQ_(\d{4})\]/);
          return parseInt(numberMatch[1], 10);
        });

        const maxReqNumber = Math.max(...reqNumbers);
        nextReqNumber = maxReqNumber + 1;
        console.log(`Next requirement number will be: ${nextReqNumber}`);
      } else {
        console.log("No existing requirements found, starting with REQ_0001");
      }

      // Format the requirement header
      const reqHeader = `[REQ_${nextReqNumber.toString().padStart(4, '0')}]`;
      const bookmarkName = `REQ_${nextReqNumber.toString().padStart(4, '0')}`;
      console.log(`Inserting requirement: ${reqHeader}`);

      // Insert the requirement header at the current selection
      const selection = context.document.getSelection();
      const insertedRange = selection.insertText(reqHeader, Word.InsertLocation.replace);
      context.load(insertedRange);

      await context.sync();
      console.log("Requirement header inserted successfully");

      // Add bookmark to the inserted requirement
      try {
        insertedRange.insertBookmark(bookmarkName);
        console.log(`Bookmark '${bookmarkName}' added successfully`);
      } catch (bookmarkError) {
        console.error("Error adding bookmark:", bookmarkError);
      }

      // Apply style if it exists
      try {
        insertedRange.style = "REQ_TITLE";
        console.log("Style 'REQ_TITLE' applied successfully");
        await context.sync();
      } catch (styleError) {
        if (styleError.code === 'InvalidArgument') {
          console.log("Style 'REQ_TITLE' not found, continuing without styling");
        } else {
          console.error("Error applying style:", styleError.message);
        }
      }

      // Move cursor to the end of the inserted text
      insertedRange.select("End");
      console.log("Cursor positioned at end of inserted text");

      // Add requirement template
      var templateRange = context.document.getSelection();

      // Insert the template text
      templateRange.insertText("\nLorsque ", Word.InsertLocation.end);
      templateRange = templateRange.insertText("[condition remplis]", Word.InsertLocation.end);
      templateRange.font.italic = true;
      templateRange = templateRange.insertText(", le ", Word.InsertLocation.end);
      templateRange.font.italic = false;
      templateRange = templateRange.insertText("[composant logiciel]", Word.InsertLocation.end);
      templateRange.font.italic = true;
      templateRange = templateRange.insertText(" doit ", Word.InsertLocation.end);
      templateRange.font.italic = false;
      templateRange = templateRange.insertText("[actions]", Word.InsertLocation.end);
      templateRange.font.italic = true;
      templateRange = templateRange.insertText(".\n\nDépend de :", Word.InsertLocation.end);
      templateRange.font.italic = false;

      console.log("Requirement template added successfully");

      // Move cursor to the end of the inserted text
      templateRange.select("End");
      console.log("Cursor positioned at end of inserted text");

      await context.sync();
    });

    console.log("=== Insert Requirement Function Completed Successfully ===");

  } catch (error) {
    console.error("Error inserting requirement:", error);

    // Show error notification if possible
    try {
      Office.context.document.settings.set("lastAction", "Error inserting requirement");
    } catch (notificationError) {
      console.error("Could not show error notification:", notificationError);
    }
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Replaces all [REQ_XXXX] placeholders in the document with actual numbered requirements
 * @param event {Office.AddinCommands.Event}
 */
async function ReplaceRequirement(event) {
  console.log("=== Starting requirement placeholder replacement ===");

  try {
    await Word.run(async (context) => {
      // Get the entire document text to search for placeholders and existing requirements
      console.log("Loading document content...");
      const documentBody = context.document.body;
      context.load(documentBody, "text");

      await context.sync();

      // Find all existing numbered requirements using regex
      const numberedRegex = /\[REQ_(\d{4})\]/g;
      const numberedMatches = documentBody.text.match(numberedRegex);

      let nextReqNumber = 1; // Default value if no requirements found

      if (numberedMatches) {
        console.log(`Found ${numberedMatches.length} existing numbered requirements`);
        // Extract numbers and find the highest one
        const reqNumbers = numberedMatches.map(match => {
          const numberMatch = match.match(/\[REQ_(\d{4})\]/);
          return parseInt(numberMatch[1], 10);
        });

        const maxReqNumber = Math.max(...reqNumbers);
        nextReqNumber = maxReqNumber + 1;
        console.log(`Starting replacement from requirement number: ${nextReqNumber}`);
      } else {
        console.log("No existing numbered requirements found, starting with REQ_0001");
      }

      // Find all placeholder requirements [REQ_XXXX]
      const placeholderRegex = /\[REQ_XXXX\]/g;
      const placeholderMatches = documentBody.text.match(placeholderRegex);

      if (!placeholderMatches) {
        console.log("No [REQ_XXXX] placeholders found in document");
        // Show completion message
        Office.context.document.settings.set("lastAction", "No placeholders found to replace");
        event.completed();
        return;
      }

      console.log(`Found ${placeholderMatches.length} placeholders to replace`);

      // Replace each placeholder one by one
      let replacementCount = 0;

      for (let i = 0; i < placeholderMatches.length; i++) {
        try {
          // Search for the placeholder text
          const searchResults = context.document.body.search("[REQ_XXXX]", { matchCase: true, matchWholeWord: false });
          context.load(searchResults, "items");

          await context.sync();

          if (searchResults.items.length > 0) {
            // Replace the first occurrence
            const currentReqNumber = nextReqNumber + i;
            const reqText = `[REQ_${currentReqNumber.toString().padStart(4, '0')}]`;
            const bookmarkName = `REQ_${currentReqNumber.toString().padStart(4, '0')}`;

            const rangeToReplace = searchResults.items[0];
            rangeToReplace.insertText(reqText, Word.InsertLocation.replace);
            context.load(rangeToReplace);

            await context.sync();

            // Add bookmark to the replaced requirement
            try {
              rangeToReplace.insertBookmark(bookmarkName);
              console.log(`Bookmark '${bookmarkName}' added successfully`);
            } catch (bookmarkError) {
              console.error(`Error adding bookmark for ${bookmarkName}:`, bookmarkError);
            }

            // Apply style if it exists
            try {
              rangeToReplace.style = "REQ_TITLE";
              await context.sync();
            } catch (styleError) {
              if (styleError.code === 'InvalidArgument') {
                // Style doesn't exist, continue without it
              } else {
                console.error("Error applying style:", styleError.message);
              }
            }

            replacementCount++;
            console.log(`Replaced placeholder ${i + 1}/${placeholderMatches.length} with ${reqText}`);
          }
        } catch (replaceError) {
          console.error(`Error replacing placeholder ${i + 1}:`, replaceError);
        }
      }

      console.log(`Successfully replaced ${replacementCount} placeholders`);
    });

    console.log("=== Replace Requirement Placeholders Function Completed Successfully ===");

  } catch (error) {
    console.error("Error replacing requirement placeholders:", error);

    // Show error notification if possible
    try {
      Office.context.document.settings.set("lastAction", "Error replacing placeholders");
    } catch (notificationError) {
      console.error("Could not show error notification:", notificationError);
    }
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

// Register the functions with Office
Office.actions.associate("InsertRequirement", InsertRequirement);
Office.actions.associate("ReplaceRequirement", ReplaceRequirement);
