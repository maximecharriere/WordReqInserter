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
async function insertNewRequirement(event) {
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
      
      // Format the requirement text
      const reqText = `[REQ_${nextReqNumber.toString().padStart(4, '0')}]`;
      const bookmarkName = `REQ_${nextReqNumber.toString().padStart(4, '0')}`;
      console.log(`Inserting requirement: ${reqText}`);
      
      // Insert the requirement at the current selection
      const selection = context.document.getSelection();
      const insertedRange = selection.insertText(reqText, Word.InsertLocation.replace);
      context.load(insertedRange);
      
      await context.sync();
      console.log("Requirement text inserted successfully");
      
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

// Register the function with Office
Office.actions.associate("insertNewRequirement", insertNewRequirement);
