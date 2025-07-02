# WordReqInserter

Word Add-in for inserting unique requirement numbers in Word documents. Automatically generates incremental requirement IDs in the format [REQ_XXXX] with bookmarks for easy cross-referencing.

## Features

- **Automatic Numbering**: Generates unique requirement IDs ([REQ_0001], [REQ_0002], etc.)
- **Smart Incrementing**: Scans document for existing requirements and continues numbering
- **Bookmark Creation**: Automatically creates bookmarks for cross-referencing
- **Style Support**: Applies REQ_TITLE style if available
- **Position Independent**: Numbers based on existing requirements, not document position

---

# Add the Requirement Inserter to Word

## On Word Online:
1. Open a Word document
2. Go under **Home** > **Add-Ins** group > **Add-Ins** command (or search for "Add-Ins") > **Advanced...** > **Load my add-in**
3. Upload the manifest.xml that can be found on the Shared Folder:  
   `\\SharedFolder\OfficeAddins\manifest.xml` (replace with actual location)

## On Word Desktop:
1. Open Word
2. Go under **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**
3. Add the catalog URL:  
   `\\SharedFolder\OfficeAddins` (replace with actual location)
4. Check "Show in Menu"
5. Restart Word Desktop
6. Go under **Home** > **Add-Ins** group > **Add-Ins** command (or search for "Add-Ins") > **Advanced...**
7. Under **Shared Folder** tab, select **WordReqInserter**

---

# Use the Requirement Inserter in Word

## Add a Requirement

### 1. Create the REQ_TITLE Style (Optional)
- The requirement will use the **REQ_TITLE** style if available
- Create a Style named **REQ_TITLE** in your document for consistent formatting
- If the style doesn't exist, requirements will use default formatting

### 2. Insert a Requirement
1. Position your cursor where you want to insert the requirement
2. Click **Home** > **ESPI** group > **Insert Requirement** command
3. A unique requirement ID will be inserted in the format **[REQ_XXXX]**
4. The ID is incremental and doesn't depend on the position in the document
5. A bookmark is automatically created for cross-referencing

## Cross-Reference to a Requirement

1. Under **Insert** > **Links** group > **Cross-reference** command
2. Select **Reference type**: **Bookmark**
3. Select your requirement bookmark (REQ_XXXX format)
4. Choose to insert the **Bookmark text** in the document
5. Click **Insert**
