# Workflow Excel VBA Macros

This repository contains the VBA modules for the **Workflow** Excel workbook. The macros help manage order stages, file attachments and synchronize data between sheets.

## Contents
- `ThisWorkbook` – Workbook event handlers.
- `frmAttach` – User form for attaching proof/email/print files.
- `modAttachBindings` – Routines called by Attach buttons.
- `modDesignAttach` – Adds "View" or "Browse" buttons on the Design sheet.
- `modFlagSync` – Syncs checkbox flags on stage sheets back to the Master table.
- `modHyperlinkUtils` – Shared helpers for working with hyperlinks.
- `modMasterSync` – Synchronizes row changes with the Master table.
- `modOrderEntrySync` – Processes Order Entry rows and refreshes stage sheets.
- `modShowAttachForm` – Central `ShowAttachForm` routine.
- `modStageSync` – Refreshes all stage sheets.
- `modSyncHelpers` – Utility routines for clearing and repopulating data.

## Usage
1. Open the Excel workbook and import these modules into the VBA editor.
2. Ensure each worksheet contains the expected tables (`tblMaster`, `tblDesign`, `tblOrderEntry`, etc.).
3. Macros will run automatically via workbook events or button clicks.
4. To refresh all stages manually, run `RefreshAllStages` from `modStageSync`.

## Notes
- The code assumes hyperlink cells contain either a hyperlink or a file path string.
- `frmAttach` now exposes `UpdateCellHyperlink` as a private routine to avoid name collisions with the helper module.
