# Sheet Set fields for the title block

The add-in uses workflow B:

- `INNO-STT` is the frame number, initial sheet number, and layout name.
- `INNO_NAME_DRAWING` is the initial sheet title.
- `PLSHEETSET` writes those values to the DST as `Sheet Number` and `Sheet Title`.
- The title block in Paper Space displays the DST values through Sheet Set fields.

## Prepare the title-block DWG

Edit the title-block source used by the template layout:

1. In the drawing-number attribute, insert the Sheet Set field **CurrentSheetNumber**.
2. In the drawing-title attribute, insert the Sheet Set field **CurrentSheetTitle**.
3. Save the title-block DWG and reload its xref in the host drawing.

Do not use the Model-space `INNO-STT` and `INNO_NAME_DRAWING` attributes as title-block
fields. Those attributes identify frames for `PLSTT` / `PLAYOUT`; the Paper-space title
block should read the sheet metadata from the DST.

## Workflow

1. Run `PLSTT` to assign `INNO-STT` and optional `INNO_NAME_DRAWING`.
2. Run `PLAYOUT`; layout names are created from `INNO-STT`.
3. Run `PLSHEETSET`.
4. Review or edit `Sheet Number` and `Drawing Name / Sheet Title`.
5. Click **Create / Update DST**.
6. Save the host DWG and run `REGEN` if the fields have not refreshed yet.

Editing Sheet Number or Sheet Title in the sheet-set table does not rename layouts and
does not write back to Model-space frame attributes. The title block follows the DST
because its values are fields, not because those objects are synchronized.
