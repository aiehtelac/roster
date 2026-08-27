# Excel templates

One formatted workbook per roster type, resolved automatically by roster type:

    templates/Templates_HO.xlsx
    templates/Templates_MO.xlsx
    templates/Templates_REG.xlsx

Names are matched exactly, so the roster suffix must be upper case.
`.xlsm` is accepted too (`Templates_HO.xlsm`); `.xlsx` wins if both exist.

The same template is reused every month — the scheduler copies it, then writes the
month name into `B1`, the start date into `I5`, PH markers into row 1, and the roster
grid from row 6 / column I. A template is never modified in place.

If the file for a roster type is missing, the app still runs and produces CSV output
only, with a warning in the sidebar.

To update a template: replace the file and push. Community Cloud redeploys on push.
