# Workbook import and category authority

**Feature ID:** `FEATURE-WORKBOOK`

The owner imports `budget.xlsx` from the local workspace. The importer previews the `Manual Budget` categories, their evaluated monthly amounts, and the history summary before any write. It reads `Actual Charges` as labeled training evidence; it does not edit the workbook.

The import bridge explicitly maps `Food & Dining` to `Food & Dining & Entertainment`, `Insurance` to `Insurance (including car)`, and `Work Reimbursement` to an excluded, retained category. The owner must accept the preview. An unreadable formula, missing required column, changed workbook hash, or ambiguous bridge blocks the import and leaves the previous plan intact.
