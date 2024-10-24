# Hengji_SnipeIT_Companion

Introduction
This is an 3rd party add-on used to maintain a custom field in SnipeIT.
Step 1: Pull response of all assets from SnipeIT.
Step 2: Generate custom asset tag.
Step 3: Put asset tag back into SnipeIT.
Step 4: Print the tag and an extra label for newly added assets.
Step 5: Generate an asset report for newly added assets.

Required Entry in SnipeIT:
1. Company (Brand)
2. Location

Release Notes:
v1.0 Auto-generate custom field based on other info; Export updated assets as an Excel spreadsheet
v1.1 Automatically updates the updated assets’ status from "in stock" to "in transfer"
v1.2 Minor debug
v1.3 Added label printing function.
v1.4 Void status update.

Work Pending:
1. Add second label printer.
2. Modification in exported asset report.
3. Confirm if need to change status upon assigning the asset tag.
4. Change location's CN to 商场地点/客户名称
5. Use of df index could be wrong, need to be checked
6. Move dedicated asset tag numbers to a private file.