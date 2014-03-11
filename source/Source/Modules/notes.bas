Attribute VB_Name = "modNotes"
Option Explicit

'       Project Notes
'       -------------
'
'0.9.1 features:
'1. DSN not saved in gsDSN bug solved
'2. Combo ItemData accessed via GetSelectedItemDataInCombo in SaveActivity
'3. Editable activity time spent
'4. DSN change moved to menu
'5. Option for not showing completed tasks
'
'0.9.1 Known bugs:
'1. Type Mismatch error 13 occures probably in CrashRecory routine
'2. Sometimes curor is busy while waking up
'
'
'0.9.2 Features:
'1. Time Spent field is made blank after entering one task.
'2. Solved Type Mismatch err 13 - it was due to saving null str in Timer tag to registry and then loading it for conversion to number
'3. For first time Save This command asks whether time spent is correct or not if that field is nor edited by user.
'4. If time_spent > time_interval then ask user if it's correct.
'5. Combobox width = largest str in combo
'6. Textbox height in Task editor reduced to std.


'
'Immediate To do:
'----------------
'
'Bug in crash recovery
'Bug in time shon in label
'
'
'
'
'
'
'
'

'To do:
'------
'
'1. What is no user is loggen on?
'
'
'
'
'
'
'Bugs:
'-----
'
'1. lblComment in frmWakeUpDlg doesn't gives proper time.
'
'
'
'
'
'
'
'
'
'Assumptions:
'------------
'
'1. Load settings must be called after settings are retrived from registry.
'2. At closing you must save closing status, closing time and current timer tag value.
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
