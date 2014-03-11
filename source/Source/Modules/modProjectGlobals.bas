Attribute VB_Name = "modProjectGlobals"
Option Explicit

Public Const gsREG_APP_NAME As String = "AutoActuals"
Public Const glWIN_LIST_SAMPLE_RATE As Integer = 2
Public Const gsTRAY_HINT As String = "Rave Time Sheet"
Public Const gsGLOBAL_USER_NAME As String = "<everyone>"
Public Const gsSTART_UP_TASK_NAME As String = "<StartUp>"
Public Const gsCLOSING_TASK_NAME As String = "<StartUp>"
Public Const gsDEFAULT_SUB_MODULE_NAME As String = "<Default>"

Public Const gnMODE_PROJECT_EDITOR_ADD As Integer = 1
Public Const gnMODE_PROJECT_EDITOR_MODIFY As Integer = 2
Public Const gnMODE_PROJECT_EDITOR_DELETE As Integer = 3
Public Const gnMODE_TASK_EDITOR_ADD As Integer = 4
Public Const gnMODE_TASK_EDITOR_MODIFY As Integer = 5
Public Const gnMODE_TASK_EDITOR_DELETE As Integer = 6
Public Const gnMODE_USER_EDITOR_ADD As Integer = 7
Public Const gnMODE_USER_EDITOR_MODIFY As Integer = 8
Public Const gnMODE_USER_EDITOR_DELETE As Integer = 9
Public Const gnMODE_TASK_TYPE_EDITOR_ADD As Integer = 10
Public Const gnMODE_TASK_TYPE_EDITOR_MODIFY As Integer = 11
Public Const gnMODE_TASK_TYPE_EDITOR_DELETE As Integer = 12
Public Const gnMODE_MODULE_EDITOR_ADD As Integer = 13
Public Const gnMODE_MODULE_EDITOR_MODIFY As Integer = 14
Public Const gnMODE_MODULE_EDITOR_DELETE As Integer = 15

Public Const gnUSER_TYPE_NEWBEE As Integer = 1
Public Const gnUSER_TYPE_PROGRAMMER As Integer = 2
Public Const gnUSER_TYPE_SENIER_PROGRAMMER As Integer = 3
Public Const gnUSER_TYPE_PROCESS_LEADER As Integer = 4
Public Const gnUSER_TYPE_ADMINISTRATOR As Integer = 5

Public Const gnSECURITY_ALLOW As Integer = 1
Public Const gnSECURITY_READ_ONLY As Integer = 2
Public Const gnSECURITY_PARTIAL_READ_ONLY As Integer = 3
Public Const gnSECURITY_DENIY As Integer = 4



Public gavntWinList() As Variant
Public gsDSN As String
Public glUser_ID As Long
Public gsUserLoginName As String
Public glUser_Type_ID As Long
Public glGlobal_User_ID As Long
Public glMachine_ID As Long
Public gbIsDebugMode As Boolean
Public gbStatusReportWasGenerated As Boolean

