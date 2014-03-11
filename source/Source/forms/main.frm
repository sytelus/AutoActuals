VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AutoActuals"
   ClientHeight    =   2940
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   4455
   ClipControls    =   0   'False
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtToDate 
         Height          =   285
         Left            =   2460
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkIncludeNotes 
         Caption         =   "Include &Notes"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   2220
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CheckBox chkIncludeEntryExitTime 
         Caption         =   "Include &Entry-Exit time analysis"
         Height          =   195
         Left            =   540
         TabIndex        =   5
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CheckBox chkIncludeProjectName 
         Caption         =   "Include &Project name"
         Height          =   195
         Left            =   540
         TabIndex        =   4
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtFromDate 
         Height          =   285
         Left            =   540
         TabIndex        =   2
         Top             =   960
         Width           =   1515
      End
      Begin VB.OptionButton opnActivityReport 
         Caption         =   "&Activity Report from:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   1755
      End
      Begin VB.OptionButton opnReportForToday 
         Caption         =   "&Caption for Today at run-time"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "&To"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&No"
      Height          =   435
      Left            =   2640
      TabIndex        =   8
      Top             =   2340
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   435
      Left            =   720
      TabIndex        =   7
      Top             =   2340
      Width           =   1155
   End
   Begin VB.Timer WakeUpTimer 
      Left            =   180
      Top             =   2460
   End
   Begin VB.Label lblReportInfo 
      Caption         =   "Caption at run time"
      Height          =   315
      Left            =   420
      TabIndex        =   9
      Top             =   2460
      Visible         =   0   'False
      Width           =   4155
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Menu for system tray"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "O&ptions"
         Begin VB.Menu mnuAutoPopup 
            Caption         =   "&Auto Popup"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuEnforceStatusReport 
            Caption         =   "&Enforce Status report reminder"
         End
         Begin VB.Menu mnuDoNotShowCompletedItems 
            Caption         =   "&Do not show completed tasks"
         End
      End
      Begin VB.Menu mnuResetTimer 
         Caption         =   "&Reset Timer..."
      End
      Begin VB.Menu mnuChangeDSN 
         Caption         =   "C&hange DSN..."
      End
      Begin VB.Menu mnuOpenDatabase 
         Caption         =   "Open &Database..."
      End
      Begin VB.Menu mnuMakeStatusReport 
         Caption         =   "&Make Status Report..."
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlWakeUpInterval As Long
Dim mlSubInterval As Long

Dim mlProject As Long
Dim mlTask As Long

Dim msActivity As String
Dim mdtLastActivationTime As Date
Dim mdtLastSaved As Date

Private Sub cmdCancel_Click()
    
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()

    Dim sMessage As String
    
    If ValidateForm(sMessage) Then

        If opnReportForToday.Value Then
            
            Call MakeStatusReport(opnReportForToday.Value)
            
        Else
        
            Call MakeStatusReport(opnReportForToday.Value, CDate(txtFromDate), CDate(txtToDate), CheckToBool(chkIncludeEntryExitTime.Value), CheckToBool(chkIncludeProjectName.Value), CheckToBool(chkIncludeNotes.Value))
        
        End If
        
        Me.Hide
        
    Else
    
        MsgBox sMessage
        
    End If

End Sub

Private Sub Form_Load()

    On Error GoTo ERR_Form_Load
    
    Dim lTimerTag As Long
    Dim bSuccess As Boolean
    Dim sErrLocation As String
    
    sErrLocation = "Form_Load"
    
    '----------------------------------------------------------------
    
    If Command$ = "debug" Then
    
        gbIsDebugMode = True
        
    Else
    
        gbIsDebugMode = False
    
    End If

    '----------------------------------------------------------------
       
    sErrLocation = "Form_Load - Checking Prev Instance"
       
    'Is prev instance is all ready running
    If PrevInstanceHandled Then
    
        'Invoked last instance successfully, end yourself
        ExitApp
        
    End If
    
    '----------------------------------------------------------------
    sErrLocation = "Form_Load - Adding Icon to sys tray"
    
    'Show the icon in System Tray - as soon as app is started
    Call AddIconInSysTray(Me, gsTRAY_HINT)
    
    '----------------------------------------------------------------
    
    sErrLocation = "Form_Load - Loading settings from registry"
  
    'Get user settings from registry - Interval, DSN etc.
    'This settings are used by many of the functions afterwards, so it must be run before anything else
    LoadSettings
    
    '----------------------------------------------------------------
  
    sErrLocation = "Form_Load - Running DSN Changed tasks"
  
    Call TasksWhenDSNIsChanged(False)
    
    If gbIsDebugMode Then
        WakeUp
    End If
    
    
    sErrLocation = "Form_Load - Starting the timer"
    
    'If AutoPopup not required do not start the timer
    If mnuAutoPopup.Checked Then
    
        'Start the timer
        WakeUpTimer.Enabled = False
        
        WakeUpTimer.Interval = mlSubInterval
        
        WakeUpTimer.Enabled = True
        
    End If
    
    '----------------------------------------------------------------
    
Exit Sub
ERR_Form_Load:

    MsgBox "Error: " & Err.Number & ": " & Err.Description & " at " & sErrLocation
    
    ReRaisErr
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'This event DOES not occure in MouseMove, but it only
    'occures when user clicks or moves mouse on icon in system tray
    'The X parameter contains wParam of Windows message for ShellNotify, not the X cordinate of MouseMove event
    
    Dim lTimeLeftInInterval As Long     'Holds how much minutes remaining before popup occures
    
    If Not Me.Visible Then
    
        If frmWakeUpDlg.Visible Then
            
            Call SetTrayHint("Activated")
            
        Else
        
            Select Case X
            
            Case 7680 'MouseMove
            
                If mnuAutoPopup.Checked Then
                
                    If IsNumeric(WakeUpTimer.Tag) Then
                    
                        lTimeLeftInInterval = mlWakeUpInterval - CLng(WakeUpTimer.Tag)
                        
                        If lTimeLeftInInterval <> 0 Then
                        
                            Call SetTrayHint(lTimeLeftInInterval & " min left")
                            
                        Else
                        
                            Call SetTrayHint("just about to popup!")
                            
                        End If
                        
                    Else
                        
                        Call SetTrayHint("Calculating time left...")
                        
                    End If
                    
                Else
                
                    Call SetTrayHint("No Auto Popup")
                
                End If
            
            Case 7695 'LeftMouseDown
            
            Case 7710 'LeftMouseUp
            
            Case 7725, 115875 '7725-LeftDblClick , 115875 is sent via another instace to wake up in this instance
                
                WakeUp
                
            Case 7740 'RightMouseDown
            
                'Sometimes mouse stays hourglass...
                Call SetMousePointer(vbDefault)
                
                PopupMenu mnuPopup, 0, , , mnuOpen
            
            Case 7755 'RightMouseUp
            
            Case 7770 'RightDblClick
                
            End Select
            
        End If
        
    Else
        
        Call SetTrayHint("dialog for status report open")
        
    End If

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error GoTo Err_QueryUnload
    
    'If debug mode then do not save DSN and interval changes in registry because that will effect real settings for currently running program (that I'm actually using it)
    If Not gbIsDebugMode Then
    
        If IsStatusReportReq Then
        
            Cancel = 1
        
        End If
        
        If Cancel = 0 Then
        
            'Following is used by crash recovery routine
            SaveSetting gsREG_APP_NAME, "LastValues", "ClosedProperly", "True"
            SaveSetting gsREG_APP_NAME, "LastValues", "TimerTag", AlternateStrIfNull(WakeUpTimer.Tag, "0")
            SaveSetting gsREG_APP_NAME, "LastValues", "ClosingTime", Now
            
            'Note down the closing event
            SaveActivityForSysTask gsCLOSING_TASK_NAME, "<Closing>"
            
        End If
        
    End If
    
    
    'Remove the icon from system tray
    If Cancel = 0 Then
    
        DeleteIconFromSysTray
        
        'Due to unknown reasons, app remains in memory - so use End
        End
        
    End If
    
Exit Sub
Err_QueryUnload:
    
    Resume Next
    
    End
    
End Sub

Private Sub mnuAutoPopup_Click()
    
    Call SetAutoPopup(Not mnuAutoPopup.Checked)
    
End Sub

Private Sub mnuChangeDSN_Click()
    
    Dim sUserResponce As String
    Dim bSuccess As Boolean
    
    bSuccess = True
    
    sUserResponce = InputBox("Enter the new DSN:", , gsDSN)
    
    'Does user pressed Cancel?
    If sUserResponce <> "" Then
            
        gsDSN = sUserResponce
            
        'Get the Global IDs and make the <Startup Entry>
        Call TasksWhenDSNIsChanged(True)
        
        'Save the new DSN in registry
        SaveSetting gsREG_APP_NAME, "Settings", "DSN", gsDSN
    
    End If

End Sub

Private Sub mnuClose_Click()

    If Not IsStatusReportReq Then
    
        'Whether to prepare status or not is already asked to user, so make the flag true to prevent asking it again in Form's Unload event
        gbStatusReportWasGenerated = True
        
        ExitApp
        
    End If

End Sub

Private Sub mnuDoNotShowCompletedItems_Click()
    
    Call SetDoNotShowCompletedItems(Not mnuDoNotShowCompletedItems.Checked)
    
End Sub

Private Sub mnuEnforceStatusReport_Click()

    Call SetEnforceStatusReport(Not mnuEnforceStatusReport.Checked)

End Sub

Private Sub mnuMakeStatusReport_Click()
    
    ShowStatusReportDialog
    
End Sub

Private Sub mnuOpen_Click()

    WakeUp

End Sub

Private Sub mnuOpenDatabase_Click()

    Dim sDatabaseFile As String
    Dim sUserResponce As String
    
    sDatabaseFile = GetSetting(gsREG_APP_NAME, "Settings", "DatabaseFile", "")
    
    sUserResponce = InputBox("Enter the name of database file (you have to enter only once):", , sDatabaseFile)
    
    If sUserResponce <> "" Then
    
        Call SaveSetting(gsREG_APP_NAME, "Settings", "DatabaseFile", sUserResponce)
        
        If Not OpenAnyFile(sUserResponce) Then
        
            MsgBox "Can not execute associated program with database file."
        
        End If
    
    End If
    
End Sub

Private Sub mnuResetTimer_Click()
    
    If MsgBox("Do you really want to reset the interval timer?", vbYesNo) = vbYes Then
    
        WakeUpTimer.Tag = "0"
        
        WakeUpTimer.Enabled = False
        
        WakeUpTimer.Enabled = True
    
    End If
    
End Sub

Private Sub opnActivityReport_Click()

    Call EnableActivityReportOptionControls(True)
    
End Sub

Private Sub opnReportForToday_Click()

    Call EnableActivityReportOptionControls(False)
    
End Sub

Private Sub WakeUpTimer_Timer()
    
    Dim lTimerEventCount As Long
    
    'If this is non-auto popup mode, it's incorrect to have time event (something went wrong)
    If Not mnuAutoPopup.Checked Then
    
        Call SetAutoPopup(False)
        
    Else
    
        'First disable the timer until this event is completed
        WakeUpTimer.Enabled = False
        
        'Get the numbers of minutes passed
        lTimerEventCount = CLng(WakeUpTimer.Tag)
        
        'At frequent numbers of minutes, refresh the list of running programms
        If (lTimerEventCount Mod glWIN_LIST_SAMPLE_RATE) = 0 Then
        
            FillWinListInArr
            
        End If
        
        
        'If number of minutes passed >= interval set
        If lTimerEventCount >= mlWakeUpInterval Then
        
            'Re-init the minute counter
            lTimerEventCount = 0
            
            'Display the "what you were doing" dialog
            If WakeUp Then
            
                'If user pressed OK, erase the list else retain the entries
                Erase gavntWinList
                
            End If
            
        Else
        
            'if minutes passed is hasn't reach the interval, increment the minute counter
            lTimerEventCount = lTimerEventCount + 1
            
        End If
        
        'Save the minute count in tag
        WakeUpTimer.Tag = CStr(lTimerEventCount)
        
        'Restart the timer
        WakeUpTimer.Enabled = True
        
    End If
    
End Sub

Private Function WakeUp() As Boolean
    
    Dim bUnload As Boolean
    Dim bIsTimerEnabled As Boolean
    Dim lTimeSpent As Long
    'Dim ofrmWakeUpDlg As frmWakeUpDlg
    Dim bUserResponce As Boolean
        
        
    'Bydefault assume that User pressed cancel - so return False
    WakeUp = False
    
    'Record the activation time
    RecordActivationTime
    
    'Save the timer status
    bIsTimerEnabled = WakeUpTimer.Enabled
    
    'Disable the timer status while dialog is on
    WakeUpTimer.Enabled = False
    
    'Update the list of windows before dialog is shown up
    FillWinListInArr
    
    'Display the dialog
    'Set ofrmWakeUpDlg = New frmWakeUpDlg
    
    bUserResponce = frmWakeUpDlg.DisplayForm(msActivity, mlWakeUpInterval, mlProject, mlTask, mdtLastSaved, bUnload, lTimeSpent)
    
    'Set ofrmWakeUpDlg = Nothing
    
    
    If bUserResponce Then
    
        'User pressed OK. So return true
        WakeUp = True
        
        'If debug mode then do not save DSN and interval changes in registry because that will effect real settings for currently running program (that I'm actually using it)
        If Not gbIsDebugMode Then
        
            'This settings should be persistant i.e. reloaded when program starts again
            
            SaveSetting gsREG_APP_NAME, "LastValues", "Project", mlProject
            
            SaveSetting gsREG_APP_NAME, "LastValues", "Task", mlTask
                    
            SaveSetting gsREG_APP_NAME, "Settings", "DSN", gsDSN
            
            SaveSetting gsREG_APP_NAME, "Settings", "WakeUpInterval", mlWakeUpInterval
        
        End If
        
        'Save the activity as entered in dialog
        Call SaveActivity(msActivity, Now, glUser_ID, GetMachineName, mlTask, lTimeSpent)
        
        'Save the last saved time
        RecordLastSavedTime
    
    End If
        
    If bUnload Then
    
        If MsgBox("Do you really want to shutdown the AutoActuals?", vbYesNo) = vbYes Then
    
            ExitApp
                        
        End If
    
    End If
           
    'Restore the timer status
    WakeUpTimer.Enabled = bIsTimerEnabled
           
End Function


Private Sub LoadSettings()

    Dim sWakeUpInterval As String       'Intermediate value obtained from registry
    Dim sProjectTempVal As String       'Temporarily stores string value obtained from Project key
    
    '
    'Following Settings retrived:
    'gsDSN, mlWakeUpInterval, AutoPopup, mlProject, mlTask
    '
    
    
    'In debug mode have small interval for faster testing and different database
    If gbIsDebugMode Then
    
        '
        'This settings depends on what r u debugging
        '
        
        'Get the DSN from registry
        gsDSN = "Actuals_Test"
    
        sWakeUpInterval = "60"
        
        SetAutoPopup False
        
        mlProject = 1
        
        mlTask = 1
        
        'In debugging mode, I want things promptly
        mlSubInterval = 100
        
    Else
 
        'Get the DSN from registry - Null DSN condition is used to detect whether user is running application first time
        gsDSN = GetSetting(gsREG_APP_NAME, "Settings", "DSN", "")
           
        'By default Interval is 60 min
        sWakeUpInterval = GetSetting(gsREG_APP_NAME, "Settings", "WakeUpInterval", "60")
    
        'Restore AutoPopup settings
        Call SetAutoPopup(GetSetting(gsREG_APP_NAME, "Settings", "AutoPopup", "True"))
        
        'Restore NoCompletedItems settings
        Call SetDoNotShowCompletedItems(GetSetting(gsREG_APP_NAME, "Settings", "NoCompletedItems", "False"))
        
        'Restore Enforce status report settings
        Call SetEnforceStatusReport(GetSetting(gsREG_APP_NAME, "Settings", "EnforceStatusReport", "True"))
        
        
        'Get the project name from the registry
        sProjectTempVal = GetSetting(gsREG_APP_NAME, "LastValues", "Project", "-1")
         
         If IsNumeric(sProjectTempVal) Then
         
            mlProject = CLng(sProjectTempVal)
            
        Else
        
            mlProject = -1
        
        End If
        
        'Get the last task ID
        mlTask = GetSetting(gsREG_APP_NAME, "LastValues", "Task", -1)
        
    End If
            
    'Convert the interval obtained from registry to long value
    mlWakeUpInterval = CLng(sWakeUpInterval)
    
    mdtLastActivationTime = GetSetting(gsREG_APP_NAME, "LastValues", "LastActivationTime", "12-Feb-1975")
    
    mdtLastSaved = GetSetting(gsREG_APP_NAME, "LastValues", "LastSavedTime", Now)
        
    'One minute = 60000 ms - Time unit is not allowed to change for the time being
    mlSubInterval = 60000
    
    gbStatusReportWasGenerated = False
    
End Sub

Private Sub OnFirstTime()
    
        msActivity = "Welcome to the AutoActuals!"
        msActivity = msActivity & vbCrLf & "----------------------------------------------"
        msActivity = msActivity & vbCrLf
        msActivity = msActivity & vbCrLf & "Hot keys:"
        msActivity = msActivity & vbCrLf & "Press Ctrl + Enter to save and close this dialog."
        msActivity = msActivity & vbCrLf & "Press Ctrl + S to save without closing this dialog."
        msActivity = msActivity & vbCrLf & "Press F1 to get hints on what you were doing."
        msActivity = msActivity & vbCrLf
        msActivity = msActivity & vbCrLf & "To add/modify/delete any item, select that item and press Alt and + key/~ key/X key."
        
        WakeUp
        
        msActivity = "Now onwards, just type here what " & vbCrLf & "you was doing and press F2 or Ctrl+S to close and save this window." & vbCrLf & "This will enable you to track your activities and write actuals."

End Sub

Private Function IsClockCorrect() As Boolean
    IsClockCorrect = (DateDiff("n", mdtLastActivationTime, Now) >= 0)
End Function

Private Sub RecoverInterval(ByVal vdtLastTime As Date, ByVal vlNewTimerTag As Long)
    
    'If restarted within 5Hrs or date is not changed
    If (DateDiff("n", vdtLastTime, Now) <= (5 * 60)) Or (Day(vdtLastTime) = Day(Now)) Then
    
        'Try to adjust the timer tag so interval is recovered
        
        'Is timer interval already exceeded
        If vlNewTimerTag < mlWakeUpInterval Then
        
            WakeUpTimer.Tag = vlNewTimerTag
            
            'Interval was recovered so don't alter LastSavedTime, it's already loaded from registry
        
        Else
        
            msActivity = "There was more then " & vbCrLf & mlWakeUpInterval & " min passed after you have resatrted."
            
            WakeUp
            
            'Reset the passed interval to 0 minutes
            WakeUpTimer.Tag = 0
            
            'Reset the LastSaved Time to current time
            mdtLastSaved = Now
        
        End If
        
        
        
    Else
        
        'Reset the passed interval to 0 minutes
        WakeUpTimer.Tag = 0
        
        'Interval was not recovers so reset last saved to current time
        mdtLastSaved = Now

    End If
    
End Sub


Private Sub FillWinListInArr()
    
    'If list has became too long, recreate it
    If GetDimension(gavntWinList) = 2 Then
    
        If UBound(gavntWinList, 2) > 100 Then
        
            Erase gavntWinList
        
        End If
        
    End If
    
    GetWinList gavntWinList, True
    
End Sub

Private Sub SetAutoPopup(ByVal bEnableAutoPopup As Boolean)
    
    mnuAutoPopup.Checked = bEnableAutoPopup
    
    WakeUpTimer.Enabled = bEnableAutoPopup
    
    mnuResetTimer.Enabled = bEnableAutoPopup
    
    If Not gbIsDebugMode Then
    
        Call SaveSetting(gsREG_APP_NAME, "Settings", "AutoPopup", bEnableAutoPopup)
        
    End If
    
End Sub

Private Function PrevInstanceHandled() As Boolean

    Dim bResult As Boolean
    Dim lLastMainFormHandle As Long     'Stores form handle of last instance
    
    'By default assume Previose instance was not there
    bResult = False
    
    'First check if previouse instance is running
    If App.PrevInstance Then
    
        'Get the main form window handle of previouse instance
        lLastMainFormHandle = GetSetting(gsREG_APP_NAME, "LastValues", "MainFormHandle", "0")
        
        'If handle is valid
        If lLastMainFormHandle <> 0 Then
        
            'Send a message to main form to ask it to popup
            If PostMessage(lLastMainFormHandle, WM_MOUSEMOVE, 0&, 7725&) Then
                
                   'Successfully invoked prev instance
                   bResult = True
            
            End If
        
        End If
    
    Else
    
        'If debug mode then do not save changes in registry because that will effect real settings for currently running program (that I'm actually using it)
        If Not gbIsDebugMode Then
        
            'No Prev Instance, so save form handle
            Call SaveSetting(gsREG_APP_NAME, "LastValues", "MainFormHandle", Me.hwnd)
            
        End If
    
    End If
    
    PrevInstanceHandled = bResult

End Function


Private Function CheckDSNAndGetGlobalIDs() As Boolean

    Dim bSuccess As Boolean
    Dim bWasDSNEmpty As Boolean
    Dim vntUsers As Variant
    Dim lRowCount As Long
    
    bSuccess = True
    
    'Is DSN is null
    If gsDSN = "" Then
    
        bWasDSNEmpty = True
        
        'Ask the user for DSN
        gsDSN = InputBox("You are probably running this application for the first time. Please enter the DSN:")
        
        'Does user pressed Cancel or blank DSN?
        If gsDSN = "" Then
           
            MsgBox "This application must be supplied valid DSN!"
            
            bSuccess = False
        
        End If
        
    Else
    
        bWasDSNEmpty = False
    
    End If
    
    
    If bSuccess Then
   
        'Get the ID of Global user
        If Not GetID("Users", "User_ID", "User_Login_Name", gsGLOBAL_USER_NAME, glGlobal_User_ID) Then
         
             MsgBox "User named " & gsGLOBAL_USER_NAME & " must exist in Users table (this user is Global user)."
             
             bSuccess = False
         
         End If
         
         gsUserLoginName = GetLoginName("default")
         
         If bSuccess Then
         
            'Get the ID of the user for current login name
            If Not GetID("Users", "User_ID", "User_Login_Name", gsUserLoginName, glUser_ID) Then
            
                'Check if more then 1 visble users exist in database - if no then this is first user so make him/her admin
                Call GetRowsInArr("SELECT * FROM Users WHERE Is_Visible = True", vntUsers, lRowCount)
                
                If (UBound(vntUsers, 2) - LBound(vntUsers, 2) + 1) = 1 Then
            
                    'Create the new user with admin rights
                    If Not AddARow("Users", Array("User_Login_Name", "User_Type_ID"), Array(gsUserLoginName, gnUSER_TYPE_ADMINISTRATOR), "User_ID", glUser_ID) Then
                    
                        MsgBox "User name '" & AlternateStrIfNull(gsUserLoginName, "<No User name>") & "' can not be added to database." & vbCrLf & "Please log on as a valid user in Windows or contact your system administrator."
                        
                        bSuccess = False
                    
                    End If
                    
                    glUser_Type_ID = gnUSER_TYPE_ADMINISTRATOR
                    
                Else
                
                    MsgBox "Your account does not exist in database. Please contact the system administrator to create the account for " & gsUserLoginName
                    
                    bSuccess = False
                
                End If
                
            Else

                'Get the user type
                If Not GetID("Users", "User_Type_ID", "User_Login_Name", gsUserLoginName, glUser_Type_ID) Then
                
                    MsgBox "Error occured while checking access rights for user."
                    
                    bSuccess = False
                
                End If
            
            End If
            
        End If
        
        If bSuccess Then
        
            'Get the ID of the machine
            If Not GetID("Machines", "Machine_ID", "Name", GetMachineName, glMachine_ID) Then
            
                'Add this machine in the list
                If Not AddARow("Machines", Array("Name"), Array(GetMachineName), "Machine_ID", glMachine_ID) Then
                
                    MsgBox "Machine name '" & AlternateStrIfNull(GetMachineName, "<No Machine name>") & "' can not be added to database." & vbCrLf & "Please give your machine a valid name from Control Panel > Network > Idntification tab or contact your system administrator."
                    
                    bSuccess = False
                
                End If
            
            End If
        
        End If
         
    End If
    
    
    'If everything gone fine then show up the wake up dialog
    If bSuccess And bWasDSNEmpty Then
    
        'Save the DSN in registry
        SaveSetting gsREG_APP_NAME, "Settings", "DSN", gsDSN
        
        'Execute the first time routines
        OnFirstTime
    
    End If
    
    
    CheckDSNAndGetGlobalIDs = bSuccess

End Function

Private Sub CrashRecovery()

    On Error GoTo ERR_CrashRecovery

    Dim dtLastClosingTime  As Date
    Dim lLastTagValue As Long
    Dim bWasCrashedLastTime As Boolean
    Dim lMinutesPassed As Long
    Dim sErrorLocation As String
    
    'Interval Crash Recovery routine
    
    
    sErrorLocation = "CrashRecovery:Get ClosedProperly"
    
    'First see if last time, app was closed properly
    If GetSetting(gsREG_APP_NAME, "LastValues", "ClosedProperly", "True") = "True" Then
    
        bWasCrashedLastTime = False
        
    Else
    
        bWasCrashedLastTime = True
    
    End If
    
    
    '
    'Now save the status that app is not yet closed properly in registry
    '
    
    'If debug mode then do not save changes in registry because that will effect real settings for currently running program (that I'm actually using it)
    If Not gbIsDebugMode Then
    
        sErrorLocation = "CrashRecovery:Save ClosedProperly"
        
        SaveSetting gsREG_APP_NAME, "LastValues", "ClosedProperly", "False"
        
        
        sErrorLocation = "CrashRecovery:Save LastActivationTime"
        
        'Save the last activation time in registry - This will be used to recover interval if crash occures
        SaveSetting gsREG_APP_NAME, "LastValues", "LastActivationTime", mdtLastActivationTime
        
    End If
    
    
    If IsClockCorrect Then
    
        'If machine was not crashed last time
        If Not bWasCrashedLastTime Then
        
            sErrorLocation = "CrashRecovery:Get ClosingTime"
            
            'It means that Last Closing Time saved in registry is valid
            dtLastClosingTime = GetSetting(gsREG_APP_NAME, "LastValues", "ClosingTime", "12-Feb-1975")

            sErrorLocation = "CrashRecovery:DateDiff dtLastClosingTime"

            'Calculate the numbers of minutes passed since last closing time
            lMinutesPassed = DateDiff("n", dtLastClosingTime, Now)
            
            sErrorLocation = "CrashRecovery:Get TimerTag"
            
            'Get how much minutes was already passed from interval
            lLastTagValue = CLng(GetSetting(gsREG_APP_NAME, "LastValues", "TimerTag", "0"))
            
            sErrorLocation = "CrashRecovery:RecoverInterval"
            
            'Following function will see if minutes already passed from interval before closing + minutes passed since closing is within interval bound or not and will take action
            Call RecoverInterval(dtLastClosingTime, lMinutesPassed + lLastTagValue)
        
        Else
            
            sErrorLocation = "CrashRecovery:DateDiff mdtLastActivationTime"
            
            'There was a crash. Their is no way to calculate exact time passed when system restarted. So calculate how much time passed since last activation
            lMinutesPassed = DateDiff("n", mdtLastActivationTime, Now)
            
            
            sErrorLocation = "CrashRecovery:ELSE RecoverInterval"
            
            'Following function will see if minutes already passed from last activation time is within interval bound or not and will take action
            Call RecoverInterval(mdtLastActivationTime, lMinutesPassed)
        
        End If
        
    Else
    
        sErrorLocation = "CrashRecovery:ELSE mdtLastSaved=Now"
        
        mdtLastSaved = Now
        
        MsgBox "The current system time is not correct. Many programs may not work properly if current system time is not correct."
    
    End If
    
Exit Sub
ERR_CrashRecovery:

    MsgBox "Error: " & Err.Number & ": " & Err.Description & " at " & sErrorLocation
    
    ReRaisErr


End Sub


Private Function SaveActivityForSysTask(ByVal vsTaskName As String, ByVal vsComment As String) As Boolean

    Dim lTaskID As Long
    Dim bSuccess As Boolean

    bSuccess = True


    'First get ID for task name
    If Not GetID("Tasks", "Task_ID", "Task_Name", vsTaskName, lTaskID) Then
     
         MsgBox "Task named " & vsTaskName & " must exist in Tasks table (this is the system task)."
         
         bSuccess = False
     
     End If
    
    
     If bSuccess Then
        
        'Then save current time
        bSuccess = SaveActivity(vsComment, Now, glUser_ID, GetMachineName, lTaskID, 0)
        
     End If
     
     SaveActivityForSysTask = bSuccess

End Function

Private Sub RecordActivationTime()
    
    'Record the activation time - this is not used however currently
    mdtLastActivationTime = Now
    
    'Save the last activation time in registry - This will be used to recover interval if crash occures
    SaveSetting gsREG_APP_NAME, "LastValues", "LastActivationTime", mdtLastActivationTime

End Sub


Private Sub RecordLastSavedTime()

    mdtLastSaved = Now
    
    'Save the last activation time in registry - This will be used to recover interval if crash occures
    SaveSetting gsREG_APP_NAME, "LastValues", "LastSavedTime", mdtLastSaved
    
End Sub


Private Sub TasksWhenDSNIsChanged(ByVal vboolSkipNotRequiredTasks As Boolean)

    On Error GoTo ERR_TasksWhenDSNIsChanged
    
    Dim sErrorLocation As String

    sErrorLocation = "TasksWhenDSNIsChanged:CheckDSNAndGetGlobalIDs"

    'Check out DSN, if not entered, then ask.
    If Not CheckDSNAndGetGlobalIDs Then
    
        'Can not DSN or user ID, so exit
        ExitApp
    
    End If
    
    sErrorLocation = "TasksWhenDSNIsChanged:CrashRecovery"
        
    If Not vboolSkipNotRequiredTasks Then
        
        '----------------------------------------------------------------
        
        'If machine is restarted then try to recover the lost minutes in interval
        CrashRecovery
        
        '----------------------------------------------------------------
    
    End If
    
    sErrorLocation = "TasksWhenDSNIsChanged:SaveActivityForSysTask"
    
    'Note down the startup time
    SaveActivityForSysTask gsSTART_UP_TASK_NAME, "<Started Up>"

    '----------------------------------------------------------------

Exit Sub
ERR_TasksWhenDSNIsChanged:

    MsgBox "Error: " & Err.Number & ": " & Err.Description & " at " & sErrorLocation
    
    ReRaisErr

End Sub


Private Sub SetDoNotShowCompletedItems(ByVal bEnableDoNotShowCompletedItems As Boolean)
    
    mnuDoNotShowCompletedItems.Checked = bEnableDoNotShowCompletedItems
    
    If Not gbIsDebugMode Then
    
        Call SaveSetting(gsREG_APP_NAME, "Settings", "NoCompletedItems", bEnableDoNotShowCompletedItems)
        
    End If
    
End Sub

Public Sub EnableActivityReportOptionControls(ByVal vboolEnable As Boolean)

    Call MakeAControlReadOnly(chkIncludeEntryExitTime, vboolEnable)
    
    Call MakeAControlReadOnly(chkIncludeProjectName, vboolEnable)
    
    Call MakeAControlReadOnly(chkIncludeNotes, vboolEnable)
    
    Call MakeAControlReadOnly(txtFromDate, vboolEnable)
    
    Call MakeAControlReadOnly(txtToDate, vboolEnable)

End Sub


Private Function ValidateForm(ByRef rsMessage As String) As Boolean

    Dim bSuccess As Boolean
    
    bSuccess = True
    
    rsMessage = "Following value(s) are not correct:"
    
    
    If Not opnReportForToday.Value Then
             
        If Not IsDate(txtFromDate) Then
        
            bSuccess = False
            
            rsMessage = rsMessage & vbCrLf & "The 'From Date' is not valid."
            
        End If
        
        
        If Not IsDate(txtToDate) Then
        
            bSuccess = False
            
            rsMessage = rsMessage & vbCrLf & "The 'To Date' is not valid."
            
        End If
    
    End If
    
    ValidateForm = bSuccess
    
End Function

Private Sub SetEnforceStatusReport(ByVal bEnableEnforceStatusReport As Boolean)
    
    mnuEnforceStatusReport.Checked = bEnableEnforceStatusReport
    
    If Not gbIsDebugMode Then
    
        Call SaveSetting(gsREG_APP_NAME, "Settings", "EnforceStatusReport", bEnableEnforceStatusReport)
        
    End If
    
End Sub


