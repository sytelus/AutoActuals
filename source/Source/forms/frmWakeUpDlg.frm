VERSION 5.00
Begin VB.Form frmWakeUpDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AutoActuals"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWakeUpDlg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton btnSaveThis 
      Caption         =   "Sa&ve This"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   28
      ToolTipText     =   "Save the current entries without closing the window (Ctrl + S)"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtTimeSpent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4620
      TabIndex        =   24
      ToolTipText     =   "How much time you spent in this activity?"
      Top             =   4860
      Width           =   495
   End
   Begin VB.ComboBox cboTask 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Text            =   "cboTask"
      ToolTipText     =   "The Task you are working on"
      Top             =   660
      Width           =   3615
   End
   Begin VB.ComboBox cboProject 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "cboProject"
      ToolTipText     =   "The Project or Task Category you are working on"
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox txtInterval 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   7
      ToolTipText     =   "After how many minutes you want this dialog to popup automatically?"
      Top             =   4860
      Width           =   375
   End
   Begin VB.Timer MinuteTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4380
      Top             =   2100
   End
   Begin VB.TextBox txtActivity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      ToolTipText     =   "Type here whatever you want!"
      Top             =   1860
      Width           =   5115
   End
   Begin VB.CheckBox chkUnload 
      Caption         =   "&Unload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Check this if you want to stop this program"
      Top             =   4260
      Width           =   915
   End
   Begin VB.TextBox txtDSN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3540
      TabIndex        =   11
      ToolTipText     =   "ODBC DSN where your entries are saved"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -60
      TabIndex        =   14
      Top             =   4620
      Width           =   5715
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -60
      TabIndex        =   13
      Top             =   1260
      Width           =   5715
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Close (Esc)"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Save and close the window (Ctrl + Enter)"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblMakeStatusReport 
      AutoSize        =   -1  'True
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   29
      ToolTipText     =   "Make the report for today's work (Ctrl + R)"
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblHints 
      AutoSize        =   -1  'True
      Caption         =   "&&"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2940
      TabIndex        =   27
      ToolTipText     =   "Get hints on what you were doing (F1)"
      Top             =   4800
      Width           =   435
   End
   Begin VB.Label Label8 
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5160
      TabIndex        =   26
      Top             =   4875
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "&Time Spent:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   25
      Top             =   4875
      Width           =   855
   End
   Begin VB.Label lblDeleteTask 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5340
      TabIndex        =   23
      ToolTipText     =   "Delete Task (Press Alt and X)"
      Top             =   730
      Width           =   135
   End
   Begin VB.Label lblModifyTask 
      AutoSize        =   -1  'True
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5050
      TabIndex        =   22
      ToolTipText     =   "Modify Task (Press Alt and ~)"
      Top             =   660
      Width           =   165
   End
   Begin VB.Label lblAddTask 
      AutoSize        =   -1  'True
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4770
      TabIndex        =   21
      ToolTipText     =   "Add New Task (Press Alt and +)"
      Top             =   690
      Width           =   165
   End
   Begin VB.Label lblAddProject 
      AutoSize        =   -1  'True
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4770
      TabIndex        =   20
      ToolTipText     =   "Add New Project (Press Alt and +)"
      Top             =   120
      Width           =   165
   End
   Begin VB.Label lblModifyProject 
      AutoSize        =   -1  'True
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5060
      TabIndex        =   18
      ToolTipText     =   "Modify Project (Press Alt and ~)"
      Top             =   180
      Width           =   165
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Current Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3660
      TabIndex        =   17
      Top             =   4260
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tas&k:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Next &Interval:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   4870
      Width           =   240
   End
   Begin VB.Label lblDisplayTime 
      AutoSize        =   -1  'True
      Caption         =   "00:00 AM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4680
      TabIndex        =   15
      Top             =   4260
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   10
      Top             =   1515
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Project:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "This caption is set at run time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblDeleteProject 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5340
      TabIndex        =   19
      ToolTipText     =   "Delete Project (Press Alt and X)"
      Top             =   200
      Width           =   135
   End
End
Attribute VB_Name = "frmWakeUpDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Prev_Unselected_Text_cboProject As String
Private Prev_Unselected_Text_cboTask As String
Private mnPrev_Project_List_Index As Integer


Const msLBL_COMMENT_CAPTION As String = "&What were you doing "

Const mlNO_RESPONCE_TIMEOUT As Long = 1

Private mbUserPressedCancel As Boolean
Private mnLastKey As Integer
Private mbIsNewLine As String
Private mdtLastSave As Date
Private mbHandleLostFocus As Boolean
Private mbUserEditedTimeSpentField As Boolean
Private mbNotesChanged As Boolean

Private msLastActivityText As String
Private mbUserWasAway As Boolean
Private mbUserResponded As Boolean
Private mlMinutesCount As Long

Private mbFormDirty As Boolean

Private mbInitialisedOnce As Boolean
Private mlInitialInterval As Long
Private mlEneteredTaskTotalMinutes As Long 'Stores total of time spent for tasks entered by user after dialog was poped up


Public Function DisplayForm(ByRef rsActivity As String, ByRef rlInterval As Long, ByRef rlProjectID As Long, ByRef rlTaskID As Long, ByVal vdtLastSaved As Date, ByRef rbUnload As Boolean, ByRef rlTimeSpent As Long) As Boolean

    Dim bSuccess As Boolean
    
    'Initialize the form
    Call InitForm(rlProjectID, rlTaskID, vdtLastSaved, rlInterval, rsActivity, rbUnload)
   
    Do
    
        'bSuccess stores whether all inputes are correct if user pressed OK
        bSuccess = True
    
        'Make the form top most on all other
        MakeFormTopMost Me
        
        'Update the time in label caption every time it pops up
        UpdateLabelCaption
        
        'Show the form modally
        Me.Show vbModal
        
        'If form is not dirty
        
        '
        'Below is commented because even if form is not dirty, some field may be having invalid init values (currently Time Spent would have invalid number after pressing Save This button)
        '
        
'        If Not mbFormDirty Then
'
'            'If form is not dirty act as if cancel was pressed
'            mbUserPressedCancel = True
'
'        End If
        
        'If user pressed OK
        If (Not mbUserPressedCancel) Then
        
            'Validate if all inputes are proper
            bSuccess = ValidateInputes
            
            'If all inputes are proper
            If bSuccess Then
            
                'Move data from form to output paramaters
                Call SetReturnData(rsActivity, gsDSN, rlInterval, rlProjectID, rlTaskID, rlTimeSpent)
                
            End If
            
        End If
        
    'If validation failed, show up the form again
    Loop Until bSuccess
    
    'Update rbUnload var even if user pressed Cancel
    rbUnload = (chkUnload.Value = vbChecked)

    'Return whether user pressed OK or not
    DisplayForm = (Not mbUserPressedCancel)
    
    'Finally unload the form
    Unload Me

End Function

Private Sub btnCancel_Click()
    
    mbUserPressedCancel = True
    
    Me.Hide
    
    DoEvents

End Sub

Private Sub btnCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    mbHandleLostFocus = False
    
End Sub


Private Sub btnSaveThis_Click()
    
    Dim sActivity As String
    Dim sDSN As String
    Dim lInterval As Long
    Dim lProjectID As Long
    Dim lTaskID As Long
    Dim lTimeSpent As Long
    Dim bSuccess As Boolean
    
    bSuccess = True
        
    'Validate if all inputes are proper
    bSuccess = ValidateInputes
    
    'If all inputes are proper
    If bSuccess Then
    
        If bSuccess Then
        
            'Move data from form to output paramaters
            Call SetReturnData(sActivity, sDSN, lInterval, lProjectID, lTaskID, lTimeSpent)
            
            'Save the activity as entered in dialog
            bSuccess = SaveActivity(sActivity, Now, glUser_ID, GetMachineName, lTaskID, lTimeSpent)
            
            If bSuccess Then
            
                'One task is enter. Force user to enter new time for next time he want to enter.
                'Note: this will fire the change event and will make form dirty
                
                mlEneteredTaskTotalMinutes = mlEneteredTaskTotalMinutes + lTimeSpent
                
                Dim lMinutesRemained As Long
                
                lMinutesRemained = MinutesPassedSinceLastSave - mlEneteredTaskTotalMinutes
                
                If lMinutesRemained > 0 Then
                
                    txtTimeSpent.Text = "?" & lMinutesRemained
                
                Else
                
                    txtTimeSpent.Text = "?"
                
                End If
                
                Call SetDirtyFlag(False)
                
            End If
            
        End If
            
    End If

End Sub

Private Sub lblHints_Click()
    
    Dim sWinList As String
    Dim i As Integer
    
    If MinutesPassedSinceLastSave <> 0 Then
    
        sWinList = "The following programs have been running during the last " & MinutesPassedSinceLastSave & " minutes" & ":" & vbCrLf & vbCrLf
        
    Else
    
        sWinList = "The following programs have been running during the last 1 minute:" & vbCrLf & vbCrLf
    
    End If
    
    If GetDimension(gavntWinList) = 2 Then
    
        For i = LBound(gavntWinList, 2) To UBound(gavntWinList, 2)
        
            sWinList = sWinList & gavntWinList(2, i) & vbCr
        
        Next i
        
    End If
    
    MsgBox sWinList
    
End Sub

Private Sub btnOK_Click()
    
    mbUserPressedCancel = False
    
    Me.Hide
    
    DoEvents

End Sub



Private Sub cboProject_Change()
    
    'If user didn't changed the notes, clear it
    If Not mbNotesChanged Then
    
        txtActivity = ""
        
    End If

    If IE4LikeCombo_Change(cboProject, Prev_Unselected_Text_cboProject) Then
    
        cboTask.Clear
        
        mnPrev_Project_List_Index = -1
    
    End If
    
    Call SetDirtyFlag(True)
    
End Sub

Private Sub cboProject_Click()
       
    'If last selected index is not equal to currently selected index then only update the task combo
    If mnPrev_Project_List_Index <> cboProject.ListIndex Then
        
        'There was new selection in Project combo, so refill the task list
        Call SetupTaskCombo(-1, False)
        
        mnPrev_Project_List_Index = cboProject.ListIndex
        
    End If
    
    Call SetDirtyFlag(True)
    
End Sub

Private Sub cboProject_LostFocus()

    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String

    If mbHandleLostFocus And (Not IsItemExistInCombo(cboProject)) Then
    
        nSecurityResult = CheckSecurity(gnMODE_PROJECT_EDITOR_ADD, sSecurityAccessDeniyMsg)
        
        If nSecurityResult = gnSECURITY_DENIY Then
        
            MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to do this task.")
            
            Call NoFailSelectItemInCombo(cboProject, 0)
            
        Else
        
            If MsgBox("Do you want to add new project named " & AlternateStrIfNull(cboProject.Text, "<New Project>") & " in your list?", vbYesNo) = vbYes Then
            
                Call ShowProjectTaskEditor(gnMODE_PROJECT_EDITOR_ADD, cboProject.Text)
                        
'If user does not creates new project don't do any thing
                        
'            Else
'
'                If cboProject.ListCount <> 0 Then
'
'                    cboProject.ListIndex = 0
'
'                End If
'
            End If
            
        End If
    
    End If
        

    
End Sub

Private Sub cboTask_Change()

    'If user didn't changed the notes, clear it
    If Not mbNotesChanged Then
    
        txtActivity = ""
        
    End If


   Call IE4LikeCombo_Change(cboTask, Prev_Unselected_Text_cboTask)
   
   Call SetDirtyFlag(True)
   
End Sub

Private Sub cboTask_LostFocus()

    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String

    If mbHandleLostFocus And (Not IsItemExistInCombo(cboTask)) And (Len(cboTask.Text) <> 0) Then
    
        'Is valid project exist in combo?
        If IsItemExistInCombo(cboProject) Then
        
            nSecurityResult = CheckSecurity(gnMODE_TASK_EDITOR_ADD, sSecurityAccessDeniyMsg)
            
            If nSecurityResult = gnSECURITY_DENIY Then
            
                MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to do this task.")
                
                Call NoFailSelectItemInCombo(cboTask, 0)
                
            Else
        
                If MsgBox("Do you want to add new task named " & AlternateStrIfNull(cboTask.Text, "<New Task>") & " in your list?", vbYesNo) = vbYes Then
                
                    Call ShowProjectTaskEditor(gnMODE_TASK_EDITOR_ADD, cboTask.Text)
                    
'If user does not creates new task don't do any thing

'                Else
'
'                    If cboTask.ListCount <> 0 Then
'
'                        cboTask.ListIndex = 0
'
'                    End If
                            
                End If
                
            End If
            
        Else
        
            MsgBox "The task you entered will not be added untill you first create the valid project."
        
        End If
    
    End If


End Sub

Private Function ValidateInputes() As Boolean

    Dim bSuccess As Boolean
    Dim sMsg As String
    
    bSuccess = True
    
    sMsg = "The following values are not correct:" & vbCr

    If Not IsNumeric(txtInterval.Text) Then
    
        sMsg = sMsg & "Interval value " & txtInterval.Text & " is not a valid number." & vbCr
        
        bSuccess = False
        
    End If
    
    
    If txtDSN.Text = "" Then
    
        sMsg = sMsg & "You must enter valid DSN set to your database." & vbCr
        
        bSuccess = False
    
    End If
    
  
    If Not IsNumber(txtTimeSpent.Text) Then
    
        sMsg = sMsg & "You must enter the valid number in the Time Spent field." & vbCr
        
        bSuccess = False
    
    End If
    
    If Len(cboTask.Text) <> 0 Then
    
        If Not IsItemExistInCombo(cboTask) Then
        
            'Execute the LostFocus
            cboTask_LostFocus
            
            'Is still there a task in combo not added to database
            If Not IsItemExistInCombo(cboTask) Then
            
                sMsg = sMsg & "The task you have specified does not exist in database. You must first add it in to database." & vbCr
                
                bSuccess = False
            
            End If
                
        End If
        
    Else
        
        sMsg = sMsg & "Task name can not be blank. Please select a task from the list or create a new one." & vbCr
        
        bSuccess = False
    
    End If
    
    If Len(cboProject.Text) <> 0 Then
    
        If Not IsItemExistInCombo(cboProject) Then
        
            'Execute the LostFocus
            cboProject_LostFocus
            
            'Is still there a project in combo not added to database
            If Not IsItemExistInCombo(cboProject) Then
            
                sMsg = sMsg & "The project you have specified does not exist in database. You must first add it in to database." & vbCr
                
                bSuccess = False
            
            End If
                
        End If
        
    Else
        
        sMsg = sMsg & "Project name can not be blank. Please select a project from the list or create a new one." & vbCr
        
        bSuccess = False
    
    End If
    
    'If other things are correct
    If bSuccess Then
    
        'Warn user about too large value of time spent
        Dim lTimeSpent As Long
        Dim lInterval As Long
    
        lTimeSpent = CLng(txtTimeSpent)
        lInterval = CLng(txtInterval)
    
        If lTimeSpent <= 0 Then
        
            bSuccess = False
            
            sMsg = sMsg & "Time Spent value specified must be number greater than 0." & vbCr
        
        'If user didn't edited time spent field and timespent is greater then interval, ask him if he is sure
        ElseIf (Not mbUserEditedTimeSpentField) And ((lTimeSpent - lInterval) > 5) Then
        
            If MsgBox("Are you sure you spent " & txtTimeSpent & " min on " & cboTask & "?", vbYesNo) <> vbYes Then
            
                bSuccess = False
                
                sMsg = sMsg & "Time Spent value needs to be changed." & vbCr
            
            End If
            
        'Else if time spent is too large, give warning
        ElseIf lTimeSpent > (lInterval * 3) Then
            
            If MsgBox("Are you sure you spent " & txtTimeSpent & " min on " & cboTask & "?", vbYesNo) <> vbYes Then
            
                bSuccess = False
                
                sMsg = sMsg & "Time Spent specified needs to be changed." & vbCr
            
            End If
            
        End If
        
    End If
    

    
    If Not bSuccess Then
        
        MsgBox sMsg
        
    End If
    
    ValidateInputes = bSuccess

End Function


Private Sub Form_Activate()

    If Not mbInitialisedOnce Then
    
        cboProject.SetFocus
        
        mbInitialisedOnce = True
    
    End If
    
    mbHandleLostFocus = True

End Sub

Private Sub Form_Deactivate()
    
    ResetLabelActivations

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    mbUserResponded = True

    If Not mbUserWasAway Then
    
        'Is this Ctrl+Enter
        If ((KeyCode = vbKeyReturn) And ((Shift And vbCtrlMask) > 0)) Then
        
            'Let the active control lost the focus so that if new task/project is in combo user will be asked to add it first
            btnOK.SetFocus
            
            btnOK_Click
            
        End If
        
        'Is this F2 or Ctrl+S
        If (KeyCode = vbKeyF2) Or ((KeyCode = vbKeyS) And ((Shift And vbCtrlMask) > 0)) Then
        
            'Let the active control lost the focus so that if new task/project is in combo user will be asked to add it first
            If btnSaveThis.Enabled Then
                
                btnSaveThis.SetFocus
                
            Else
            
                btnOK.SetFocus
                
            End If
            
            'Let the lost focus event occure
            DoEvents
            
            btnSaveThis_Click
        
        End If
        
        
        'Is this F1 or Alt+H
        If (KeyCode = vbKeyF1) Or ((KeyCode = vbKeyH) And ((Shift And vbAltMask) > 0)) Then
        
            'Show hints
            lblHints_Click
        
        End If
        
        'Is this Ctrl + R
        If ((KeyCode = vbKeyR) And ((Shift And vbCtrlMask) > 0)) Then
        
            'Make the report
            lblMakeStatusReport_Click
            
        End If
        
    End If
    
    'Is user asking to refresh
    If KeyCode = vbKeyF5 Then
    
        DoRefresh
    
    End If
    
    '-----------------------------
    'Check if shortcut key pressed
    '-----------------------------
    
    'Is ALt key pressed
    If (Shift And vbAltMask) > 0 Then
    
        Select Case KeyCode
        
            Case 187        'Plus key: + or =
            
                If Me.ActiveControl Is cboProject Then
                
                    lblAddProject_Click
                
                ElseIf Me.ActiveControl Is cboTask Then
                
                    lblAddTask_Click
                
                End If
            
            Case 192        'Tilde key: ~
            
                If Me.ActiveControl Is cboProject Then
                
                    lblModifyProject_Click
                
                ElseIf Me.ActiveControl Is cboTask Then
                
                    lblModifyTask_Click
                
                End If
            
            
            Case vbKeyX
            
                If Me.ActiveControl Is cboProject Then
                
                    lblDeleteProject_Click
                
                ElseIf Me.ActiveControl Is cboTask Then
                
                    lblDeleteTask_Click
                
                End If
            
        
        End Select
        
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        
    If mbUserWasAway Then
        
        MakeUserUnAway
        
        If Me.ActiveControl Is txtActivity Then
            
            KeyAscii = 0
            
        End If
        
    End If
    
    'Is this Enter but focus is not on Activity text box?
    If KeyAscii = 10 Then
    
        If Not (Me.ActiveControl Is txtActivity) Then
        
            KeyAscii = 0
        
        End If
    
    End If

End Sub



Private Sub Form_Load()

    mbInitialisedOnce = False

End Sub

Private Sub Form_LostFocus()
    
    ResetLabelActivations
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetLabelActivations
    
    If mbUserWasAway Then
        
        MakeUserUnAway
        
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mbUserPressedCancel = True
End Sub


Private Sub lblAddProject_Click()
    
    If IsItemExistInCombo(cboProject) Then
    
        Call ShowProjectTaskEditor(gnMODE_PROJECT_EDITOR_ADD, "<New Project>")
        
    Else
    
        Call ShowProjectTaskEditor(gnMODE_PROJECT_EDITOR_ADD, cboProject.Text)
        
    End If

End Sub

Private Sub lblAddProject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblAddProject, True)

End Sub


Private Sub lblAssTask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetLabelStyle(lblAddTask, True)

End Sub



Private Sub lblAddTask_Click()
    
    If IsItemExistInCombo(cboTask) Then
    
        Call ShowProjectTaskEditor(gnMODE_TASK_EDITOR_ADD, "<New Task>")
        
    Else
    
        Call ShowProjectTaskEditor(gnMODE_TASK_EDITOR_ADD, cboTask.Text)
        
    End If

End Sub

Private Sub lblAddTask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblAddTask, True)

End Sub

Private Sub lblDeleteProject_Click()
    
    Dim bSuccess As Boolean
    Dim nSelectedItemIndexInCombo As Integer
    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String
    Dim lProjectID As Long
    Dim vntProjectForUser As Variant
    Dim lRowCount As Long
    Dim sSQL As String
    

    bSuccess = True
        
    nSecurityResult = CheckSecurity(gnMODE_PROJECT_EDITOR_DELETE, sSecurityAccessDeniyMsg)
    
    If nSecurityResult = gnSECURITY_DENIY Then
    
        MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to perform this action.")
        
    Else
    
        If IsItemExistInCombo(cboProject) Then
        
            If MsgBox("Are you sure you want to delete the project " & cboProject.Text & " from your list?", vbYesNoCancel) = vbYes Then
                
                lProjectID = GetSelectedItemDataInCombo(cboProject)
                
                'Delete the membership from the project
                bSuccess = DeleteRows("Task_Category_Memberships", Array("Task_Category_ID", "User_ID"), Array(lProjectID, glUser_ID))
                
                'Check if it is really deleted
                sSQL = "SELECT *"
                sSQL = sSQL & " FROM Task_Category_Memberships"
                sSQL = sSQL & " Where (Task_Category_ID = " & lProjectID & ")"
                sSQL = sSQL & " AND ((User_ID = " & glUser_ID & ")"
                sSQL = sSQL & " OR (User_ID = " & glGlobal_User_ID & "))"
                
                Call GetRowsInArr(sSQL, vntProjectForUser, lRowCount)
                
                If IsEmpty(vntProjectForUser) Then
                
                    If bSuccess Then
                
                        nSelectedItemIndexInCombo = GetSelectedItemInCombo(cboProject)
                        
                        'Delete from combo too
                        Call cboProject.RemoveItem(nSelectedItemIndexInCombo)
                        
                        Call NoFailSelectItemInCombo(cboProject, nSelectedItemIndexInCombo + 1, nSelectedItemIndexInCombo - 1)
                        
                    End If
                    
                Else
                
                    MsgBox "The project can not be removed because it is allocated globally to every user."

                End If
            
            End If
            
        Else
        
            Call NoFailSelectItemInCombo(cboProject, 0)
            
        End If
        
    End If

End Sub

Private Sub lblDeleteProject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetLabelStyle(lblDeleteProject, True)
    
End Sub

Private Sub lblDeleteTask_Click()
    
    Dim bSuccess As Boolean
    Dim nSelectedItemIndexInCombo As Integer
    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String
    Dim sSQL As String
    Dim lTaskID As Long
    Dim vntTaskForUser As Variant
    Dim lRowCount As Long

    bSuccess = True
        
    nSecurityResult = CheckSecurity(gnMODE_TASK_EDITOR_DELETE, sSecurityAccessDeniyMsg)
    
    If nSecurityResult = gnSECURITY_DENIY Then
    
        MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to perform this action.")
        
    Else
    
        If IsItemExistInCombo(cboTask) Then
        
            If MsgBox("Are you sure you want to delete the task " & cboTask.Text & " from your list?", vbYesNoCancel) = vbYes Then
            
                lTaskID = GetSelectedItemDataInCombo(cboTask)
                
                'Delete the membership from the task
                bSuccess = DeleteRows("Task_Memberships", Array("Task_ID", "User_ID"), Array(lTaskID, glUser_ID))
                
                'Check if it is really deleted
                sSQL = "SELECT *"
                sSQL = sSQL & " FROM Task_Memberships"
                sSQL = sSQL & " Where (Task_ID = " & lTaskID & ")"
                sSQL = sSQL & " AND ((User_ID = " & glUser_ID & ")"
                sSQL = sSQL & " OR (User_ID = " & glGlobal_User_ID & "))"
                
                Call GetRowsInArr(sSQL, vntTaskForUser, lRowCount)
                
                If IsEmpty(vntTaskForUser) Then
                                
                    If bSuccess Then
                
                        nSelectedItemIndexInCombo = GetSelectedItemInCombo(cboTask)
                        
                        'Delete from combo too
                        Call cboTask.RemoveItem(nSelectedItemIndexInCombo)
                        
                        Call NoFailSelectItemInCombo(cboTask, nSelectedItemIndexInCombo + 1, nSelectedItemIndexInCombo - 1)
                        
                    End If
                    
                Else
                
                    MsgBox "The task can not be removed because it is allocated globally to every user."
                
                End If
            
            End If
            
        Else
        
            Call NoFailSelectItemInCombo(cboTask, 0)
            
        End If
        
    End If

End Sub

Private Sub lblDeleteTask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblDeleteTask, True)
    
End Sub



Private Sub lblHints_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetLabelStyle(lblHints, True)

End Sub

Private Sub lblMakeStatusReport_Click()
    
    btnCancel_Click
    
    ShowStatusReportDialog
    
End Sub

Private Sub lblMakeStatusReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetLabelStyle(lblMakeStatusReport, True)

End Sub

Private Sub lblModifyProject_Click()

    If IsItemExistInCombo(cboProject) Then
    
        Call ShowProjectTaskEditor(gnMODE_PROJECT_EDITOR_MODIFY, cboProject.Text, GetSelectedItemDataInCombo(cboProject))
        
    Else
    
        Call ShowProjectTaskEditor(gnMODE_PROJECT_EDITOR_ADD, cboProject.Text)
        
    End If
    
End Sub

Private Sub lblModifyProject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetLabelStyle(lblModifyProject, True)
    
End Sub


Private Sub lblModifyTask_Click()
    
    If IsItemExistInCombo(cboTask) Then
    
        Call ShowProjectTaskEditor(gnMODE_TASK_EDITOR_MODIFY, cboTask.Text, GetSelectedItemDataInCombo(cboTask))
        
    Else
    
        Call ShowProjectTaskEditor(gnMODE_TASK_EDITOR_ADD, cboTask.Text)
        
    End If

End Sub

Private Sub lblModifyTask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblModifyTask, True)
    
End Sub

Private Sub MinuteTimer_Timer()
        
    mlMinutesCount = mlMinutesCount + 1
    
    UpdateLabelCaption
    
    If (Not mbUserResponded) And (mlMinutesCount = mlNO_RESPONCE_TIMEOUT) Then
    
        'Replace the last activity text because user is away and others should not see what he was doing last time
        msLastActivityText = txtActivity.Text
        
        txtActivity.Text = "User is away." & vbCrLf & "Press any key to see" & vbCrLf & "last activity text."
        
        txtActivity.SelStart = 0
        txtActivity.SelLength = Len(txtActivity.Text)
        
        mbUserWasAway = True
    
    End If
    
End Sub



Private Sub txtActivity_Change()

    Call SetDirtyFlag(True)
    
    mbNotesChanged = True
    
End Sub

Private Sub txtActivity_KeyPress(KeyAscii As Integer)
    
    If mbIsNewLine Then
    
        If (Chr$(KeyAscii) >= "a") And (Chr$(KeyAscii) <= "z") Then
        
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        
        End If
        
        mbIsNewLine = False
    
    End If
    
    If KeyAscii = 13 Then
    
        mbIsNewLine = True
    
    End If
    
    mnLastKey = KeyAscii
    
End Sub

Private Sub UpdateLabelCaption()

    Dim lTimeSpent As Long
    Dim bFormDirtyState As Boolean
    Dim sTimeSpentInTextBox As String

    lTimeSpent = MinutesPassedSinceLastSave

    'If time spent is 0 then make it minimal
    If lTimeSpent = 0 Then
    
        lTimeSpent = 1
        
        lblComment.Caption = msLBL_COMMENT_CAPTION & "since last " & lTimeSpent & " minute:"
        
        sTimeSpentInTextBox = CStr(lTimeSpent)
    
    'If time spent is negative do something corrective
    ElseIf lTimeSpent < 0 Then
    
        lTimeSpent = -1
        
        lblComment.Caption = "Your system date is not correct!"
        
        lblComment.ForeColor = vbRed
        
        sTimeSpentInTextBox = "???"
    
    'If time spent is too much do something
    ElseIf lTimeSpent > (mlInitialInterval * 10) Then
    
        lblComment.Caption = msLBL_COMMENT_CAPTION & "in the last " & lTimeSpent & " minutes (so much time!)?"
        
        sTimeSpentInTextBox = "?" & CStr(lTimeSpent)
    
    'Otherwise timespent is correct
    Else
    
        lblComment.Caption = msLBL_COMMENT_CAPTION & "in the last " & lTimeSpent & " minutes?"
        
        sTimeSpentInTextBox = CStr(lTimeSpent)
    
    End If
    
    
    If Not mbUserEditedTimeSpentField Then
    
        'Save the current dirty status
        bFormDirtyState = mbFormDirty
            
            'Update the edit field
            txtTimeSpent.Text = sTimeSpentInTextBox
            
            mbUserEditedTimeSpentField = False
        
        'Restore the original dirty status
        SetDirtyFlag (bFormDirtyState)
        
    End If
    
    'Update the clock in label
    lblDisplayTime.Caption = FormattedTime
    
    'Millitary time has smaller width - so position the label according to it's current width
    lblDisplayTime.Left = txtActivity.Left + txtActivity.Width - lblDisplayTime.Width

End Sub

Private Sub DoRefresh()
    
    Dim lSelectedProjectId As Long
    Dim lSelectedTaskId As Long

    'Save the last Project ID which was selected
    If cboProject.ListIndex <> -1 Then
    
        lSelectedProjectId = cboProject.ItemData(cboProject.ListIndex)
        
    Else
    
        lSelectedProjectId = -1
    
    End If
    
    'Save the last Project ID which was selected
    If cboTask.ListIndex <> -1 Then
    
        lSelectedTaskId = cboTask.ItemData(cboTask.ListIndex)
        
    Else
    
        lSelectedTaskId = -1
    
    End If
    
    'Now fillup the project and task combos and try to select last selected items
    Call SetupProjectCombo(lSelectedProjectId)
    
    Call SetupTaskCombo(lSelectedTaskId, False)
    
End Sub

Private Sub SetupProjectCombo(ByVal vlProjectIDToBeSelected As Long)

    Dim sSQL As String
    
    sSQL = "SELECT DISTINCT Task_Categories.Task_Category_ID, Task_Categories.Task_Category_Name"
    sSQL = sSQL & " FROM Task_Categories, Task_Category_Memberships"
    sSQL = sSQL & " WHERE (Task_Categories.Task_Category_ID = Task_Category_Memberships.Task_Category_ID)"
    sSQL = sSQL & " AND ((Task_Category_Memberships.User_ID = " & glUser_ID & ")"
    sSQL = sSQL & " OR (Task_Category_Memberships.User_ID = " & glGlobal_User_ID & "))"
    
    
    
    Call PopulateCombo(cboProject, sSQL)
    
    Call SelectItemInCombo(cboProject, vlProjectIDToBeSelected)
    
    Call SetProperComboWidth(cboProject)

End Sub

Private Sub SetupTaskCombo(ByVal vlTaskIDToBeSelected As Long, ByVal vbFillOnlyIfEmpty As Boolean)

    Dim bFillTaskCombo As Boolean
    Dim sSQL As String

    'If something is selected in Project combo then only tasks could be filled
    If cboProject.ListIndex <> -1 Then
    
        'First detremine whether it is required to fill the task combo
        bFillTaskCombo = (Not vbFillOnlyIfEmpty) Or (vbFillOnlyIfEmpty And (cboTask.ListCount = 0))
        
        'If it is required to fill the task combo then
        If bFillTaskCombo Then
            
            'Fill the task list according to project ID
            sSQL = "SELECT DISTINCT Tasks.Task_ID, Tasks.Task_Name, Tasks.Created_On"
            sSQL = sSQL & " FROM Tasks, Task_Memberships"
            sSQL = sSQL & " WHERE (Tasks.Task_ID = Task_Memberships.Task_ID)"
            sSQL = sSQL & " AND (Tasks.Task_Category_ID = " & cboProject.ItemData(cboProject.ListIndex) & ")"
            sSQL = sSQL & " AND ((Task_Memberships.User_ID = " & glUser_ID & ")"
            sSQL = sSQL & " OR (Task_Memberships.User_ID = " & glGlobal_User_ID & "))"
            sSQL = sSQL & " AND (Tasks.Is_Visible = True)"
            If frmMain.mnuDoNotShowCompletedItems.Checked Then
            
                sSQL = sSQL & " AND (Tasks.Is_Completed = False)"
            
            End If
            sSQL = sSQL & " ORDER BY Tasks.Created_On DESC"
            
            Call PopulateCombo(cboTask, sSQL)
        
            'If task list got filled, by default select first task
            If cboTask.ListCount <> 0 Then
            
                cboTask.ListIndex = 0
            
            End If
            
        End If
        
        'Try to select the task ID that was passed
        Call SelectItemInCombo(cboTask, vlTaskIDToBeSelected)
        
    Else
    
        'If nothing is selected in Project combo then clear the task list
        cboTask.Clear
        
    End If
    
    Call SetProperComboWidth(cboTask)

End Sub

Private Function InitForm(ByVal vlProjectID As Long, ByVal vlTaskID As Long, ByVal vdtLastSaved As Date, ByVal vlInterval As Long, ByVal vsActivity As String, ByRef rbUnload As Boolean) As Boolean
    
    mbUserEditedTimeSpentField = False
    
    'This var stores last value of ListIndex in Project combo.
    mnPrev_Project_List_Index = -2
    
    'Show the version in the caption
    Call AddVerInCaption(Me)
    
    'Get the list of projects and select the project that was selected last time
    Call SetupProjectCombo(vlProjectID)
    
    'Set up the task list according to project combo box selection
    Call SetupTaskCombo(vlTaskID, True)
    
    rbUnload = False
    
    mnLastKey = 0
    
    mlMinutesCount = 0
    
    mbUserWasAway = False
    
    mbUserResponded = False
    
    mbIsNewLine = True
    
    mdtLastSave = vdtLastSaved
    
    chkUnload.Value = vbUnchecked
    
    txtInterval.Text = CStr(vlInterval)
    
    mlInitialInterval = vlInterval
    
    txtDSN.Text = gsDSN
    
    txtActivity.Text = vsActivity
    
    txtActivity.SelStart = 0
    
    txtActivity.SelLength = Len(txtActivity.Text)
    
    'Set the flag that user haven't changed the notes
    mbNotesChanged = False
    
    'Enable the minute timer, on firing of which labels will be updated when the form is modal
    MinuteTimer.Enabled = True
    
    mbHandleLostFocus = True
    
    mlEneteredTaskTotalMinutes = 0
    
    'Allow user to re-enter same activity when popup
    Call SetDirtyFlag(True)
    
End Function


Function MinutesPassedSinceLastSave() As Long

    MinutesPassedSinceLastSave = DateDiff("n", mdtLastSave, Now)

End Function

Function ShowProjectTaskEditor(ByVal vnMode As Integer, Optional ByRef rsProjectTaskName As String, Optional ByVal vlProjectTaskID As Long) As Boolean

    Dim avntMembers As Variant
    Dim avntAllItems As Variant
    Dim avntMemberModules As Variant
    Dim avntAllModules As Variant
    Dim avntProjectTaskData As Variant
    Dim lTaskTypeID As Long
    Dim lProjectID As Long
    Dim lRowCount As Long
    Dim nTaskPriority As Variant
    Dim nTaskTimeAllocated As Variant
    Dim dtTastExpectedStartTime As Variant
    Dim sSQL As String
    Dim bSuccess As Boolean
    Dim ofrmProjectTaskEditor As frmProjectTaskEditor
    Dim bIsCompleted As Boolean
    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String


    bSuccess = True
    
    nSecurityResult = CheckSecurity(vnMode, sSecurityAccessDeniyMsg)
    
    If nSecurityResult = gnSECURITY_DENIY Then
    
        MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to do this task.")
        
    Else

        Select Case vnMode
        
            Case gnMODE_PROJECT_EDITOR_ADD
            
                'What's going on:
                'ProjectTask editor's DisplayForm takes two array: one for member users and other for all availabel users.
                'In Add mode, member user is empty and All user is filled with all available user.
                'So fire the query to get array of user_id and login name and pass it to DisplayForm.
                '
                        
                avntMembers = Empty
                
                avntAllItems = Empty
                
                'Get list of all users
                sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE Is_Visible = True"
                
                Call GetRowsInArr(sSQL, avntAllItems, lRowCount)
                
                'Set default user
                sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE User_ID = " & glUser_ID
                
                Call GetRowsInArr(sSQL, avntMembers, lRowCount)
                
                'Before showing modal form make this form as normal form (which is otherwise topmost form)
                MakeFormNonTopMost Me
                
                'Don't handle event for LostFocus of txtProject and txtTask
                mbHandleLostFocus = False
                
                Set ofrmProjectTaskEditor = New frmProjectTaskEditor
                
                'Now display the form
                bSuccess = ofrmProjectTaskEditor.DisplayForm(vnMode, nSecurityResult, -1, -1, rsProjectTaskName, avntMembers, avntAllItems)
                
                Set ofrmProjectTaskEditor = Nothing
                
                'Start hadling LostFocus as usual
                mbHandleLostFocus = True
                
                'Make the form top most again
                MakeFormTopMost Me
                
                If bSuccess Then
                
                    'If user pressed OK then add the entry in database
                    bSuccess = AddNewProject(rsProjectTaskName, avntMembers)
                    
                End If
                
            Case gnMODE_PROJECT_EDITOR_MODIFY
            
                avntMembers = Empty
                
                avntAllItems = Empty
                
                sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE Is_Visible = True"
                
                Call GetRowsInArr(sSQL, avntAllItems, lRowCount)
                
                sSQL = "SELECT Users.User_ID, Users.User_Login_Name"
                sSQL = sSQL & " FROM Task_Category_Memberships, Users"
                sSQL = sSQL & " WHERE (Task_Category_Memberships.Task_Category_ID = " & vlProjectTaskID & ")"
                sSQL = sSQL & " AND (Task_Category_Memberships.User_ID = Users.User_ID)"
                sSQL = sSQL & " AND (Is_Visible = True)"
                
                Call GetRowsInArr(sSQL, avntMembers, lRowCount)
                
                'Before showing modal form make this form as normal form (which is otherwise topmost form)
                MakeFormNonTopMost Me
                
                'Don't handle event for LostFocus of txtProject and txtTask
                mbHandleLostFocus = False
                
                
                Set ofrmProjectTaskEditor = New frmProjectTaskEditor
                
                'Now display the form
                bSuccess = ofrmProjectTaskEditor.DisplayForm(vnMode, nSecurityResult, vlProjectTaskID, -1, rsProjectTaskName, avntMembers, avntAllItems)
                
                Set ofrmProjectTaskEditor = Nothing
                
                'Start hadling LostFocus as usual
                mbHandleLostFocus = True
                
                'Make the form top most again
                MakeFormTopMost Me
                
                If bSuccess Then
                
                    'If user pressed OK then add the entry in database
                    bSuccess = ModifyProject(vlProjectTaskID, rsProjectTaskName, avntMembers)
                    
                End If
        
            Case gnMODE_TASK_EDITOR_ADD
            
                'Is valid project exist?
                If IsItemExistInCombo(cboProject) Then
                                
                    'Set default task type
                    lTaskTypeID = -1
                    
                    'Set user memberships
                    avntMembers = Empty
                    
                    avntAllItems = Empty
                    
                    'Get list of all users
                    sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE Is_Visible = True"
                    
                    Call GetRowsInArr(sSQL, avntAllItems, lRowCount)
                    
                    'Set default user
                    sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE (User_ID = " & glUser_ID & ") AND (Is_Visible = True)"
                    
                    Call GetRowsInArr(sSQL, avntMembers, lRowCount)
                    
                    'Set module memberships
                    avntMemberModules = Empty
                    
                    avntAllModules = Empty
                    
                    'Get all the sub-modules
                    'First get current project ID
                    lProjectID = GetSelectedItemDataInCombo(cboProject)
                    
                    'Get all submodules associated with project
                    sSQL = "SELECT Sub_Modules.Sub_Module_ID, Modules.Module_Name & '.' & Sub_Modules.Sub_Module_Name"
                    sSQL = sSQL & " FROM Sub_Modules, Modules"
                    sSQL = sSQL & " WHERE (Modules.Module_ID = Sub_Modules.Module_ID)"
                    sSQL = sSQL & " AND (Modules.Task_Category_ID =" & lProjectID & ")"
                    sSQL = sSQL & " AND (Sub_Modules.Is_Visible = True)"
                    
                    Call GetRowsInArr(sSQL, avntAllModules, lRowCount)
                    
                    'Set default value of other parameters
                    nTaskPriority = Empty
                    nTaskTimeAllocated = Empty
                    dtTastExpectedStartTime = Empty
                    bIsCompleted = False
                    
                    
                    'Before showing modal form make this form as normal form (which is otherwise topmost form)
                    MakeFormNonTopMost Me
                    
                    'Don't handle event for LostFocus of txtProject and txtTask
                    mbHandleLostFocus = False
                    
                    
                    Set ofrmProjectTaskEditor = New frmProjectTaskEditor
                    
                    'Now display the form
                    bSuccess = ofrmProjectTaskEditor.DisplayForm(vnMode, nSecurityResult, -1, lProjectID, rsProjectTaskName, avntMembers, avntAllItems, avntMemberModules, avntAllModules, lTaskTypeID, nTaskPriority, nTaskTimeAllocated, dtTastExpectedStartTime, bIsCompleted)
                    
                    Set ofrmProjectTaskEditor = Nothing
                    
                    
                    'Start hadling LostFocus as usual
                    mbHandleLostFocus = True
                    
                    'Make the form top most again
                    MakeFormTopMost Me
                    
                    If bSuccess Then
                    
                        'If user pressed OK then add the entry in database
                        bSuccess = AddNewTask(rsProjectTaskName, avntMembers, avntMemberModules, lTaskTypeID, nTaskPriority, nTaskTimeAllocated, dtTastExpectedStartTime, bIsCompleted)
                        
                    End If
                    
                Else
                
                    MsgBox "You must first add the project before you add the task in it."
                
                End If
                
            Case gnMODE_TASK_EDITOR_MODIFY
            
                avntMembers = Empty
                
                avntAllItems = Empty
                
                sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE Is_Visible = True"
                
                Call GetRowsInArr(sSQL, avntAllItems, lRowCount)
                
                sSQL = "SELECT Users.User_ID, Users.User_Login_Name"
                sSQL = sSQL & " FROM Task_Memberships, Users"
                sSQL = sSQL & " WHERE (Task_Memberships.Task_ID = " & vlProjectTaskID & ")"
                sSQL = sSQL & " AND (Task_Memberships.User_ID = Users.User_ID)"
                sSQL = sSQL & " AND (Is_Visible = True)"
                
                Call GetRowsInArr(sSQL, avntMembers, lRowCount)
                
                
                'Set module memberships
                avntMemberModules = Empty
                
                avntAllModules = Empty
                
                'First get current project ID
                lProjectID = GetSelectedItemDataInCombo(cboProject)
                
                'Get the member modules
                sSQL = "SELECT Sub_Modules.Sub_Module_ID, Modules.Module_Name & '.' & Sub_Modules.Sub_Module_Name"
                sSQL = sSQL & " FROM Sub_Modules, Modules, Task_Sub_Module_Memberships"
                sSQL = sSQL & " WHERE (Modules.Module_ID = Sub_Modules.Module_ID)"
                sSQL = sSQL & " AND (Modules.Task_Category_ID =" & lProjectID & ")"
                sSQL = sSQL & " AND (Task_Sub_Module_Memberships.Task_ID =" & vlProjectTaskID & ")"
                sSQL = sSQL & " AND (Task_Sub_Module_Memberships.Sub_Module_ID = Sub_Modules.Sub_Module_ID)"
                sSQL = sSQL & " AND (Sub_Modules.Is_Visible = True)"
                
                Call GetRowsInArr(sSQL, avntMemberModules, lRowCount)
                
                'Get all submodules associated with project
                sSQL = "SELECT Sub_Modules.Sub_Module_ID, Modules.Module_Name & '.' & Sub_Modules.Sub_Module_Name"
                sSQL = sSQL & " FROM Sub_Modules, Modules"
                sSQL = sSQL & " WHERE (Modules.Module_ID = Sub_Modules.Module_ID)"
                sSQL = sSQL & " AND (Modules.Task_Category_ID =" & lProjectID & ")"
                sSQL = sSQL & " AND (Sub_Modules.Is_Visible = True)"
                
                Call GetRowsInArr(sSQL, avntAllModules, lRowCount)
                
                sSQL = "SELECT Task_Type_ID, Priority, Time_Alloted, Expected_Start,Is_Completed"
                sSQL = sSQL & " FROM Tasks"
                sSQL = sSQL & " WHERE Task_ID = " & vlProjectTaskID
                
                bSuccess = GetRowsInArr(sSQL, avntProjectTaskData, lRowCount)
                
                If bSuccess Then
                
                    If Not IsEmpty(avntProjectTaskData) Then
                    
                        lTaskTypeID = avntProjectTaskData(LBound(avntProjectTaskData, 1), LBound(avntProjectTaskData, 2)) & ""
                        nTaskPriority = avntProjectTaskData(LBound(avntProjectTaskData, 1) + 1, LBound(avntProjectTaskData, 2)) & ""
                        nTaskTimeAllocated = avntProjectTaskData(LBound(avntProjectTaskData, 1) + 2, LBound(avntProjectTaskData, 2)) & ""
                        dtTastExpectedStartTime = avntProjectTaskData(LBound(avntProjectTaskData, 1) + 3, LBound(avntProjectTaskData, 2))
                        bIsCompleted = avntProjectTaskData(LBound(avntProjectTaskData, 1) + 4, LBound(avntProjectTaskData, 2))
                    
                    End If
                    
                End If
                
                'Before showing modal form make this form as normal form (which is otherwise topmost form)
                MakeFormNonTopMost Me
                
                'Don't handle event for LostFocus of txtProject and txtTask
                mbHandleLostFocus = False
                
                
                Set ofrmProjectTaskEditor = New frmProjectTaskEditor
                
                'Now display the form
                bSuccess = ofrmProjectTaskEditor.DisplayForm(vnMode, nSecurityResult, vlProjectTaskID, lProjectID, rsProjectTaskName, avntMembers, avntAllItems, avntMemberModules, avntAllModules, lTaskTypeID, nTaskPriority, nTaskTimeAllocated, dtTastExpectedStartTime, bIsCompleted)
                
                Set ofrmProjectTaskEditor = Nothing
                
                
                'Start hadling LostFocus as usual
                mbHandleLostFocus = True
                
                'Make the form top most again
                MakeFormTopMost Me
                
                If bSuccess Then
                
                    'If user pressed OK then add the entry in database
                    bSuccess = ModifyTask(vlProjectTaskID, rsProjectTaskName, avntMembers, avntMemberModules, lTaskTypeID, nTaskPriority, nTaskTimeAllocated, dtTastExpectedStartTime, bIsCompleted)
                    
                End If
        
        End Select
        
    End If
    
    ShowProjectTaskEditor = bSuccess

End Function



'
'==========================================================================================
' Routine Name : AddNewProject
' Purpose      : Adds a new project in database
' Parameters   : vsProjectName - name of new project to be added
'                vavntMembers - Array containing User_ID and User_Login_Name
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 14-Aug-1998 05:01 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Private Function AddNewProject(ByVal vsProjectName As String, ByVal vavntMembers As Variant) As Boolean

    On Error GoTo ERR_AddNewProject

    'Routine specific local vars here
    Dim lNewProjectID As Long
    
    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "frmWakeUpDlg.AddNewProject"

        
    bSuccess = AddARow("Task_Categories", _
        Array("Task_Category_Name", "Created_By", "Created_On", "Is_Visible", "Last_Modified_By", "Last_Modified_On"), _
        Array(vsProjectName, glUser_ID, Now, True, glUser_ID, Now), "Task_Category_ID", lNewProjectID)
   
    If bSuccess Then
        
        bSuccess = AddRowsWithConst("Task_Category_Memberships", "Task_Category_ID", lNewProjectID, _
            Array("User_ID"), vavntMembers, Array(0), True)
            
    End If
    
    If bSuccess Then
    
        cboProject.AddItem vsProjectName
        
        cboProject.ItemData(cboProject.NewIndex) = lNewProjectID
        
        cboProject.ListIndex = cboProject.NewIndex

    End If


    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    AddNewProject = bSuccess

Exit Function

ERR_AddNewProject:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


'
'==========================================================================================
' Routine Name : ModifyProject
' Purpose      : Modifies the tables related to Task Categories
' Parameters   : vlProjectID: ID of the row in Task_Category table to modify
'                vsProjectName: Modified name of the project
'                vavntMembers: Modified list of project members
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 09-Sep-1998 02:08 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Private Function ModifyProject(ByVal vlProjectID As Long, ByVal vsProjectName As String, ByVal vavntMembers As Variant) As Boolean

    On Error GoTo ERR_ModifyProject

    'Routine specific local vars here
    Dim nSelectedItemIndexInCombo As Integer

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "frmWakeUpDlg.ModifyProject"

        
    bSuccess = ModifyARow("Task_Categories", "Task_Category_ID", vlProjectID, Array("Task_Category_Name", "Last_Modified_By", "Last_Modified_On"), Array(vsProjectName, glUser_ID, Now))
    
    If bSuccess Then
        
        bSuccess = AddRowsWithConst("Task_Category_Memberships", "Task_Category_ID", vlProjectID, _
            Array("User_ID"), vavntMembers, Array(0), True)
                
    End If

    
    If bSuccess Then

        nSelectedItemIndexInCombo = GetSelectedItemInCombo(cboProject)
        
        cboProject.List(nSelectedItemIndexInCombo) = vsProjectName
        
        cboProject.ListIndex = nSelectedItemIndexInCombo
        
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    ModifyProject = bSuccess

Exit Function

ERR_ModifyProject:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


'
'==========================================================================================
' Routine Name : AddNewTask
' Purpose      : Adds a new task in database
' Parameters   : vsTaskName - name of new task to be added
'                vavntMembers - Array containing User_ID and User_Login_Name
'                vavntMemberModules - Array containing Sub_Module_ID, Sub_Module_Name included in the task
'                vlTaskTypeID - ID specifieng the task type
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 14-Aug-1998 05:01 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Private Function AddNewTask(ByVal vsTaskName As String, vavntMembers As Variant, ByVal vavntMemberModules As Variant, ByVal vlTaskTypeID As Long, ByVal vnTaskPriority As Variant, ByVal vnTaskTimeAllocated As Variant, ByVal vdtTaskExpectedStartTime As Variant, ByVal vboolTaskCompleted As Boolean) As Boolean

    On Error GoTo ERR_AddNewTask

    'Routine specific local vars here
    Dim lNewTaskID As Long
    Dim lParentProjectID As Long
    
    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "frmWakeUpDlg.AddNewTask"

    
    lParentProjectID = GetSelectedItemDataInCombo(cboProject)
    
    bSuccess = AddARow("Tasks", _
        Array("Task_Name", "Time_Spent", "Priority", "Time_Alloted", "Expected_Start", "Is_Completed", "Created_By", "Created_On", "Is_Visible", "Task_Category_ID", "Task_Type_ID", "Last_Modified_By", "Last_Modified_On", "Is_Visible"), _
        Array(vsTaskName, 0, ReturnNullIfBlank(vnTaskPriority), ReturnNullIfBlank(vnTaskTimeAllocated), ReturnNullIfBlank(vdtTaskExpectedStartTime, True), vboolTaskCompleted, glUser_ID, Now, True, lParentProjectID, vlTaskTypeID, glUser_ID, Now, True), "Task_ID", lNewTaskID)
   
   
    If bSuccess Then
        
        bSuccess = AddRowsWithConst("Task_Memberships", "Task_ID", lNewTaskID, _
            Array("User_ID"), vavntMembers, Array(0), True)
            
        bSuccess = bSuccess And AddRowsWithConst("Task_Sub_Module_Memberships", "Task_ID", lNewTaskID, _
            Array("Sub_Module_ID"), vavntMemberModules, Array(0), True)
            
    End If
    
    If bSuccess Then
    
        cboTask.AddItem vsTaskName
        
        cboTask.ItemData(cboTask.NewIndex) = lNewTaskID
        
        cboTask.ListIndex = cboTask.NewIndex

    End If


    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    AddNewTask = bSuccess

Exit Function

ERR_AddNewTask:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function



'
'==========================================================================================
' Routine Name : ModifyTask
' Purpose      : Modifies the tables related to Task
' Parameters   : vlTaskID: ID of the row in Task table to modify
'                vsTaskName: Modified name of the task
'                vavntMembers: Modified list of task members
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 09-Sep-1998 02:08 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Private Function ModifyTask(ByVal vlTaskID As Long, ByVal vsTaskName As String, ByVal vavntMembers As Variant, ByVal vavntMemberModules As Variant, ByVal vlTaskTypeID As Long, ByVal vnTaskPriority As Variant, ByVal vnTaskTimeAllocated As Variant, ByVal vdtTaskExpectedStartTime As Variant, ByVal vboolTaskCompleted As Boolean) As Boolean

    On Error GoTo ERR_ModifyTask

    'Routine specific local vars here
    Dim nSelectedItemIndexInCombo As Integer

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "frmWakeUpDlg.ModifyTask"

    
    'Modify the main task entry
    bSuccess = ModifyARow("Tasks", "Task_ID", vlTaskID, Array("Task_Name", "Task_Type_ID", "Priority", "Time_Alloted", "Expected_Start", "Is_Completed", "Last_Modified_By", "Last_Modified_On"), _
                                                        Array(vsTaskName, vlTaskTypeID, ReturnNullIfBlank(vnTaskPriority), ReturnNullIfBlank(vnTaskTimeAllocated), ReturnNullIfBlank(vdtTaskExpectedStartTime, True), vboolTaskCompleted, glUser_ID, Now))
    
    If bSuccess Then
        
        'Modify task-user memberships
        bSuccess = AddRowsWithConst("Task_Memberships", "Task_ID", vlTaskID, _
            Array("User_ID"), vavntMembers, Array(0), True)
            
        'Modify task-sub modules memberships
        bSuccess = bSuccess And AddRowsWithConst("Task_Sub_Module_Memberships", "Task_ID", vlTaskID, _
            Array("Sub_Module_ID"), vavntMemberModules, Array(0), True)
                
    End If

    
    If bSuccess Then

        nSelectedItemIndexInCombo = GetSelectedItemInCombo(cboTask)
        
        cboTask.List(nSelectedItemIndexInCombo) = vsTaskName
        
        cboTask.ListIndex = nSelectedItemIndexInCombo
        
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    ModifyTask = bSuccess

Exit Function

ERR_ModifyTask:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


Private Sub ResetLabelActivations()

    Call SetLabelStyle(lblModifyProject, False)
    Call SetLabelStyle(lblDeleteProject, False)
    Call SetLabelStyle(lblModifyTask, False)
    Call SetLabelStyle(lblDeleteTask, False)
    Call SetLabelStyle(lblAddProject, False)
    Call SetLabelStyle(lblAddTask, False)
    Call SetLabelStyle(lblHints, False)
    Call SetLabelStyle(lblMakeStatusReport, False)
    
End Sub


Private Function SetLabelStyle(ByRef rlbl As Label, ByVal vIsMouseOver As Boolean)

    If vIsMouseOver Then
    
        ResetLabelActivations
    
        rlbl.ForeColor = vbRed
    
    Else
     
        rlbl.ForeColor = vbButtonText
     
    End If

End Function

Private Sub SetDirtyFlag(ByVal vboolFlag As Boolean)

    btnSaveThis.Enabled = vboolFlag
    
    mbFormDirty = vboolFlag

End Sub

Private Sub txtActivity_LostFocus()
    
    If mbUserWasAway Then
        
        MakeUserUnAway
    
    End If
    
End Sub

Private Sub txtInterval_Change()
    Call SetDirtyFlag(True)
End Sub

Private Sub txtTimeSpent_Change()
    
    mbUserEditedTimeSpentField = True

    Call SetDirtyFlag(True)
    
End Sub


Private Sub SetReturnData(ByRef rsActivity As String, ByRef gsDSN As String, ByRef rlInterval As Long, ByRef rlProjectID As Long, ByRef rlTaskID As Long, ByRef rlTimeSpent As Long)

    rsActivity = txtActivity.Text
    gsDSN = txtDSN.Text
    rlInterval = CLng(txtInterval.Text)
    rlProjectID = GetSelectedItemDataInCombo(cboProject)
    rlTaskID = GetSelectedItemDataInCombo(cboTask)
    rlTimeSpent = CLng(txtTimeSpent.Text)

End Sub


Private Sub MakeUserUnAway()
        
        mbUserWasAway = False
        
        'Replace last Activity text
        txtActivity.Text = msLastActivityText
        
        txtActivity.SelStart = 0
        txtActivity.SelLength = Len(txtActivity.Text)

End Sub
