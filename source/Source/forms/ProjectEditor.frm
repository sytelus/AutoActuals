VERSION 5.00
Begin VB.Form frmProjectTaskEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caption at run time"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   Icon            =   "ProjectEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCompleted 
      Caption         =   "Task is completed"
      Height          =   255
      Left            =   6780
      TabIndex        =   44
      Top             =   4440
      Width           =   1635
   End
   Begin VB.TextBox txtStartTime 
      Height          =   285
      Left            =   7860
      TabIndex        =   20
      ToolTipText     =   "Expected time and date for when to start this task"
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox txtTimeAlloted 
      Height          =   285
      Left            =   6420
      TabIndex        =   18
      ToolTipText     =   "How much time is allocated for this task (in min)"
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox txtPriority 
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      ToolTipText     =   "What is the allocated priority for this task"
      Top             =   660
      Width           =   315
   End
   Begin VB.ComboBox cboTaskType 
      Height          =   315
      Left            =   5160
      Sorted          =   -1  'True
      TabIndex        =   14
      Text            =   "cboTask"
      ToolTipText     =   "Type of the task"
      Top             =   120
      Width           =   1755
   End
   Begin VB.Frame fraModuleMembership 
      Caption         =   "Module A&ffected By Task"
      Height          =   3195
      Left            =   4320
      TabIndex        =   29
      Top             =   1200
      Width           =   4035
      Begin VB.ListBox lstMemberModules 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "ProjectEditor.frx":014A
         Left            =   180
         List            =   "ProjectEditor.frx":0151
         MultiSelect     =   2  'Extended
         TabIndex        =   22
         ToolTipText     =   "Modules affected by the task"
         Top             =   540
         Width           =   1575
      End
      Begin VB.ListBox lstAllModules 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "ProjectEditor.frx":0165
         Left            =   2280
         List            =   "ProjectEditor.frx":016C
         MultiSelect     =   2  'Extended
         TabIndex        =   24
         ToolTipText     =   "Modules available in the project"
         Top             =   540
         Width           =   1575
      End
      Begin VB.CommandButton btnMoveToMemberModule 
         Caption         =   "<"
         Height          =   375
         Left            =   1860
         TabIndex        =   25
         Top             =   540
         Width           =   315
      End
      Begin VB.CommandButton btnMoveToAllModules 
         Caption         =   ">"
         Height          =   375
         Left            =   1860
         TabIndex        =   26
         Top             =   1140
         Width           =   315
      End
      Begin VB.CommandButton btnMoveAllToMemberModules 
         Caption         =   "<<"
         Height          =   375
         Left            =   1860
         TabIndex        =   27
         Top             =   1800
         Width           =   315
      End
      Begin VB.CommandButton btnMoveAllToAllModules 
         Caption         =   ">>"
         Height          =   375
         Left            =   1860
         TabIndex        =   28
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label lblAddModule 
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
         Left            =   2760
         TabIndex        =   36
         ToolTipText     =   "Add New User"
         Top             =   2820
         Width           =   165
      End
      Begin VB.Label lblDeleteModule 
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
         Left            =   3330
         TabIndex        =   35
         ToolTipText     =   "Delete User"
         Top             =   2895
         Width           =   135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Affe&cted Modules:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "A&vailable Modules:"
         Height          =   195
         Left            =   2280
         TabIndex        =   23
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblModifyModule 
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
         Height          =   180
         Left            =   3045
         TabIndex        =   34
         ToolTipText     =   "Modify User"
         Top             =   2880
         Width           =   165
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   3000
      TabIndex        =   8
      Top             =   4680
      Width           =   1155
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   960
      TabIndex        =   7
      Top             =   4680
      Width           =   1155
   End
   Begin VB.TextBox txtProjectName 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame fraMembership 
      Caption         =   "Project Member&ship:"
      Height          =   3195
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4035
      Begin VB.CommandButton btnMoveAllToAllItems 
         Caption         =   ">>"
         Height          =   375
         Left            =   1860
         TabIndex        =   12
         Top             =   2400
         Width           =   315
      End
      Begin VB.CommandButton btnMoveAllToMember 
         Caption         =   "<<"
         Height          =   375
         Left            =   1860
         TabIndex        =   11
         Top             =   1800
         Width           =   315
      End
      Begin VB.CommandButton btnMoveToAllItems 
         Caption         =   ">"
         Height          =   375
         Left            =   1860
         TabIndex        =   10
         Top             =   1140
         Width           =   315
      End
      Begin VB.CommandButton btnMoveToMember 
         Caption         =   "<"
         Height          =   375
         Left            =   1860
         TabIndex        =   9
         Top             =   540
         Width           =   315
      End
      Begin VB.ListBox lstAllItems 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "ProjectEditor.frx":017D
         Left            =   2280
         List            =   "ProjectEditor.frx":0184
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         ToolTipText     =   "Available users in the system"
         Top             =   540
         Width           =   1575
      End
      Begin VB.ListBox lstMemberItems 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "ProjectEditor.frx":0195
         Left            =   180
         List            =   "ProjectEditor.frx":019C
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         ToolTipText     =   "Which users can see this item?"
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label lblModifyUser 
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
         Height          =   180
         Left            =   3045
         TabIndex        =   32
         ToolTipText     =   "Modify User"
         Top             =   2880
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Available Users:"
         Height          =   195
         Left            =   2280
         TabIndex        =   5
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Members:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   690
      End
      Begin VB.Label lblDeleteUser 
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
         Left            =   3330
         TabIndex        =   33
         ToolTipText     =   "Delete User"
         Top             =   2895
         Width           =   135
      End
      Begin VB.Label lblAddUser 
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
         Left            =   2760
         TabIndex        =   31
         ToolTipText     =   "Add New User"
         Top             =   2820
         Width           =   165
      End
   End
   Begin VB.Frame fraBreak1 
      Height          =   75
      Left            =   0
      TabIndex        =   30
      Top             =   780
      Width           =   4215
   End
   Begin VB.Label lblSecurityMessage 
      Caption         =   "Security message will be displayed here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   45
      Top             =   660
      Visible         =   0   'False
      Width           =   3690
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "&Start Time:"
      Height          =   195
      Left            =   7020
      TabIndex        =   19
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Time A&lloted:"
      Height          =   195
      Left            =   5460
      TabIndex        =   17
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lblPriority 
      AutoSize        =   -1  'True
      Caption         =   "P&riority:"
      Height          =   195
      Left            =   4320
      TabIndex        =   15
      Top             =   720
      Width           =   510
   End
   Begin VB.Label lblCancelPosFortaskEditor 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancel Btn Positioner for task dlg"
      Height          =   675
      Left            =   4920
      TabIndex        =   43
      Top             =   4260
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblOKPosForTaskEditor 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OK Btn Positioner for task editors"
      Height          =   675
      Left            =   2460
      TabIndex        =   42
      Top             =   4380
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblCancelPosForNonTaskEditor 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancel Btn Positioner for non-task dlg"
      Height          =   555
      Left            =   2640
      TabIndex        =   41
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblOKPosForNonTaskEditor 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OK Btn Positioner for non-task editors"
      Height          =   555
      Left            =   600
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "&Task Type:"
      Height          =   195
      Left            =   4320
      TabIndex        =   13
      Top             =   180
      Width           =   810
   End
   Begin VB.Label lblAddTaskType 
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
      Left            =   6990
      TabIndex        =   39
      ToolTipText     =   "Add New Task"
      Top             =   90
      Width           =   165
   End
   Begin VB.Label lblModifyTaskType 
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
      Left            =   7275
      TabIndex        =   38
      ToolTipText     =   "Modify Task"
      Top             =   60
      Width           =   165
   End
   Begin VB.Label lblDeleteTaskType 
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
      Left            =   7560
      TabIndex        =   37
      ToolTipText     =   "Delete Task"
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblItemName 
      AutoSize        =   -1  'True
      Caption         =   "&Project Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1005
   End
End
Attribute VB_Name = "frmProjectTaskEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbUserPressedCancel As Boolean
Private Prev_Unselected_Text_cboTaskType As String
Private mlItemID As Long
Private mlItemParentID As Long


Function DisplayForm(ByVal vnMode As Integer, ByVal vnSecurityResult As Integer, ByVal vlItemID As Long, ByVal lItemParentID As Long, ByRef rsProjectTaskName As String, ByRef ravntMembers As Variant, ByRef ravntAllItems As Variant, Optional ByRef ravntMemberModules As Variant, Optional ByRef ravntAllModules As Variant, Optional ByRef rlTaskTypeID As Long, Optional ByRef rnTaskPriority As Variant, Optional ByRef rnTaskTimeAllocated As Variant, Optional ByRef rdtTaskExpectedStartTime As Variant, Optional ByRef rbIsCompleted As Boolean) As Boolean

    Dim bSuccess As Boolean
    
    Dim bValidateSuccess As Boolean
    
    bSuccess = True
    
    bValidateSuccess = False
    
    
    Select Case vnMode
    
        Case gnMODE_TASK_EDITOR_ADD, gnMODE_TASK_EDITOR_MODIFY
        
            bSuccess = InitForm(vnMode, vnSecurityResult, vlItemID, lItemParentID, rsProjectTaskName, ravntMembers, ravntAllItems, ravntMemberModules, ravntAllModules, rlTaskTypeID, rnTaskPriority, rnTaskTimeAllocated, rdtTaskExpectedStartTime, rbIsCompleted)
            
        Case Else
        
            bSuccess = InitForm(vnMode, vnSecurityResult, vlItemID, lItemParentID, rsProjectTaskName, ravntMembers, ravntAllItems)
    
            
    End Select
    
    If bSuccess Then
    
        Do
    
            Me.Show vbModal
            
            bSuccess = Not mbUserPressedCancel
            
            If bSuccess Then
            
                bValidateSuccess = ValidateForm(vnMode)
                
                If bValidateSuccess Then
                
                    rsProjectTaskName = txtProjectName.Text
                    
                    Call FillArrayFromList(lstMemberItems, ravntMembers)
                    
                    Select Case vnMode
                    
                        Case gnMODE_TASK_EDITOR_ADD, gnMODE_TASK_EDITOR_MODIFY
                    
                            Call FillArrayFromList(lstMemberModules, ravntMemberModules)
                            
                            If Len(cboTaskType.Text) <> 0 Then
                            
                                rlTaskTypeID = GetSelectedItemDataInCombo(cboTaskType)
                                
                            Else
                            
                                rlTaskTypeID = -1
                            
                            End If
                            
                            rnTaskPriority = txtPriority
                            
                            rnTaskTimeAllocated = txtTimeAlloted
                            
                            rdtTaskExpectedStartTime = txtStartTime
                            
                            rbIsCompleted = CheckToBool(chkCompleted.Value)
                            
                    End Select
                    
                End If
                       
            End If
            
        Loop Until (bValidateSuccess) Or (mbUserPressedCancel)
    
    End If
    
    DisplayForm = bSuccess
    
    Unload Me

End Function

Private Function InitForm(ByVal vnMode As Integer, ByVal vnSecurityResult As Integer, ByVal vlItemID As Long, ByVal vlItemParentID As Long, ByVal vsProjectName As String, ByVal vavntMembers As Variant, ByVal vavntAllItems As Variant, Optional ByVal vavntMemberModules As Variant, Optional ByVal vavntAllModules As Variant, Optional ByVal vlTaskTypeID As Long, Optional ByVal vnTaskPriority As Variant, Optional ByVal vnTaskTimeAllocated As Variant, Optional ByVal vdtTaskExpectedStartTime As Variant, Optional ByVal vboolIsCompleted As Boolean) As Boolean

    Dim bSuccess As Boolean
    
    bSuccess = True

    txtProjectName.Text = vsProjectName
    
    mlItemID = vlItemID
    mlItemParentID = vlItemParentID
    
    
    'Fill user list boxes
    Call FillListFromArray(lstMemberItems, vavntMembers)
    
    Call FillListFromArray(lstAllItems, vavntAllItems)
    
    Call MakeExclusiveLists(lstMemberItems, lstAllItems)
    
    
    Select Case vnMode
            
        Case gnMODE_TASK_EDITOR_ADD, gnMODE_TASK_EDITOR_MODIFY
        
            'Fill Module list boxes
            Call FillListFromArray(lstMemberModules, vavntMemberModules)
    
            Call FillListFromArray(lstAllModules, vavntAllModules)
    
            Call MakeExclusiveLists(lstMemberModules, lstAllModules)
    
            'Fill task type combo
            Call SetupTaskTypeCombo(vlTaskTypeID)
            
            'Set other parameters
            txtPriority.Text = vnTaskPriority & ""
            
            txtTimeAlloted.Text = vnTaskTimeAllocated & ""
            
            txtStartTime.Text = vdtTaskExpectedStartTime & ""
            
            'Display it's value
            chkCompleted.Value = BoolToCheck(vboolIsCompleted)
            
            btnOK.Left = lblOKPosForTaskEditor.Left
            
            btnCancel.Left = lblCancelPosFortaskEditor.Left
            
        Case Else
        
            btnOK.Left = lblOKPosForNonTaskEditor.Left
            
            btnCancel.Left = lblCancelPosForNonTaskEditor.Left
            
            Me.Width = fraModuleMembership.Left - 1

    End Select

        
    Select Case vnMode
    
        Case gnMODE_PROJECT_EDITOR_ADD, gnMODE_PROJECT_EDITOR_MODIFY
        
            fraMembership.Caption = "Project Membership"
            
            lblItemName.Caption = "&Project Name:"
        
        
        Case gnMODE_TASK_EDITOR_ADD, gnMODE_TASK_EDITOR_MODIFY
        
            fraMembership.Caption = "Task Membership"
            
            lblItemName.Caption = "&Task Name:"
        
        
        Case gnMODE_TASK_TYPE_EDITOR_ADD, gnMODE_TASK_TYPE_EDITOR_MODIFY
        
            fraMembership.Caption = "Task Type Membership"
            
            lblItemName.Caption = "&Task Type Name:"
        
    
    End Select
    

    
    Select Case vnMode
    
        Case gnMODE_PROJECT_EDITOR_ADD
        
            Me.Caption = "Add New Project"
            
        Case gnMODE_PROJECT_EDITOR_MODIFY
        
            Me.Caption = "Modify The Project"
            
        Case gnMODE_TASK_EDITOR_ADD
        
            Me.Caption = "Add New Task"
            
        Case gnMODE_TASK_EDITOR_MODIFY
        
            Me.Caption = "Modify The Task"
        
        Case gnMODE_TASK_TYPE_EDITOR_ADD
        
            Me.Caption = "Add New Task Type"
            
        Case gnMODE_TASK_TYPE_EDITOR_MODIFY
        
            Me.Caption = "Modify The Task Type"
        
    End Select
    
    txtProjectName.SelLength = Len(txtProjectName.Text)
    
    
    Select Case vnSecurityResult
    
        Case gnSECURITY_ALLOW
            'No need to anything
        
        Case gnSECURITY_DENIY
            'Programming mistake. Show message and simulate cancel
            MsgBox "The internal security violation in frmProjectTaskEditor has occured. Please contact the support."
            
            ExitApp
            
        
        Case gnSECURITY_PARTIAL_READ_ONLY
            'Make the frames read-only
            Call MakeControlsReadOnly(Me, fraMembership)
            Call EnableActiveLabels(False, True)
           
            fraBreak1.Top = lblSecurityMessage.Top - 85
            
            lblSecurityMessage.Visible = True
            lblSecurityMessage.Caption = "You can not alter the membership because you do not have required rights."

        
        Case gnSECURITY_READ_ONLY
            Call MakeControlsReadOnly(Me)
            Call EnableActiveLabels(False, False)
            
            fraBreak1.Top = lblSecurityMessage.Top - 85
            
            lblSecurityMessage.Visible = True
            lblSecurityMessage.Caption = "You can not alter the filed values because you do not have required rights."
            
            btnCancel.Enabled = True
    
    End Select
    
    InitForm = bSuccess

End Function

Private Sub btnCancel_Click()
    
    mbUserPressedCancel = True
    
    Me.Hide
    
    DoEvents

End Sub



Private Sub btnMoveAllToAllItems_Click()
    Call MoveListToList(False, lstMemberItems, lstAllItems, True)
End Sub

Private Sub btnMoveAllToAllModules_Click()
    Call MoveListToList(False, lstMemberModules, lstAllModules, True)
End Sub

Private Sub btnMoveAllToMember_Click()
    Call MoveListToList(False, lstAllItems, lstMemberItems, True)
End Sub

Private Sub btnMoveAllToMemberModules_Click()
    Call MoveListToList(False, lstAllModules, lstMemberModules, True)
End Sub

Private Sub btnMoveToAllItems_Click()
    Call MoveListToList(True, lstMemberItems, lstAllItems, True)
End Sub

Private Sub btnMoveToAllModules_Click()
    Call MoveListToList(True, lstMemberModules, lstAllModules, True)
End Sub

Private Sub btnMoveToMember_Click()
    Call MoveListToList(True, lstAllItems, lstMemberItems, True)
End Sub

Private Sub btnMoveToMemberModule_Click()
    Call MoveListToList(True, lstAllModules, lstMemberModules, True)
End Sub

Private Sub btnOK_Click()
    
    mbUserPressedCancel = False
    
    Me.Hide
    
    DoEvents

End Sub

Private Sub cboTaskType_Change()
    
    Call IE4LikeCombo_Change(cboTaskType, Prev_Unselected_Text_cboTaskType)
    
End Sub

Private Sub cboTaskType_LostFocus()

    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String

    If (Not IsItemExistInCombo(cboTaskType)) And (Len(cboTaskType.Text) <> 0) Then
    
        nSecurityResult = CheckSecurity(gnMODE_TASK_TYPE_EDITOR_ADD, sSecurityAccessDeniyMsg)
        
        If nSecurityResult = gnSECURITY_DENIY Then
        
            MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to do this task.")
            
            Call NoFailSelectItemInCombo(cboTaskType, 0)
            
        Else
    
            If MsgBox("Do you want to add new task typed named " & AlternateStrIfNull(cboTaskType.Text, "<New Task Type>") & " in your list?", vbYesNo) = vbYes Then
            
                Call ShowTaskTypeEditor(gnMODE_TASK_TYPE_EDITOR_ADD, cboTaskType.Text)
                
            Else
            
                If cboTaskType.ListCount <> 0 Then
                
                    cboTaskType.ListIndex = 0
                    
                End If
                        
            End If
            
        End If
    
    End If

End Sub

Private Sub Form_Deactivate()

    ResetLabelActivations

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Is ALt key pressed
    If (Shift And vbAltMask) > 0 Then
    
        Select Case KeyCode
        
            Case 187        'Plus key: + or =
            
                If Me.ActiveControl Is lstAllItems Then
                
                    lblAddUser_Click
                
                ElseIf Me.ActiveControl Is cboTaskType Then
                
                    lblAddTaskType_Click
                
                ElseIf Me.ActiveControl Is lstAllModules Then
                
                    lblAddModule_Click
                
                End If
            
            Case 192        'Tilde key: ~
            
                If Me.ActiveControl Is lstAllItems Then
                
                    lblModifyUser_Click
                
                ElseIf Me.ActiveControl Is cboTaskType Then
                
                    lblModifyTaskType_Click
                
                ElseIf Me.ActiveControl Is lstAllModules Then
                
                    lblModifyModule_Click
                
                End If
            
            Case vbKeyX
            
                If Me.ActiveControl Is lstAllItems Then
                
                    lblDeleteUser_Click
                
                ElseIf Me.ActiveControl Is cboTaskType Then
                
                    lblDeleteTaskType_Click
                
                ElseIf Me.ActiveControl Is lstAllModules Then
                
                    lblDeleteModule_Click
                
                End If
            
        
        End Select
        
    End If

End Sub

Private Sub Form_LostFocus()
    
    ResetLabelActivations
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ResetLabelActivations
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    mbUserPressedCancel = True
    
End Sub



Private Sub fraMembership_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblAddModule_Click()
    
    Call ShowModuleEditor(gnMODE_MODULE_EDITOR_ADD)

End Sub

Private Sub lblAddModule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblAddModule, True)

End Sub

Private Sub lblAddTaskType_Click()
    
    If IsItemExistInCombo(cboTaskType) Then
    
        Call ShowTaskTypeEditor(gnMODE_TASK_TYPE_EDITOR_ADD, "<New Task Type>")
        
    Else
    
        Call ShowTaskTypeEditor(gnMODE_TASK_TYPE_EDITOR_ADD, cboTaskType.Text)
        
    End If

End Sub

Private Sub lblAddTaskType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblAddTaskType, True)

End Sub

Private Sub lblAddUser_Click()

    Dim sLoginName As String
    Dim lUserID As Long
    
    If AddNewUser(sLoginName, lUserID) Then
    
        Call lstAllItems.AddItem(sLoginName)
            
        lstAllItems.ItemData(lstAllItems.NewIndex) = lUserID

    End If

End Sub

Private Sub lblAddUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblAddUser, True)
    
End Sub

Private Sub lblDeleteModule_Click()
    
    Dim nListIndex As Integer
    
    If lstAllModules.SelCount = 0 Then
    
        MsgBox "First select the sunmodules(s) you want to delete."
        
    Else
    
        If MsgBox("Are you sure you want to delete " & lstAllModules.SelCount & " submodules(s)?", vbYesNoCancel) = vbYes Then
        
            For nListIndex = lstAllModules.ListCount - 1 To 0 Step -1
            
                If lstAllModules.Selected(nListIndex) Then
                
                    If ModifyARow("Sub_Modules", "Sub_Module_ID", lstAllModules.ItemData(nListIndex), Array("Is_Visible", "Last_Modified_By", "Last_Modified_On"), Array(False, glUser_ID, Now)) Then
                    
                        Call lstAllModules.RemoveItem(nListIndex)
                    
                    End If
                
                End If
            
            Next nListIndex
            
        End If
    
    End If

End Sub

Private Sub lblDeleteModule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblDeleteModule, True)

End Sub

Private Sub lblDeleteTaskType_Click()

    Dim bSuccess As Boolean
    Dim nSelectedItemIndexInCombo As Integer
    Dim lTaskTypeID As Long
    Dim vntTaskTypeForUser As Variant
    Dim lRowCount As Long
    Dim sSQL As String
    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String
    
    nSecurityResult = CheckSecurity(gnMODE_TASK_TYPE_EDITOR_DELETE, sSecurityAccessDeniyMsg)
    
    If nSecurityResult = gnSECURITY_DENIY Then
    
        MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to do this task.")
        
    Else
    
        If IsItemExistInCombo(cboTaskType) Then
        
            If MsgBox("Are you sure you want to delete the task type " & cboTaskType.Text & " from your list?", vbYesNoCancel) = vbYes Then
            
                lTaskTypeID = GetSelectedItemDataInCombo(cboTaskType)
    
                'Delete the membership from the project
                bSuccess = DeleteRows("Task_Type_Memberships", Array("Task_Type_ID", "User_ID"), Array(lTaskTypeID, glUser_ID))
                
                'Check if it is really deleted
                sSQL = "SELECT *"
                sSQL = sSQL & " FROM Task_Type_Memberships"
                sSQL = sSQL & " Where (Task_Type_ID = " & lTaskTypeID & ")"
                sSQL = sSQL & " AND ((User_ID = " & glUser_ID & ")"
                sSQL = sSQL & " OR (User_ID = " & glGlobal_User_ID & "))"
                
                Call GetRowsInArr(sSQL, vntTaskTypeForUser, lRowCount)
                
                If IsEmpty(vntTaskTypeForUser) Then
                
                    If bSuccess Then
                
                        nSelectedItemIndexInCombo = GetSelectedItemInCombo(cboTaskType)
                        
                        'Delete from combo too
                        Call cboTaskType.RemoveItem(nSelectedItemIndexInCombo)
                        
                        Call NoFailSelectItemInCombo(cboTaskType, nSelectedItemIndexInCombo + 1, nSelectedItemIndexInCombo - 1)
                        
                    End If
                    
                Else
                
                    MsgBox "The task type can not be removed because it is allocated globally to every user."
    
                End If
            
            End If
            
        Else
        
            Call NoFailSelectItemInCombo(cboTaskType, 0)
            
        End If
        
    End If

End Sub

Private Sub lblDeleteTaskType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblDeleteTaskType, True)

End Sub

Private Sub lblDeleteUser_Click()
    
    Dim nListIndex As Integer
    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String
    Dim lUserID As Long
    
    nSecurityResult = CheckSecurity(gnMODE_USER_EDITOR_DELETE, sSecurityAccessDeniyMsg)
    
    If nSecurityResult = gnSECURITY_DENIY Then
    
        MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to do this task.")
        
    Else
    
        If lstAllItems.SelCount = 0 Then
        
            MsgBox "First select the user(s) you want to delete."
            
        Else
        
            If MsgBox("Are you sure you want to delete " & lstAllItems.SelCount & " user(s)?", vbYesNoCancel) = vbYes Then
            
                For nListIndex = lstAllItems.ListCount - 1 To 0 Step -1
                
                    If lstAllItems.Selected(nListIndex) Then
                    
                        lUserID = lstAllItems.ItemData(nListIndex)
                        
                        If lUserID <> glGlobal_User_ID Then
                        
                            If ModifyARow("Users", "User_ID", lUserID, Array("Is_Visible", "Last_Modified_By", "Last_Modified_On"), Array(False, glUser_ID, Now)) Then
                            
                                Call lstAllItems.RemoveItem(nListIndex)
                            
                            End If
                            
                        Else
                            
                            Call MsgBox("You have selected " & gsGLOBAL_USER_NAME & " for deletion, but it can not be deleted as it is System user.")
                            
                        End If
                    
                    End If
                
                Next nListIndex
                
            End If
        
        End If
        
    End If

End Sub

Private Sub lblDeleteUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblDeleteUser, True)

End Sub

Private Sub lblModifyModule_Click()

    'Is valid selection in list is done?
    If lstAllModules.SelCount = 0 Then
    
        MsgBox "First must first select module/submodule the you want to modify."
        
    ElseIf lstAllModules.SelCount > 1 Then
    
        MsgBox "You must select only one module/submodule that you want to modify."
        
    Else

        Call ShowModuleEditor(gnMODE_MODULE_EDITOR_MODIFY)
        
    End If

End Sub

Private Sub lblModifyModule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblModifyModule, True)

End Sub



Private Sub lblModifyTaskType_Click()

    If IsItemExistInCombo(cboTaskType) Then
    
        Call ShowTaskTypeEditor(gnMODE_TASK_TYPE_EDITOR_MODIFY, cboTaskType.Text, GetSelectedItemDataInCombo(cboTaskType))
        
    Else
    
        Call ShowTaskTypeEditor(gnMODE_PROJECT_EDITOR_ADD, cboTaskType.Text)
        
    End If
    
End Sub

Private Sub lblModifyTaskType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblModifyTaskType, True)

End Sub

Private Sub lblModifyUser_Click()

    Dim sLoginName As String

    If lstAllItems.SelCount = 0 Then
    
        MsgBox "First select the user you want to modify."
        
    ElseIf lstAllItems.SelCount > 1 Then
    
        MsgBox "You must select only one user who you want to modify."
        
    Else
    
        If ModifyUser(lstAllItems.ItemData(lstAllItems.ListIndex), sLoginName) Then
            
            lstAllItems.List(lstAllItems.ListIndex) = sLoginName

        End If
    
    End If

End Sub

Private Sub lblModifyUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SetLabelStyle(lblModifyUser, True)
    
End Sub

Private Sub lstAllItems_DblClick()

    Call btnMoveToMember_Click

End Sub



Private Sub lstAllModules_DblClick()
    Call btnMoveToMemberModule_Click
End Sub

Private Sub lstMemberItems_DblClick()

    btnMoveToAllItems_Click

End Sub

Private Function AddNewUser(ByRef rsNewLoginName As String, ByRef rlNewUserID As Long) As Boolean

    Dim bSuccess As Boolean
    Dim ofrmUserEditor As frmUserEditor
    Dim sFirstName As String
    Dim sLastName As String
    Dim sLoginName As String
    Dim lUserID As Long
    Dim lUserType As Long
    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String

    bSuccess = True
        
    nSecurityResult = CheckSecurity(gnMODE_USER_EDITOR_ADD, sSecurityAccessDeniyMsg)
    
    If nSecurityResult = gnSECURITY_DENIY Then
    
        MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to perform this action")
        
    Else
    
        Set ofrmUserEditor = New frmUserEditor
        
        sFirstName = "<New User>"
        
        bSuccess = ofrmUserEditor.DisplayForm(gnMODE_USER_EDITOR_ADD, nSecurityResult, sFirstName, sLastName, sLoginName, GetUserTypes, lUserType)
        
        If bSuccess Then
            
            'Add in database
            bSuccess = AddARow("Users", Array("First_Name", "Last_Name", "User_Login_Name", "User_Type_ID", "Is_Visible", "Created_By", "Created_On", "Last_Modified_By", "Last_Modified_On"), Array(sFirstName, sLastName, sLoginName, lUserType, True, glUser_ID, Now, glUser_ID, Now), "User_ID", lUserID)
            
            'Update the listboxes
            If bSuccess Then
            
                rlNewUserID = lUserID
                
                rsNewLoginName = sLoginName
            
            End If
        
        End If
        
        Set ofrmUserEditor = Nothing
        
    End If
    
    AddNewUser = bSuccess
    
End Function


Private Function ModifyUser(ByVal vlUserID As Long, ByRef rsNewLoginName As String) As Boolean

    Dim ofrmUserEditor As frmUserEditor
    Dim sFirstName As String
    Dim sLastName As String
    Dim sLoginName As String
    Dim lUserID As Long
    Dim bSuccess As Boolean
    Dim sSQL As String
    Dim lRowCount As Long
    Dim vavntUserData As Variant
    Dim lUserType As Long
    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String

    bSuccess = True
        
    nSecurityResult = CheckSecurity(gnMODE_USER_EDITOR_MODIFY, sSecurityAccessDeniyMsg)
    
    If nSecurityResult = gnSECURITY_DENIY Then
    
        MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to perform this action.")
        
    Else
    
        sSQL = "SELECT First_Name, Last_Name, User_Login_Name, User_Type_ID"
        sSQL = sSQL & " FROM Users"
        sSQL = sSQL & " WHERE User_ID = " & vlUserID
        
        bSuccess = GetRowsInArr(sSQL, vavntUserData, lRowCount)
        
        If bSuccess Then
        
            If Not IsEmpty(vavntUserData) Then
            
                sFirstName = vavntUserData(LBound(vavntUserData, 1), LBound(vavntUserData, 2)) & ""
                sLastName = vavntUserData(LBound(vavntUserData, 1) + 1, LBound(vavntUserData, 2)) & ""
                sLoginName = vavntUserData(LBound(vavntUserData, 1) + 2, LBound(vavntUserData, 2)) & ""
                lUserType = vavntUserData(LBound(vavntUserData, 1) + 3, LBound(vavntUserData, 2))
            
            End If
            
            Set ofrmUserEditor = New frmUserEditor
                
            bSuccess = ofrmUserEditor.DisplayForm(gnMODE_USER_EDITOR_ADD, nSecurityResult, sFirstName, sLastName, sLoginName, GetUserTypes, lUserType)
            
            If bSuccess Then
            
                bSuccess = ModifyARow("Users", "User_ID", vlUserID, Array("First_Name", "Last_Name", "User_Login_Name", "User_Type_ID", "Last_Modified_By", "Last_Modified_On"), Array(sFirstName, sLastName, sLoginName, lUserType, glUser_ID, Now))
                
                If bSuccess Then
                
                    rsNewLoginName = sLoginName
                    
                End If
            
            End If
            
            Set ofrmUserEditor = Nothing
            
        End If
        
    End If
    
    ModifyUser = bSuccess
    
End Function


Private Sub ResetLabelActivations()

    Call SetLabelStyle(lblModifyUser, False)
    Call SetLabelStyle(lblDeleteUser, False)
    Call SetLabelStyle(lblAddUser, False)
    
    Call SetLabelStyle(lblAddTaskType, False)
    Call SetLabelStyle(lblModifyTaskType, False)
    Call SetLabelStyle(lblDeleteTaskType, False)
    
    Call SetLabelStyle(lblAddModule, False)
    Call SetLabelStyle(lblModifyModule, False)
    Call SetLabelStyle(lblDeleteModule, False)
    
End Sub

Private Function GetUserTypes() As Variant

    Dim sSQL As String
    Dim avntUserTypes As Variant
    Dim lRowCount As Long
    
    
    'Get list of all users
    sSQL = "SELECT User_Type_ID, Description FROM User_Types"
    
    Call GetRowsInArr(sSQL, avntUserTypes, lRowCount)
    
    GetUserTypes = avntUserTypes

End Function


Private Sub SetupTaskTypeCombo(ByVal vlTaskTypeIDToBeSelected As Long)

    Dim sSQL As String
    
    sSQL = "SELECT DISTINCT Task_Types.Task_Type_ID, Task_Types.Task_Type_Name"
    sSQL = sSQL & " FROM Task_Types, Task_Type_Memberships"
    sSQL = sSQL & " WHERE (Task_Types.Task_Type_ID = Task_Type_Memberships.Task_Type_ID)"
    sSQL = sSQL & " AND ((Task_Type_Memberships.User_ID = " & glUser_ID & ")"
    sSQL = sSQL & " OR (Task_Type_Memberships.User_ID = " & glGlobal_User_ID & "))"
    
    
    
    Call PopulateCombo(cboTaskType, sSQL)
    
    Call SelectItemInCombo(cboTaskType, vlTaskTypeIDToBeSelected)
    
    Call SetProperComboWidth(cboTaskType)

End Sub

Private Function ShowTaskTypeEditor(ByVal vnMode As Integer, ByRef rsTaskTypeName As String, Optional ByVal vlTaskTypeID As Long) As Boolean

    Dim avntMembers As Variant
    Dim avntAllItems As Variant
    Dim lRowCount As Long
    Dim sSQL As String
    Dim bSuccess As Boolean
    Dim ofrmProjectTaskEditor As frmProjectTaskEditor
    Dim nSecurityResult As Integer
    Dim sSecurityAccessDeniyMsg As String

    bSuccess = True
    
    nSecurityResult = CheckSecurity(vnMode, sSecurityAccessDeniyMsg)
    
    If nSecurityResult = gnSECURITY_DENIY Then
    
        MsgBox AlternateStrIfNull(sSecurityAccessDeniyMsg, "You do not have proper access rights to do this task.")
        
    Else

        Select Case vnMode
        
            Case gnMODE_TASK_TYPE_EDITOR_ADD
            
                        
                avntMembers = Empty
                
                avntAllItems = Empty
                
                'Get list of all users
                sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE Is_Visible = True"
                
                Call GetRowsInArr(sSQL, avntAllItems, lRowCount)
                
                'Set default user
                sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE User_ID = " & glUser_ID
                
                Call GetRowsInArr(sSQL, avntMembers, lRowCount)
                
                Set ofrmProjectTaskEditor = New frmProjectTaskEditor
                
                'Now display the form
                bSuccess = ofrmProjectTaskEditor.DisplayForm(vnMode, nSecurityResult, -1, -1, rsTaskTypeName, avntMembers, avntAllItems)
                
                Set ofrmProjectTaskEditor = Nothing
               
                If bSuccess Then
                
                    'If user pressed OK then add the entry in database
                    bSuccess = AddNewTaskType(rsTaskTypeName, avntMembers)
                    
                End If
                
            Case gnMODE_TASK_TYPE_EDITOR_MODIFY
            
                avntMembers = Empty
                
                avntAllItems = Empty
                
                sSQL = "SELECT User_ID, User_Login_Name FROM Users WHERE Is_Visible = True"
                
                Call GetRowsInArr(sSQL, avntAllItems, lRowCount)
                
                sSQL = "SELECT Users.User_ID, Users.User_Login_Name"
                sSQL = sSQL & " FROM Task_Type_Memberships, Users"
                sSQL = sSQL & " WHERE (Task_Type_Memberships.Task_Type_ID = " & vlTaskTypeID & ")"
                sSQL = sSQL & " AND (Task_Type_Memberships.User_ID = Users.User_ID)"
                sSQL = sSQL & " AND (Is_Visible = True)"
                
                
                Call GetRowsInArr(sSQL, avntMembers, lRowCount)
                
                Set ofrmProjectTaskEditor = New frmProjectTaskEditor
                
                'Now display the form
                bSuccess = ofrmProjectTaskEditor.DisplayForm(vnMode, nSecurityResult, vlTaskTypeID, -1, rsTaskTypeName, avntMembers, avntAllItems)
                
                Set ofrmProjectTaskEditor = Nothing
                
                If bSuccess Then
                
                    'If user pressed OK then add the entry in database
                    bSuccess = ModifyTaskType(vlTaskTypeID, rsTaskTypeName, avntMembers)
                    
                End If
        
        End Select
        
    End If
    
    ShowTaskTypeEditor = bSuccess

End Function


Private Function AddNewTaskType(ByVal vsTaskTypeName As String, vavntMembers As Variant) As Boolean

    On Error GoTo ERR_AddNewTaskType

    'Routine specific local vars here
    Dim lNewTaskTypeID As Long
    
    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "frmProjectTaskEditor.AddNewTaskType"

        
    bSuccess = AddARow("Task_Types", _
        Array("Task_Type_Name", "Created_By", "Created_On", "Is_Visible", "Last_Modified_By", "Last_Modified_On"), _
        Array(vsTaskTypeName, glUser_ID, Now, True, glUser_ID, Now), "Task_Type_ID", lNewTaskTypeID)
   
    If bSuccess Then
        
        bSuccess = AddRowsWithConst("Task_Type_Memberships", "Task_Type_ID", lNewTaskTypeID, _
            Array("User_ID"), vavntMembers, Array(0), True)
            
    End If
    
    If bSuccess Then
    
        cboTaskType.AddItem vsTaskTypeName
        
        cboTaskType.ItemData(cboTaskType.NewIndex) = lNewTaskTypeID
        
        cboTaskType.ListIndex = cboTaskType.NewIndex

    End If


    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    AddNewTaskType = bSuccess

Exit Function

ERR_AddNewTaskType:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


Private Function ModifyTaskType(ByVal vlTaskTypeID As Long, ByVal vsTaskTypeName As String, ByVal vavntMembers As Variant) As Boolean

    On Error GoTo ERR_ModifyTaskType

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
    sErrorLocation = "frmProjectTaskEditor.ModifyTaskType"

        
    bSuccess = ModifyARow("Task_Types", "Task_Type_ID", vlTaskTypeID, Array("Task_Type_Name", "Last_Modified_By", "Last_Modified_On"), Array(vsTaskTypeName, glUser_ID, Now))
    
    If bSuccess Then
        
        bSuccess = AddRowsWithConst("Task_Type_Memberships", "Task_Type_ID", vlTaskTypeID, _
            Array("User_ID"), vavntMembers, Array(0), True)
                
    End If

    
    If bSuccess Then

        nSelectedItemIndexInCombo = GetSelectedItemInCombo(cboTaskType)
        
        cboTaskType.List(nSelectedItemIndexInCombo) = vsTaskTypeName
        
        cboTaskType.ListIndex = nSelectedItemIndexInCombo
        
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    ModifyTaskType = bSuccess

Exit Function

ERR_ModifyTaskType:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


Private Function ShowModuleEditor(ByVal vnMode As Integer) As Boolean

    Dim ofrmTextListEditor As frmTextListEditor
    Dim sModuleName As String
    Dim avntSubModules As Variant
    Dim lSubModuleID As Long
    Dim lModuleID As Long
    Dim sSQL As String
    Dim lRowCount As Long
    Dim bSuccess As Boolean
    Dim avntSavedMemberModules As Variant
    Dim avntAllModules As Variant
    Dim avntActualMemberModules As Variant
    
    bSuccess = True
    
    Select Case vnMode
    
        Case gnMODE_MODULE_EDITOR_ADD
        
            sModuleName = "<New Module>"
            
            avntSubModules = Empty
            
            'Call Text&List editor
            Set ofrmTextListEditor = New frmTextListEditor
            
                bSuccess = ofrmTextListEditor.DisplayForm("Add New Module and Submodules", "Module Name:", "Submodules:", "Module", "Submodule", sModuleName, avntSubModules)
            
            Set ofrmTextListEditor = Nothing
            
            If bSuccess Then
            
                'Add it to database
                bSuccess = AddNewModule(sModuleName, avntSubModules)
                
                'Refresh list box
                Call LoadAllModulesInList
                
                'Make both the lists exclusive
                Call MakeExclusiveLists(lstMemberModules, lstAllModules)
            
            End If
            
        
        Case gnMODE_MODULE_EDITOR_MODIFY
        
            'Get name of parent module
            lSubModuleID = lstAllModules.ItemData(lstAllModules.ListIndex)
            
            sSQL = "SELECT Module_ID FROM Sub_Modules WHERE Sub_Module_ID = " & lSubModuleID
            
            Call GetColValue(sSQL, lModuleID)
            
            sSQL = "SELECT Module_Name FROM Modules WHERE Module_ID = " & lModuleID
            
            Call GetColValue(sSQL, sModuleName)
            
            'Now get the all submodules associated with module
            sSQL = "SELECT Sub_Module_ID, Sub_Module_Name"
            sSQL = sSQL & " FROM Sub_Modules"
            sSQL = sSQL & " WHERE (Module_ID = " & lModuleID & ")"
            sSQL = sSQL & " AND (Is_visible = True)"
            
            'Default value
            avntSubModules = Empty
                        
            Call GetRowsInArr(sSQL, avntSubModules, lRowCount)
            
            'Save the current status of members
            Call FillArrayFromList(lstMemberModules, avntSavedMemberModules)
            
            'Call Text&List editor
            Set ofrmTextListEditor = New frmTextListEditor
            
                bSuccess = ofrmTextListEditor.DisplayForm("Modify Module and Submodules", "Module Name:", "Submodules:", "Module", "Submodule", sModuleName, avntSubModules)
            
            Set ofrmTextListEditor = Nothing
            
            If bSuccess Then
            
                'Add it to database
                bSuccess = ModifyModule(lModuleID, sModuleName, avntSubModules)
                
                'Refresh list box
                LoadAllModulesInList
                
                'Get all submodules in arr
                Call FillArrayFromList(lstAllModules, avntAllModules)
                
                'Get new member modules
                Call MakeArrayOfIncludedIDs(avntSavedMemberModules, avntAllModules, avntActualMemberModules)
                
                'Refill the list
                Call FillListFromArray(lstMemberModules, avntActualMemberModules)
                
                'Finally make both the lists exclusive
                Call MakeExclusiveLists(lstMemberModules, lstAllModules)
            
            End If
    
    
    End Select
    
    
    ShowModuleEditor = bSuccess

End Function



Private Function AddNewModule(ByVal vsModuleName As String, vavntSubModules As Variant) As Boolean

    On Error GoTo ERR_AddNewModule

    'Routine specific local vars here
    Dim lNewModuleID As Long
    Dim lKeyValue As Long
    
    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "frmProjectTaskEditor.AddNewModule"

        
    bSuccess = AddARow("Modules", _
        Array("Module_Name", "Task_Category_ID", "Created_By", "Created_On", "Is_Visible", "Last_Modified_By", "Last_Modified_On"), _
        Array(vsModuleName, mlItemParentID, glUser_ID, Now, True, glUser_ID, Now), "Module_ID", lNewModuleID)
   
    If bSuccess Then
    
        'Is sub modules array is empty?
        If IsArrayEmpty(vavntSubModules) Then
        
            'Add a row in this array with default submodule
            vavntSubModules = Empty
            
            'Array for two columns and one row
            ReDim vavntSubModules(0 To 1, 0 To 0)
            
            'Set the name of default submodule
            vavntSubModules(1, 0) = gsDEFAULT_SUB_MODULE_NAME
        
        End If
        
        bSuccess = AddRowsWithConst("Sub_Modules", "Module_ID", lNewModuleID, _
            Array("Sub_Module_Name"), vavntSubModules, Array(1), False, "Sub_Module_ID", lKeyValue)
            
    End If
    

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    AddNewModule = bSuccess

Exit Function

ERR_AddNewModule:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function

Private Function ModifyModule(ByVal vlModuleID As Long, ByVal vsModuleName As String, ByVal vavntSubModules As Variant) As Boolean

    On Error GoTo ERR_ModifyModule
    
    Dim lKeyValue As Long

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "frmProjectTaskEditor.ModifyModule"

        
    'Modify main row for Module
    bSuccess = ModifyARow("Modules", "Module_ID", vlModuleID, Array("Module_Name", "Last_Modified_By", "Last_Modified_On"), Array(vsModuleName, glUser_ID, Now))
    
    If bSuccess Then
    
        'Is sub modules array is empty?
        If IsArrayEmpty(vavntSubModules) Then
        
            'Add a row in this array with default submodule
            vavntSubModules = Empty
            
            'Array for two columns and one row
            ReDim vavntSubModules(0 To 1, 0 To 0)
            
            'Set the name of default submodule
            vavntSubModules(1, 0) = gsDEFAULT_SUB_MODULE_NAME
            vavntSubModules(0, 0) = -1      'Invalid primary key
        
        End If
        
        bSuccess = AddRowsWithConst("Sub_Modules", "Module_ID", vlModuleID, _
            Array("Sub_Module_Name"), vavntSubModules, Array(1), True, "Sub_Module_ID", lKeyValue, 0)
            
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    ModifyModule = bSuccess

Exit Function

ERR_ModifyModule:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function

Function LoadAllModulesInList() As Boolean
    
    Dim avntMemberModules As Variant
    Dim avntAllModules As Variant
    Dim sSQL As String
    Dim lRowCount As Long
    
    'Set module memberships
    avntMemberModules = Empty
    
    avntAllModules = Empty
    
    If mlItemParentID <> -1 Then
    
        'Get all submodules associated with project
        sSQL = "SELECT Sub_Modules.Sub_Module_ID, Modules.Module_Name & '.' & Sub_Modules.Sub_Module_Name"
        sSQL = sSQL & " FROM Sub_Modules, Modules"
        sSQL = sSQL & " WHERE (Modules.Module_ID = Sub_Modules.Module_ID)"
        sSQL = sSQL & " AND (Modules.Task_Category_ID =" & mlItemParentID & ")"
        sSQL = sSQL & " AND (Sub_Modules.Is_Visible = True)"
        
        Call GetRowsInArr(sSQL, avntAllModules, lRowCount)
        
        'Fill the list box from this array
        Call FillListFromArray(lstAllModules, avntAllModules)
        
        
        If mlItemID <> -1 Then
            
            'Get the member modules
        sSQL = "SELECT Sub_Modules.Sub_Module_ID, Modules.Module_Name & '.' & Sub_Modules.Sub_Module_Name"
            sSQL = sSQL & " FROM Sub_Modules, Modules, Task_Sub_Module_Memberships"
            sSQL = sSQL & " WHERE (Modules.Module_ID = Sub_Modules.Module_ID)"
            sSQL = sSQL & " AND (Modules.Task_Category_ID =" & mlItemParentID & ")"
            sSQL = sSQL & " AND (Task_Sub_Module_Memberships.Task_ID =" & mlItemID & ")"
            sSQL = sSQL & " AND (Task_Sub_Module_Memberships.Sub_Module_ID = Sub_Modules.Sub_Module_ID)"
            sSQL = sSQL & " AND (Sub_Modules.Is_Visible = True)"
            
            Call GetRowsInArr(sSQL, avntMemberModules, lRowCount)
            
            'Fill the list box from this array
            Call FillListFromArray(lstMemberModules, avntMemberModules)
            
        Else
        
            'MsgBox "Can not retrive the list of sub-modules for the task because you haven't selected valid task."
            
        End If
        
        
    Else
    
        'MsgBox "Can not retrive the list of member sub-modules for the project because you haven't selected valid project."
        
    End If
    
End Function

Private Sub lstMemberModules_DblClick()
    btnMoveToAllModules_Click
End Sub


Private Sub MakeArrayOfIncludedIDs(ByVal vavntArrToCheckForItem As Variant, ByVal vavntArrWithItems As Variant, ByRef ravntResultArr As Variant)

    Dim nArrToCheckForItemIndex As Integer
    Dim lItemID As Long
    Dim nFirstColNum As Integer
    Dim bItemFound As Boolean
    Dim nItemIndexFound As Integer
    Dim nResultRowNum As Integer
    
    ravntResultArr = Empty
    
    nResultRowNum = 0
    
    If Not IsArrayEmpty(vavntArrToCheckForItem) Then
    
        nFirstColNum = LBound(vavntArrToCheckForItem, 1)
        
        For nArrToCheckForItemIndex = LBound(vavntArrToCheckForItem, 2) To UBound(vavntArrToCheckForItem, 2)
        
            lItemID = vavntArrToCheckForItem(nFirstColNum, nArrToCheckForItemIndex)
            
            If ItemExistInArray(vavntArrWithItems, nFirstColNum, lItemID, nItemIndexFound) Then
            
                If IsEmpty(ravntResultArr) Then
               
                    ReDim ravntResultArr(nFirstColNum To nFirstColNum + 1, 0 To nResultRowNum)
            
                Else
                    
                    ReDim Preserve ravntResultArr(nFirstColNum To nFirstColNum + 1, 0 To nResultRowNum)
                    
                End If
                
                
                ravntResultArr(nFirstColNum, nResultRowNum) = lItemID
                
                ravntResultArr(nFirstColNum + 1, nResultRowNum) = vavntArrWithItems(nFirstColNum + 1, nItemIndexFound)
                
                nResultRowNum = nResultRowNum + 1
            
            End If
        
        Next nArrToCheckForItemIndex
    
    End If

End Sub


Private Function SetLabelStyle(ByRef rlbl As Label, ByVal vIsMouseOver As Boolean)

    If vIsMouseOver Then
    
        ResetLabelActivations
    
        rlbl.ForeColor = vbRed
    
    Else
     
        rlbl.ForeColor = vbButtonText
     
    End If

End Function

Private Function ValidateForm(ByVal vnMode As Integer) As Boolean


    Dim bSuccess As Boolean
    Dim sMsg As String
    
    bSuccess = True
    
    sMsg = "Following errors occured: " & vbCrLf
    
    If Len(txtProjectName.Text) = 0 Then
        
        bSuccess = False
        
        sMsg = sMsg & vbCrLf & GetItemName(vnMode) & " name " & " can't be blank"
        
    End If
    
    Select Case vnMode
    
        Case gnMODE_TASK_EDITOR_ADD, gnMODE_TASK_EDITOR_MODIFY
        
                If Len(txtPriority.Text) <> 0 Then
                    
                    If Not IsNumber(txtPriority.Text) Then
                        
                        bSuccess = False
                    
                        sMsg = sMsg & vbCrLf & "You must enter valid number in Priority field"
                        
                    End If
        
                End If
    
                If Len(txtTimeAlloted.Text) <> 0 Then
                    
                    If Not IsNumber(txtTimeAlloted.Text) Then
                        
                        bSuccess = False
                    
                        sMsg = sMsg & vbCrLf & "You must enter valid number of minutes in Time Alloted field"
                        
                    End If
        
                End If
    
                If Len(txtStartTime.Text) <> 0 Then
                    
                    If Not IsDate(txtStartTime.Text) Then
                        
                        bSuccess = False
                    
                        sMsg = sMsg & vbCrLf & "You must enter valid date/time in Start Time field"
                        
                    End If
        
                End If
    
    End Select
    
    
    If Not bSuccess Then
    
        MsgBox sMsg
    
    End If
    
    ValidateForm = bSuccess


End Function


Private Function GetItemName(ByVal vnMode As Integer) As String

    Select Case vnMode
    
        Case gnMODE_PROJECT_EDITOR_ADD, gnMODE_PROJECT_EDITOR_MODIFY
        
            GetItemName = "Project"
        
        Case gnMODE_TASK_TYPE_EDITOR_ADD, gnMODE_TASK_EDITOR_MODIFY
        
            GetItemName = "Task"
        
        Case gnMODE_TASK_TYPE_EDITOR_ADD, gnMODE_TASK_TYPE_EDITOR_MODIFY
        
            GetItemName = "Task Type"
    
    End Select

End Function

Private Sub EnableActiveLabels(ByVal vboolEnable As Boolean, ByVal vboolAffectOnlyUserLabels As Boolean)
    
    If Not vboolAffectOnlyUserLabels Then
    
        lblAddModule.Enabled = vboolEnable
        lblModifyModule.Enabled = vboolEnable
        lblDeleteModule.Enabled = vboolEnable
        
        lblAddTaskType.Enabled = vboolEnable
        lblModifyTaskType.Enabled = vboolEnable
        lblDeleteTaskType.Enabled = vboolEnable
        
    End If
    
    lblAddUser.Enabled = vboolEnable
    lblModifyUser.Enabled = vboolEnable
    lblDeleteUser.Enabled = vboolEnable

End Sub
