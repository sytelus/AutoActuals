VERSION 5.00
Begin VB.Form frmUserEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caption At Run Time"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   Icon            =   "UserEditor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   3315
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtLoginName 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "Login name used by user to log on to the WIndows"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cboUserType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "This determines access right for the user"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lblLastName 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lblLoginName 
         AutoSize        =   -1  'True
         Caption         =   "Login Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   780
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   360
      TabIndex        =   8
      Top             =   2460
      Width           =   1155
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2100
      TabIndex        =   9
      Top             =   2460
      Width           =   1155
   End
End
Attribute VB_Name = "frmUserEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbUserPressedCancel As Boolean


Function DisplayForm(ByVal vnMode As Integer, ByVal vnSecurityResult As Integer, ByRef rsFirstName As String, ByRef rsLastName As String, ByRef rsLoginName As String, ByVal vavntUserTypes As Variant, ByRef rlUserType As Long) As Boolean

    Dim bSuccess As Boolean
    Dim bValidateSuccess As Boolean
    
    bSuccess = True
    
    bValidateSuccess = False
    

    bSuccess = InitForm(vnMode, vnSecurityResult, rsFirstName, rsLastName, rsLoginName, vavntUserTypes, rlUserType)
    
    
    If bSuccess Then
    
        Do
    
            Me.Show vbModal
            
            bSuccess = Not mbUserPressedCancel
            
            If bSuccess Then
            
                bValidateSuccess = ValidateForm(vnMode, rsLoginName)
                
                If bValidateSuccess Then
                
                    bSuccess = SetReturnValues(rsFirstName, rsLastName, rsLoginName, rlUserType)
                    
                End If
                       
            End If
            
        Loop Until (bValidateSuccess) Or (mbUserPressedCancel)
    
    End If
    
    DisplayForm = bSuccess
    
    Unload Me

End Function


Private Sub btnOK_Click()
    
    mbUserPressedCancel = False
    
    Me.Hide
    
    DoEvents

End Sub


Private Sub btnCancel_Click()
    
    mbUserPressedCancel = True
    
    Me.Hide
    
    DoEvents

End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'This is in case user presses the Close button in system menu
    mbUserPressedCancel = True
    
End Sub

Private Function InitForm(ByVal vnMode As Integer, ByVal vnSecurityResult As Integer, ByVal vsFirstName As String, ByVal vsLastName As String, ByVal vsLoginName As String, ByVal vavntUserTypes As Variant, ByVal vlUserType As Long) As Boolean

    Dim bSuccess As Boolean
    
    bSuccess = True
    

    txtFirstName = vsFirstName
    
    txtLastName = vsLastName
    
    txtLoginName = vsLoginName
    
    txtFirstName.SelLength = Len(txtFirstName)
    
    Call FillListFromArray(cboUserType, vavntUserTypes)
    
    Call SelectItemInCombo(cboUserType, vlUserType)

    Select Case vnMode
    
        Case gnMODE_USER_EDITOR_ADD
        
            Me.Caption = "Add New User"
        
        Case gnMODE_USER_EDITOR_MODIFY
        
            Me.Caption = "Modify User"
    
    End Select
    
    txtFirstName.SelLength = Len(txtFirstName)
    
    
    'Setup the security
    Select Case vnSecurityResult
    
        Case gnSECURITY_ALLOW
            'No need to anything
        
        Case gnSECURITY_DENIY
            'Programming mistake. Show message and simulate cancel
            MsgBox "The internal security violation in frmUserEditor has occured. Please contact the support."
            
            ExitApp
            
        
        Case gnSECURITY_PARTIAL_READ_ONLY
            'Not applicable
                    
        Case gnSECURITY_READ_ONLY
            Call MakeControlsReadOnly(Me)
            
            btnCancel.Enabled = True
    
    End Select

    
    InitForm = bSuccess

End Function


Function SetReturnValues(ByRef rsFirstName As String, ByRef rsLastName As String, ByRef rsLoginName As String, ByRef rlUserType As Long)

    Dim bSuccess As Boolean
    
    bSuccess = True
    
    rsFirstName = txtFirstName
    
    rsLastName = txtLastName
    
    rsLoginName = txtLoginName
    
    rlUserType = cboUserType.ItemData(cboUserType.ListIndex)
    
    
    SetReturnValues = bSuccess

End Function


Private Function ValidateForm(ByVal vnMode As Integer, ByVal vsOriginalLoginName As String) As Boolean

    Dim bSuccess As Boolean
    Dim sMsg As String
    
    bSuccess = True
    
    sMsg = "Following errors occured: "
    
    If Len(txtLoginName.Text) = 0 Then
        
        bSuccess = False
        
        sMsg = sMsg & vbCrLf & vbCrLf & "Login name can not be the blank. Please enter valid login name for the user."
        
    Else
    
        If Not CheckDuplicateLoginName(vnMode, vsOriginalLoginName, sMsg) Then
            
            bSuccess = False
            
        End If
        
    End If
    
    If Not bSuccess Then
    
        MsgBox sMsg
    
    End If
    
    ValidateForm = bSuccess

End Function


Function CheckDuplicateLoginName(ByVal vnMode As Integer, ByVal vsOriginalLoginName As String, ByRef rsErrMessages As String) As Boolean

    Dim bSuccess As Boolean
    
    bSuccess = True

    Select Case vnMode
    
        Case gnMODE_USER_EDITOR_ADD
        
            'In add mode visible login name should not be there
            If IsVisibleUserExists(txtLoginName) Then
            
                bSuccess = False
                
                rsErrMessages = rsErrMessages & vbCrLf & "The Login name you typed is already in use by other user. As two user can not have same login name, you must assign different login name for this user."
           
            End If
        
        Case gnMODE_USER_EDITOR_MODIFY
        
            'In modify mode new login name typed should not be already exist
            If vsOriginalLoginName <> txtLoginName Then
            
                'In add mode visible login name should not be there
                If IsVisibleUserExists(txtLoginName) Then
                 
                    bSuccess = False
                    
                    rsErrMessages = rsErrMessages & vbCrLf & "The Login name you typed is already in use by other user. As two user can not have same login name, you must assign different login name for this user."
                
                End If
                
            End If
    
    End Select

    CheckDuplicateLoginName = bSuccess

End Function
