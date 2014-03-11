VERSION 5.00
Begin VB.Form frmTextListEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caption At Run Time"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   180
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3255
      Begin VB.ListBox lstList 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "ModuleEditor.frx":0000
         Left            =   180
         List            =   "ModuleEditor.frx":0007
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   420
         Width           =   2895
      End
      Begin VB.Label lblListModify 
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
         Left            =   2685
         TabIndex        =   9
         ToolTipText     =   "Modify User"
         Top             =   2760
         Width           =   165
      End
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         Caption         =   "Sub Modules:"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lblListDelete 
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
         Left            =   2940
         TabIndex        =   8
         ToolTipText     =   "Delete User"
         Top             =   2775
         Width           =   135
      End
      Begin VB.Label lblListAdd 
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
         Left            =   2400
         TabIndex        =   7
         ToolTipText     =   "Add New User"
         Top             =   2700
         Width           =   165
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   300
      TabIndex        =   4
      Top             =   3840
      Width           =   1155
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Task Type:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmTextListEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbUserPressedCancel As Boolean
Private msTextFieldName As String
Private msListFieldName As String


Function DisplayForm(ByVal vsFormCaption As String, ByVal vsLabelTextCaption As String, ByVal vsLabelListCaption As String, ByVal vsTextFieldName As String, ByVal vsListFieldName As String, ByRef rsText As String, ByRef ravntList As Variant) As Boolean

    Dim bSuccess As Boolean
    Dim bValidateSuccess As Boolean
    
    
    bSuccess = True
    
    bValidateSuccess = False
    

    bSuccess = InitForm(vsFormCaption, vsLabelTextCaption, vsLabelListCaption, vsTextFieldName, vsListFieldName, rsText, ravntList)
    
    If bSuccess Then
    
        Do
    
            Me.Show vbModal
            
            bSuccess = Not mbUserPressedCancel
            
            If bSuccess Then
            
                bValidateSuccess = ValidateForm(vsTextFieldName, vsListFieldName)
                
                If bValidateSuccess Then
                
                    bSuccess = SetReturnValues(rsText, ravntList)
                    
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




Private Sub Form_Deactivate()

    ResetLabelActivations

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Is ALt key pressed
    If (Shift And vbAltMask) > 0 Then
    
        Select Case KeyCode
        
            Case 187        'Plus key: + or =
            
                If Me.ActiveControl Is lstList Then
                
                    lblListAdd_Click
                
                End If
            
            Case 192        'Tilde key: ~
            
                If Me.ActiveControl Is lstList Then
                
                    lblListModify_Click
                
                End If
            
            Case vbKeyX
            
                If Me.ActiveControl Is lstList Then
                
                    lblListDelete_Click
                
                End If
            
        
        End Select
        
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetLabelActivations

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'This is in case user presses the Close button in system menu
    mbUserPressedCancel = True
    
End Sub

Private Function InitForm(ByVal vsFormCaption As String, ByVal vsLabelTextCaption As String, ByVal vsLabelListCaption As String, ByVal vsTextFieldName As String, ByVal vsListFieldName As String, ByVal vsText As String, ByVal vavntList As Variant) As Boolean

    Dim bSuccess As Boolean
    
    bSuccess = True
    
    
    msTextFieldName = vsTextFieldName
    msListFieldName = vsListFieldName
    
    Me.Caption = vsFormCaption
    
    lblName.Caption = vsLabelTextCaption
    
    lblList.Caption = vsLabelListCaption
    
    
    txtName = vsText
    
    Call FillListFromArray(lstList, vavntList)

    txtName.SelLength = Len(txtName)
    
    InitForm = bSuccess

End Function


Function SetReturnValues(ByRef rsName As String, ByRef ravntList As Variant)

    Dim bSuccess As Boolean
    
    bSuccess = True
    
        
    rsName = txtName
    
    bSuccess = FillArrayFromList(lstList, ravntList)
    
    SetReturnValues = bSuccess

End Function


Private Function ValidateForm(ByVal vsTextFieldName As String, ByVal vsListFieldName As String) As Boolean

    Dim bSuccess As Boolean
    Dim sMsg As String
    
    bSuccess = True
    
    sMsg = "Following errors occured: "
    
    If Len(txtName.Text) = 0 Then
        
        bSuccess = False
        
        sMsg = sMsg & vbCrLf & vbCrLf & msTextFieldName & " can't be blank."
        
    End If
    
    If Not bSuccess Then
    
        MsgBox sMsg
    
    End If
    
    ValidateForm = bSuccess

End Function



Private Sub lblListAdd_Click()

    Dim sItemName As String
    Dim sItemDesc As String
    Dim sUserResponse As String
    
    sUserResponse = InputBox("Name of new " & msListFieldName & ":", "Add new " & msListFieldName, "<New " & msListFieldName & ">")
    
    'If user didn't pressed the Cancel
    If Len(sUserResponse) <> 0 Then
    
        'Is item already there in the list?
        If GetItemIndexInList(lstList, sUserResponse) <> -1 Then
        
            MsgBox "The " & msListFieldName & " " & sUserResponse & " already exist in the list."
            
        Else
            
            'Add it in the list
            lstList.AddItem sUserResponse
            lstList.ItemData(lstList.NewIndex) = -1
        
        End If
    
    End If

End Sub

Private Sub lblListAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetLabelStyle(lblListAdd, True)

End Sub

Private Sub lblListDelete_Click()

    Dim nListIndex As Integer
    
    If lstList.SelCount > 0 Then
        
        If MsgBox("Do you really want to delete " & lstList.SelCount & " item(s)?", vbYesNoCancel) = vbYes Then
        
            For nListIndex = lstList.ListCount - 1 To 0 Step -1
            
                If lstList.Selected(nListIndex) Then
                
                    Call lstList.RemoveItem(nListIndex)
                
                End If
            
            Next nListIndex
        
        End If
        
    Else
    
        Call MsgBox("Please first select the items you want to delete.")
    
    End If

End Sub

Private Sub lblListDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetLabelStyle(lblListDelete, True)

End Sub

Private Sub lblListModify_Click()

    Dim sItemName As String
    Dim sItemDesc As String
    Dim sUserResponse As String
    
    'Is valid selection in list is done?
    If lstList.SelCount = 0 Then
    
        MsgBox "First select the " & msListFieldName & " you want to modify."
        
    ElseIf lstList.SelCount > 1 Then
    
        MsgBox "You must select only one " & msListFieldName & "  that you want to modify."
        
    Else
    
        sUserResponse = InputBox("Modify the " & msListFieldName & ":", "Modify " & msListFieldName, lstList.Text)
        
        'If user didn't pressed the Cancel
        If Len(sUserResponse) <> 0 Then
        
            'Is item already there in the list?
            If (GetItemIndexInList(lstList, sUserResponse) <> -1) And (GetItemIndexInList(lstList, sUserResponse) <> lstList.ListIndex) Then
            
                MsgBox "The " & msListFieldName & " " & sUserResponse & " already exist in the list."
                
            Else
                
                'Add it in the list
                lstList.List(lstList.ListIndex) = sUserResponse
            
            End If
        
        End If
    
    End If

End Sub

Private Sub lblListModify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call SetLabelStyle(lblListModify, True)

End Sub

Private Sub ResetLabelActivations()

    Call SetLabelStyle(lblListModify, False)
    Call SetLabelStyle(lblListDelete, False)
    Call SetLabelStyle(lblListAdd, False)
    
End Sub

Private Sub lstList_DblClick()
    lblListModify_Click
End Sub

Private Sub lstList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetLabelActivations
End Sub


Private Function SetLabelStyle(ByRef rlbl As Label, ByVal vIsMouseOver As Boolean)

    If vIsMouseOver Then
    
        ResetLabelActivations
    
        rlbl.ForeColor = vbRed
    
    Else
     
        rlbl.ForeColor = vbButtonText
     
    End If

End Function

