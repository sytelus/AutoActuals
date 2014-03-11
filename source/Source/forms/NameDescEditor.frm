VERSION 5.00
Begin VB.Form frmNameDescEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Name And Description"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "NameDescEditor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   4395
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   780
         Width           =   3255
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   465
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   840
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmNameDescEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbUserPressedCancel As Boolean


Public Function DisplayForm(ByVal vsFormCaption As String, ByRef rsName As String, ByRef rsDescription As String) As Boolean

    Dim bSuccess As Boolean
    Dim bValidateSuccess As Boolean
    
    bSuccess = True
    
    bValidateSuccess = False
   

    bSuccess = InitForm(vsFormCaption, rsName, rsDescription)
        
    If bSuccess Then
    
        Do
    
            Me.Show vbModal
            
            bSuccess = Not mbUserPressedCancel
            
            If bSuccess Then
            
                bValidateSuccess = ValidateForm(rsName)
                
                If bValidateSuccess Then
                
                    bSuccess = SetReturnValues(rsName, rsDescription)
                    
                End If
                       
            End If
            
        Loop Until (bValidateSuccess) Or (mbUserPressedCancel)
    
    End If
    
    DisplayForm = bSuccess
    
    Unload Me

End Function

Private Function InitForm(ByVal vsFormCaption As String, ByVal vsName As String, ByVal vsDescription As String) As Boolean

    Dim bSuccess As Boolean
    
    bSuccess = True
    
        Me.Caption = vsFormCaption
        
        txtName = vsName
        
        txtDescription = vsDescription
    
    InitForm = bSuccess

End Function

Private Function SetReturnValues(ByVal vsName As String, ByVal vsDescription As String) As Boolean
    
    Dim bSuccess As Boolean
    
    bSuccess = True
    
    
        vsName = txtName
        
        vsDescription = txtDescription

    
    SetReturnValues = bSuccess

End Function
Private Function ValidateForm(ByVal vsName As String) As Boolean

    Dim bSuccess As Boolean
    
    bSuccess = True
    
        If Len(txtName) = 0 Then
     
            MsgBox "Name field can not be blank. Please type valid name."
        
            bSuccess = False
        
        End If
    
    
    ValidateForm = bSuccess

End Function

Private Sub btnCancel_Click()
    
    mbUserPressedCancel = True
    
    Me.Hide
    
    DoEvents

End Sub

Private Sub btnOK_Click()
    
    mbUserPressedCancel = False
    
    Me.Hide
    
    DoEvents

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'This is in case user presses the Close button in system menu
    mbUserPressedCancel = True
    
End Sub
