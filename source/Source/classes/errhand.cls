VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function GlobalErrorHandler(Optional ByVal vsErrorLocation As Variant, _
                  Optional ByVal vbDisplayError As Variant) As Integer
Dim sMessage As String
Dim nUserResponce As Integer, nResult As Integer
  
  Err.Raise Err.Number, Err.Source, Err.Description
  
  If Not IsMissing(vsErrorLocation) Then
    sMessage = "Error " & Err.Number & " at " & vsErrorLocation & ": " & Error$
  Else
    sMessage = "Error " & Err.Number & ": " & Error$
  End If
  
  nUserResponce = MsgBox(sMessage, vbAbortRetryIgnore + vbCritical)
  Select Case nUserResponce
    Case vbRetry
      nResult = gnERROR_RESUME
    Case vbIgnore
      nResult = gnERROR_RESUME_NEXT
    Case vbAbort
      nResult = gnERROR_EXIT_APP
  End Select
  GlobalErrorHandler = nResult
End Function


