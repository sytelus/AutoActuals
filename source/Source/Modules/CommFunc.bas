Attribute VB_Name = "modCommFunc"
Option Explicit

'Contains commaon functions used by various modules in this project.




'
'==========================================================================================
' Routine Name : SetTrayHint
' Purpose      : Sets the hint that is displayed in tray
' Parameters   : vsHint: Hint text
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 06-Aug-1998 12:03 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function SetTrayHint(ByVal vsHint As String, Optional ByVal vIsInBracket As Boolean = True) As Boolean

    On Error GoTo ERR_SetTrayHint

    'Routine specific local vars here
    Dim sHint As String

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler

    'By default assume everything gone fine
    bSuccess = True

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "modCommFunc.SetTrayHint"


    If vIsInBracket Then
    
        sHint = gsTRAY_HINT & " (" & vsHint & ")"
        
    Else
    
        sHint = vsHint
    
    End If


    Call ModifySysTrayIcon(frmMain, sHint)


    'Return success status of function
    SetTrayHint = bSuccess

Exit Function

ERR_SetTrayHint:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


Public Sub ExitApp(Optional ByVal vsMessage As String = "")
    
    If vsMessage <> "" Then
    
        MsgBox vsMessage
        
    End If
        
        Unload frmMain
        
        End
    
End Sub


Public Function IsVisibleUserExists(ByVal vsUserLoginName As String) As Boolean

        Dim sSQL As String
        Dim vntUserVisible As Variant
        
        sSQL = "SELECT Is_Visible From Users WHERE (User_Login_Name = '" & vsUserLoginName & "') AND (Is_Visible = True)"
        
        Call GetColValue(sSQL, vntUserVisible)

        IsVisibleUserExists = Not IsEmpty(vntUserVisible)
        
End Function

Public Function CheckSecurity(ByVal vnMode As Integer, ByRef rsSecurityAccessDeniyMsg As String) As Integer

    CheckSecurity = gnSECURITY_ALLOW

    Select Case vnMode
    
        'Creating new module in project
        Case gnMODE_MODULE_EDITOR_ADD
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
                'Others are allowed to have personal projects so allow add
            End Select
            
        'Modifieng existing module in project
        Case gnMODE_MODULE_EDITOR_MODIFY
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
                'Others are allowed to have personal projects so allow modify
            End Select

        'Creating new projects
        Case gnMODE_PROJECT_EDITOR_ADD
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
                    rsSecurityAccessDeniyMsg = "You can not add new projects because you do not have required rights."
                Case gnUSER_TYPE_PROGRAMMER, gnUSER_TYPE_SENIER_PROGRAMMER
                    CheckSecurity = gnSECURITY_PARTIAL_READ_ONLY
            End Select
        
        'Modifyieng existing projects
        Case gnMODE_PROJECT_EDITOR_MODIFY
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_READ_ONLY
                Case gnUSER_TYPE_PROGRAMMER, gnUSER_TYPE_SENIER_PROGRAMMER
                    CheckSecurity = gnSECURITY_PARTIAL_READ_ONLY
            End Select
        
        Case gnMODE_TASK_EDITOR_ADD
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
                    rsSecurityAccessDeniyMsg = "You can not add new tasks because you do not have required rights."
                Case gnUSER_TYPE_PROGRAMMER
                    CheckSecurity = gnSECURITY_PARTIAL_READ_ONLY
            End Select
        
        
        Case gnMODE_TASK_EDITOR_MODIFY
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_READ_ONLY
                Case gnUSER_TYPE_PROGRAMMER
                    CheckSecurity = gnSECURITY_PARTIAL_READ_ONLY
            End Select
        
        
        Case gnMODE_TASK_TYPE_EDITOR_ADD
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
                    rsSecurityAccessDeniyMsg = "You can not add new Task Types because you do not have required rights."
                Case gnUSER_TYPE_PROGRAMMER
                    CheckSecurity = gnSECURITY_PARTIAL_READ_ONLY
            End Select
        
        
        Case gnMODE_TASK_TYPE_EDITOR_MODIFY
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_READ_ONLY
                Case gnUSER_TYPE_PROGRAMMER
                    CheckSecurity = gnSECURITY_PARTIAL_READ_ONLY
            End Select
        
        
        Case gnMODE_USER_EDITOR_ADD
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE, gnUSER_TYPE_PROGRAMMER, gnUSER_TYPE_SENIER_PROGRAMMER, gnUSER_TYPE_PROCESS_LEADER
                    CheckSecurity = gnSECURITY_DENIY
                    rsSecurityAccessDeniyMsg = "You can not add new user because you do not have required rights."
            End Select
        
        
        Case gnMODE_USER_EDITOR_MODIFY
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE, gnUSER_TYPE_PROGRAMMER, gnUSER_TYPE_SENIER_PROGRAMMER, gnUSER_TYPE_PROCESS_LEADER
                    CheckSecurity = gnSECURITY_READ_ONLY
            End Select
            
        Case gnMODE_MODULE_EDITOR_DELETE
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
            End Select

        
        Case gnMODE_PROJECT_EDITOR_DELETE
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
                    rsSecurityAccessDeniyMsg = "You can not delete the project because you do not have required rights."
            End Select
        
        
        Case gnMODE_TASK_EDITOR_DELETE
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
                    rsSecurityAccessDeniyMsg = "You can not delete the Task because you do not have required rights."
            End Select
        
        
        Case gnMODE_TASK_TYPE_EDITOR_DELETE
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE
                    CheckSecurity = gnSECURITY_DENIY
            End Select
        
        
        Case gnMODE_USER_EDITOR_DELETE
            Select Case glUser_Type_ID
                Case gnUSER_TYPE_NEWBEE, gnUSER_TYPE_PROGRAMMER, gnUSER_TYPE_SENIER_PROGRAMMER, gnUSER_TYPE_PROCESS_LEADER
                    CheckSecurity = gnSECURITY_DENIY
                    rsSecurityAccessDeniyMsg = "You can not delete the user because you do not have required rights."
            End Select
        
    End Select

End Function


Public Function SaveActivity(ByVal vsActivity As String, ByVal vdtTime As Date, ByVal vlUser_ID As Long, ByVal vsMachineName As String, ByVal vlTask As Long, ByVal vlTimeSpent As Long) As Boolean

    Dim sSQL As String
    Dim bSuccess As Boolean
    Dim lTotalMinSpentInTask As Long
    Dim vntColValue As Variant
    
    bSuccess = True
        
    'Save the activity info
    bSuccess = AddARow("Activities", Array("Entry_Time", "Task_ID", "Activity", "User_ID", "Machine_ID", "Time_Spent", "Is_Visible"), _
        Array(vdtTime, vlTask, vsActivity, glUser_ID, glMachine_ID, vlTimeSpent, True))
        
    If bSuccess Then
        
        'Get the time spent on task
        sSQL = "SELECT Time_Spent FROM Tasks WHERE Task_ID = " & vlTask
        
        Call GetColValue(sSQL, vntColValue)
        
        If IsNumber(vntColValue) Then
        
            lTotalMinSpentInTask = vntColValue + vlTimeSpent
            
        Else
        
            lTotalMinSpentInTask = vlTimeSpent
            
        End If
        
        'Save the total time spent on task
        bSuccess = ModifyARow("Tasks", "Task_ID", vlTask, Array("Time_Spent"), Array(lTotalMinSpentInTask))
        
    End If
    
    SaveActivity = bSuccess
    
End Function



Public Function GetUserDetails(ByVal vlUserID As Long, ByRef rsFirstName As String, ByRef rsLastName As String, ByRef rsLoginName As String) As Boolean

    Dim vntUserDetails As Variant
    Dim nFirstColNum As Integer
    Dim nFirstRowNum As Integer
    Dim sSQL As String
    Dim bSuccess As Boolean
    
    bSuccess = True
    
    sSQL = "SELECT First_Name, Last_Name, User_Login_Name"
    sSQL = sSQL & " FROM Users"
    sSQL = sSQL & " WHERE (User_ID = " & vlUserID & ")"

    Call GetColValue(sSQL, vntUserDetails)
    
    If Not IsArrayEmpty(vntUserDetails) Then
    
        nFirstColNum = LBound(vntUserDetails, 1)
        
        nFirstRowNum = LBound(vntUserDetails, 2)
        
        rsFirstName = vntUserDetails(nFirstColNum, nFirstRowNum)
        
        rsLastName = vntUserDetails(nFirstColNum + 1, nFirstRowNum)
        
        rsLoginName = vntUserDetails(nFirstColNum + 2, nFirstRowNum)
        
    Else
    
        bSuccess = False
    
    End If
    
    GetUserDetails = bSuccess

End Function

Function GetUserDisplayName() As String

    Dim nFirstColumn As Integer
    Dim sFirstName As String
    Dim sLastName As String
    Dim sLoginName As String
    Dim sReturn As String

    'Get users first name, last name, login name
    Call GetUserDetails(glUser_ID, sFirstName, sLastName, sLoginName)
    
    'Generate the user name that could be displayed in report
    If Len(sFirstName) <> 0 Then
        
        sReturn = sFirstName
        
    Else
        
        sReturn = AlternateStrIfNull(sLoginName, gsUserLoginName)
        
    End If
    
    sReturn = StrConv(sReturn, vbProperCase)
    
    GetUserDisplayName = sReturn

End Function

Public Function GetTaskID(ByVal vsTaskName As String) As Long

    Dim lTaskID As Long
    
    If Not GetID("Tasks", "Task_ID", "Task_Name", vsTaskName, lTaskID) Then
     
         Call Err.Raise(vbObjectError + 9999, "GetTaskID", "Task named " & vsTaskName & " must exist in Tasks table (this is the system task).")
     
     End If
     
     GetTaskID = lTaskID

End Function
