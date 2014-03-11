Attribute VB_Name = "modReport"
Option Explicit


Public Sub MakeStatusReport(ByVal vboolReportForToday As Boolean, Optional ByVal vdtFromDate As Date, Optional ByVal vdtToDate As Date, Optional ByVal vboolIncludeEntryExitTime As Boolean, Optional ByVal vboolIncludeProjectName As Boolean, Optional ByVal vboolIncludeNotes As Boolean, Optional frm As Form = Nothing)

    On Error GoTo ERR_MakeStatusReport

    'This constatnts are useful in late binding
    Const wdWindowStateMaximize = 1
    Const wdListApplyToWholeList = 0
    Const wdNumberParagraph = 1
    Const wdNumberGallery = 2

    Dim vntData As Variant
    Dim sFromDate As String
    Dim sToDate As String
    Dim sSQL As String
    Dim lRowCount As Long
    Dim lRowIndex As Long
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.
    'Dim oWordApplication As Word.Application '--FOR DEBUG MODE--
    Dim oWordApplication As Object
    
    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)
    
    
    gbStatusReportWasGenerated = True
    
    '
    'First set up From and To dates
    '
    
    'If Today's status report is to be generated then
    If vboolReportForToday Then
    
        'Get data for dates starting from today
        sFromDate = Format(Date, "dd mmm yyyy")
        
        'And tommorow
        sToDate = Format(DateAdd("d", 1, sFromDate), "dd mmm yyyy")
        
    Else
        
        'Else start & end date is as specified by usr
        sFromDate = Format(vdtFromDate, "dd mmm yyyy")
        
        sToDate = Format(vdtToDate, "dd mmm yyyy")
    
    End If
    
    'Following is for debug purpose
    'sDate = "28 sep 1998"
    
    '
    'Now setup sSQL
    '
    
    If vboolReportForToday Then
        
        sSQL = GetSQLForStatusReport(sFromDate, sToDate)
        
    Else
    
        sSQL = GetSQLForActivityReport(sFromDate, sToDate)
    
    End If
    
    'Get the data using sSQL
    Call GetRowsInArr(sSQL, vntData, lRowCount)
    
    'If number of rows got is not zero
    If lRowCount <> 0 Then
        
        'Remove the previous Word obj initiated if any
        Set oWordApplication = Nothing
        
        'Create new instance
        Set oWordApplication = CreateObject("Word.Application")
        'Set oWordApplication = New Word.Application
        
        'If Word successfully initiated then
        If Not (oWordApplication Is Nothing) Then
        
            With oWordApplication
                
                'Make the Word visible
                .Visible = True
                
                'This parameter is now not used but is kept any how
                If Not (frm Is Nothing) Then
                
                    Call MakeFormNonTopMost(frm)
                    
                End If
                
                'Maximize the Word window
                .WindowState = wdWindowStateMaximize
                
                'Make the Word window the Active window
                .Activate
                
                'Start new Word document
                .Documents.Add
                
                'If report is to be generated for today
                If vboolReportForToday Then
                    
                    'Call function to write report for today
                    Call WriteReportForToday(oWordApplication, vntData)
                    
                Else
                    
                    'Call function to write activity report
                    Call WriteDetailedReport(oWordApplication, vntData, sFromDate, sToDate, vdtFromDate, vdtToDate, vboolIncludeEntryExitTime, vboolIncludeProjectName, vboolIncludeNotes)
                
                End If
                
                'Make me top again! --- this is not used now
                If Not (frm Is Nothing) Then
                
                    Call MakeFormTopMost(frm)
                    
                End If
            
            End With
            
        Else
            
            'Creation of Word obj has failed so show the error
            MsgBox "This feature requires the Microsoft Word 6.0 or higher. AutoActuals was not able to initiate the Word. The probable reason could be that you have not installed Microsoft Word 6.0 or heigher or current installation is currupted or another application has locked the Word."
        
        End If
        
    Else
        
        'No data fetched for given dates
        MsgBox "No activity found on " & sFromDate
        
    End If
    
    'Destroy the Word obj - this doesn't closes the Word application
    Set oWordApplication = Nothing
    
    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)
    
Exit Sub
ERR_MakeStatusReport:

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)
    
    Set oWordApplication = Nothing
    
    MsgBox "There was an error occured while creating the status report: " & vbCrLf & Err.Number & " : " & Err.Description
    
    Exit Sub

End Sub

Public Sub ShowStatusReportDialog()

    With frmMain

        .opnReportForToday.Caption = "&Generate today's status report"
        
        .opnReportForToday.Value = vbChecked
        
        .txtFromDate = Format(Date, "dd mmm yyyy")
        
        .txtToDate = Format(DateAdd("d", 1, .txtFromDate), "dd mmm yyyy")
        
        .EnableActivityReportOptionControls (False)
    
        .Show
        
    End With

End Sub

Private Function GetSQLForStatusReport(ByVal vsFromDate As String, ByVal vsToDate As String) As String

    Dim sSQL As String
    
    sSQL = "SELECT Tasks.Task_Name, SUM(Activities.Time_Spent), MIN(Activities.Entry_Time)"
    sSQL = sSQL & " FROM Tasks, Activities"
    sSQL = sSQL & " WHERE ("
    sSQL = sSQL & " (Tasks.Task_ID = Activities.Task_ID)"
    sSQL = sSQL & " AND (Activities.Entry_Time >= #" & vsFromDate & "#)"
    sSQL = sSQL & " AND (Activities.Entry_Time < #" & vsToDate & "#)"
    sSQL = sSQL & " AND (Not ((Tasks.Task_Name = '" & gsSTART_UP_TASK_NAME & "') "
    sSQL = sSQL & " AND (Tasks.Task_Name = '" & gsCLOSING_TASK_NAME & "')))"
    sSQL = sSQL & " )"
    sSQL = sSQL & " GROUP BY Tasks.Task_Name"
    sSQL = sSQL & " ORDER BY 3"
    
    GetSQLForStatusReport = sSQL

End Function

Private Function GetSQLForActivityReport(ByVal vsFromDate As String, ByVal vsToDate As String) As String

    Dim sSQL As String
    
    sSQL = "SELECT Tasks.Task_Name, SUM(Activities.Time_Spent), MIN(Activities.Entry_Time), Task_Categories.Task_Category_Name"
    sSQL = sSQL & " FROM Tasks, Activities, Task_Categories"
    sSQL = sSQL & " WHERE ("
    sSQL = sSQL & " (Tasks.Task_ID = Activities.Task_ID)"
    sSQL = sSQL & " AND (Tasks.Task_Category_ID = Task_Categories.Task_Category_ID)"
    sSQL = sSQL & " AND (Activities.Entry_Time >= #" & vsFromDate & "#)"
    sSQL = sSQL & " AND (Activities.Entry_Time < #" & vsToDate & "#)"
    sSQL = sSQL & " AND (Not ((Tasks.Task_Name = '" & gsSTART_UP_TASK_NAME & "') "
    sSQL = sSQL & " AND (Tasks.Task_Name = '" & gsCLOSING_TASK_NAME & "')))"
    sSQL = sSQL & " )"
    sSQL = sSQL & " GROUP BY Tasks.Task_Name, Task_Categories.Task_Category_Name"
    sSQL = sSQL & " ORDER BY 3"
    
    GetSQLForActivityReport = sSQL

End Function

Private Sub WriteReportForToday(oWordApplication As Object, vntData As Variant)

    'This constatnts are useful in late binding
    Const wdWindowStateMaximize = 1
    Const wdListApplyToWholeList = 0
    Const wdNumberParagraph = 1
    Const wdNumberGallery = 2
    
    Dim lRowIndex As Long
    Dim sNameForRegards As String
    Dim nFirstColumn As Integer
    
    
    'Determine the first column number in the array
    nFirstColumn = LBound(vntData, 1)
    
    'Get the name of user to be write in Regards section
    sNameForRegards = GetUserDisplayName

    
    With oWordApplication
        
        'Scheduled tasks completed
        .Selection.Font.Bold = True
        .Selection.Font.Size = .Selection.Font.Size + 2
        
        .Selection.TypeText Text:="Scheduled tasks completed:"
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10
        
        .Selection.TypeParagraph
                    
        .Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=.Parent.ListGalleries( _
            wdNumberGallery).ListTemplates(1), ContinuePreviousList:=False, ApplyTo:= _
            wdListApplyToWholeList
        
        'Scheduled tasks entries
        For lRowIndex = LBound(vntData, 2) To UBound(vntData, 2)
            
            'Create new row
'                    .Selection.TypeText Text:=AlternateStrIfNull(vntData(nFirstColumn, lRowIndex), "<No task name entered>") _
'                        & vbTab & " - " & vbTab & AlternateStrIfNull(vntData(nFirstColumn + 1, lRowIndex), "<No Time Spent entry available>") & " min"
            .Selection.TypeText Text:=AlternateStrIfNull(vntData(nFirstColumn, lRowIndex), "<No task name entered>") _
                & vbTab & " - " & vbTab & _
                MinToHrMin(vntData(nFirstColumn + 1, lRowIndex), , , "No data avalable about time spent", 5)
        
            .Selection.TypeParagraph
        
        Next lRowIndex
        
        'Endup the list
        .Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
        
        
        'heading 2
        
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = True
        .Selection.Font.Size = .Selection.Font.Size + 2
        
        .Selection.TypeText Text:="Unscheduled tasks completed:"
        
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10
        
        .Selection.TypeText Text:="<none>"
        
        
        'heading 3

        .Selection.TypeParagraph
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = True
        .Selection.Font.Size = .Selection.Font.Size + 2
        
        .Selection.TypeText Text:="Scheduled Tasks not completed:"
        
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10
        
        .Selection.TypeText Text:="<none>"
        
        'heading 4

        .Selection.TypeParagraph
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = True
        .Selection.Font.Size = .Selection.Font.Size + 2
        
        .Selection.TypeText Text:="Issues:"
        
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10
        
        .Selection.TypeText Text:="<none>"
        
        
        'heading 5

        .Selection.TypeParagraph
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = True
        .Selection.Font.Size = .Selection.Font.Size + 2
        
        .Selection.TypeText Text:="Risks:"
        
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10
        
        .Selection.TypeText Text:="<none>"
        
        
        'Regards
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10

        .Selection.TypeText Text:="Regards,"
        .Selection.TypeParagraph
        
        sNameForRegards = GetUserDisplayName
        
        .Selection.TypeText Text:=sNameForRegards & "."
        
    End With

End Sub

Private Sub WriteDetailedReport(oWordApplication As Object, vntData As Variant, ByVal vsFromDate As String, ByVal vsToDate As String, ByVal vdtFromDate As Date, ByVal vdtToDate As Date, ByVal vboolIncludeEntryExitTime As Boolean, ByVal vboolIncludeProjectName As Boolean, ByVal vboolIncludeNotes As Boolean)

    'This constatnts are useful in late binding
    Const wdWindowStateMaximize = 1
    Const wdListApplyToWholeList = 0
    Const wdNumberParagraph = 1
    Const wdNumberGallery = 2

    Dim lRowIndex As Long
    Dim nFirstColumn As Integer
    Dim sUserNameForDisplay As String
    Dim sStringForTask As String
    
    
    'Get the number of first column in passed data
    nFirstColumn = LBound(vntData, 1)
    
    'Get the user's name to be displayed
    sUserNameForDisplay = GetUserDisplayName

    With oWordApplication
    
        .Selection.Font.Bold = True
        .Selection.Font.Size = .Selection.Font.Size + 2
        
        'Task list heading
        .Selection.TypeText Text:="Following tasks were carried out by " _
            & sUserNameForDisplay _
            & " between " & vsFromDate & " and " & vsToDate
        
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10
        
        .Selection.TypeParagraph
                    
        .Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=.Parent.ListGalleries( _
            wdNumberGallery).ListTemplates(1), ContinuePreviousList:=False, ApplyTo:= _
            wdListApplyToWholeList
        
        'Write tasks
        For lRowIndex = LBound(vntData, 2) To UBound(vntData, 2)
            
            If vboolIncludeProjectName Then
            
                sStringForTask = vntData(nFirstColumn + 3, lRowIndex) & " : "
                
            Else
            
                sStringForTask = ""
            
            End If
            
            sStringForTask = sStringForTask & AlternateStrIfNull(vntData(nFirstColumn, lRowIndex), "<No task name entered>") _
                & vbTab & " - " & vbTab _
                & MinToHrMin(vntData(nFirstColumn + 1, lRowIndex), , , "No data avalable about time spent", 5)
                
            Call .Selection.TypeText(Text:=sStringForTask)
        
            .Selection.TypeParagraph
        
        Next lRowIndex
        
        'Endup the list
        .Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
        
        'Include Entry-Exit report
        If vboolIncludeEntryExitTime Then
            
            'Entry-Exit time analysis
            Call WriteEntryExitReport(oWordApplication, vsFromDate, vsToDate)
            
        End If
        
    End With

End Sub

Private Sub WriteEntryExitReport(oWordApplication As Object, ByVal vsFromDate As String, ByVal vsToDate As String)

    'This constatnts are useful in late binding
    Const wdWindowStateMaximize = 1
    Const wdListApplyToWholeList = 0
    Const wdNumberParagraph = 1
    Const wdNumberGallery = 2
    Const wdCell = 12
    Const wdTableFormatSimple2 = 2
    Const wdRed = 6
    Const wdAuto = 0
    
    Dim sSQL As String
    Dim lRowCount As Long
    Dim lRowIndex As Long
    Dim vntData As Variant
    Dim lStartedTaskID As Long
    Dim lStoppedTaskID As Long
    Dim colStartTimes As New Collection
    Dim colStopTimes As New Collection
    
    '
    'Step 1 - Get Start/Stop task ID
    '
    
    'Get teh Task IDs of <Started> and <Stopped> tasks
    lStartedTaskID = GetTaskID(gsSTART_UP_TASK_NAME)
    
    lStoppedTaskID = GetTaskID(gsCLOSING_TASK_NAME)
    
    
    
    '
    'Step 2 - Get all start times in collection
    '
    
    'Get all raws for <Started> tasks occuring between the dates
    sSQL = "SELECT Activities.Entry_Time"
    sSQL = sSQL & " FROM Activities"
    sSQL = sSQL & " WHERE ("
    sSQL = sSQL & " (Activities.Entry_Time >= #" & vsFromDate & "#)"
    sSQL = sSQL & " AND (Activities.Entry_Time < #" & vsToDate & "#)"
    sSQL = sSQL & " AND (Activities.Task_ID = " & lStartedTaskID & ")"
    sSQL = sSQL & " )"
    sSQL = sSQL & " ORDER BY Activities.Entry_Time" 'This is needed for calculation algo to work
    
    'Get the data using sSQL
    Call GetRowsInArr(sSQL, vntData, lRowCount)
    
    'Get first/last start/stop time in coll
    Call GetStartStopTimeInColl(vntData, colStartTimes)
    
    
    
    '
    'Step 3 - Get all stop times in collection
    '
    
    'Free the memory
    vntData = Empty
    
    'Get all raws for <Stopped> tasks occuring between the dates
    sSQL = "SELECT Activities.Entry_Time"
    sSQL = sSQL & " FROM Activities"
    sSQL = sSQL & " WHERE ("
    sSQL = sSQL & " (Activities.Entry_Time >= #" & vsFromDate & "#)"
    sSQL = sSQL & " AND (Activities.Entry_Time < #" & vsToDate & "#)"
    sSQL = sSQL & " AND (Activities.Task_ID = " & lStoppedTaskID & ")"
    sSQL = sSQL & " )"
    sSQL = sSQL & " ORDER BY Activities.Entry_Time DESC" 'This is needed for calculation algo to work
    
    'Get the data using sSQL
    Call GetRowsInArr(sSQL, vntData, lRowCount)
    
    'Get first/last start/stop time in coll
    Call GetStartStopTimeInColl(vntData, colStopTimes)
    
    
    
    '
    'Step 4 - calculate req quantities
    '

    Dim dtLatestStartTime As Date
    Dim dtEarliestStartTime As Date
    Dim dtLatestStopTime As Date
    Dim dtEarliestStopTime As Date
    Dim dtAvgStartTime As Date
    Dim dtAvgStopTime As Date
    
    dtLatestStartTime = GetMaxTimeValInCol(colStartTimes)
    dtEarliestStartTime = GetMinTimeValInCol(colStartTimes)
    dtLatestStopTime = GetMaxTimeValInCol(colStopTimes)
    dtEarliestStopTime = GetMinTimeValInCol(colStopTimes)
    dtAvgStartTime = GetAvgTimeInCol(colStartTimes)
    dtAvgStopTime = GetAvgTimeInCol(colStopTimes)
   
    
    With oWordApplication
        
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = True
        .Selection.Font.Size = .Selection.Font.Size + 2
        
        .Selection.TypeText Text:="Entry-Exit time analysis:"
        
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10
        
        .Selection.TypeText Text:="Earliest entry time : " & FormattedTime(dtEarliestStartTime)
        .Selection.TypeParagraph
        .Selection.TypeText Text:="Latest entry time : " & FormattedTime(dtLatestStartTime)
        .Selection.TypeParagraph
        .Selection.TypeText Text:="Earliest exit time : " & FormattedTime(dtEarliestStopTime)
        .Selection.TypeParagraph
        .Selection.TypeText Text:="Latest exit time : " & FormattedTime(dtLatestStopTime)
        .Selection.TypeParagraph
        .Selection.TypeText Text:="Average entry time : " & FormattedTime(dtAvgStartTime)
        .Selection.TypeParagraph
        .Selection.TypeText Text:="Average exit time : " & FormattedTime(dtAvgStopTime)
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        
        .Selection.Font.Bold = True
        .Selection.Font.Size = .Selection.Font.Size + 1
        .Selection.TypeText Text:="Details:"
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.Font.Bold = False
        .Selection.Font.Size = 10
        
        .Selection.Tables.Add Range:=.Selection.Range, NumRows:=2, NumColumns:= _
        3
        .Selection.Tables(1).AutoFormat Format:=wdTableFormatSimple2, ApplyBorders _
        :=True, ApplyShading:=True, ApplyFont:=True, ApplyColor:=True, _
        ApplyHeadingRows:=True, ApplyLastRow:=False, ApplyFirstColumn:=True, _
        ApplyLastColumn:=False, AutoFit:=True
        .Selection.TypeText Text:="Date"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:="Entry Time"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:="Exit Time"
        .Selection.MoveRight Unit:=wdCell
        
        Dim vntColStartTimeItem As Variant
        Dim vntStopTime As Variant
        Dim lTotalMinSpent As Long
        Dim lNumOfDays As Long
        
        lTotalMinSpent = 0
        lNumOfDays = 0
        
        For Each vntColStartTimeItem In colStartTimes
            
            .Selection.TypeText Text:=GetFormattedDatePart(vntColStartTimeItem)
            .Selection.MoveRight Unit:=wdCell
            
            .Selection.TypeText Text:=FormattedTime(vntColStartTimeItem)
            .Selection.MoveRight Unit:=wdCell
            
            
            vntStopTime = GetColItem(colStopTimes, GetFormattedDatePart(vntColStartTimeItem))
            
            If IsEmpty(vntStopTime) Then
                
                .Selection.TypeText Text:="N/A"
                
            Else
                
                .Selection.TypeText Text:=FormattedTime(vntStopTime)
                
                lTotalMinSpent = lTotalMinSpent + DateDiff("n", vntColStartTimeItem, vntStopTime)
                
                lNumOfDays = lNumOfDays + 1
                
            End If
            
            .Selection.MoveRight Unit:=wdCell
            
        Next vntColStartTimeItem
        
        
        .Selection.SelectRow
        .Selection.Cut
        .Selection.TypeParagraph
        
        If lNumOfDays <> 0 Then
        
            Dim lAvgMinutes As Long
            
            lAvgMinutes = lTotalMinSpent / lNumOfDays
            
            .Selection.TypeText Text:="Average time spent/day : " & MinToHrMin(lAvgMinutes, , , , 5)
            .Selection.TypeParagraph
            .Selection.Font.Size = 8
            .Selection.Font.ColorIndex = wdRed
            .Selection.TypeText Text:="(This may not be accurate because of unrecorded true stop times due to improper system shutdowns)"
            .Selection.Font.ColorIndex = wdAuto
            
        Else
        
            .Selection.TypeText Text:="(Average time spent/day not calculated because there was no days found having Start and Stop time recorded properly.)"
        
        End If
        
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.Font.Size = 12
        .Selection.Font.Bold = True
        .Selection.TypeText Text:="Report generated on " & Format$(Now, "dd mmm, yyyy hh:mm AM/PM") & " by AutoActuals Ver. " & App.Major & "." & App.Minor & "." & App.Revision
        
    End With

End Sub


Private Sub GetStartStopTimeInColl(ByVal vavntDateTimes As Variant, ByRef roclStartStopTimes As Collection)

    Dim nArrIndex As Integer
    Dim dtCompare As Date
    Dim nFirstCol As Integer
    Dim dtDateInArr As Date

    Call ClearCollection(roclStartStopTimes)
    
    dtCompare = CDate(0)
    
    If Not IsArrayEmpty(vavntDateTimes) Then
    
        nFirstCol = LBound(vavntDateTimes, 1)
    
        For nArrIndex = LBound(vavntDateTimes, 2) To UBound(vavntDateTimes, 2)
        
            dtDateInArr = vavntDateTimes(nFirstCol, nArrIndex)
            
            If dtCompare <> GetDatePart(dtDateInArr) Then
            
                dtCompare = GetDatePart(dtDateInArr)
                
                Call roclStartStopTimes.Add(dtDateInArr, GetFormattedDatePart(dtCompare))
            
            End If
         
        Next nArrIndex
    
    End If

End Sub

Private Function GetAvgTimeInCol(oclCol As Collection) As Date

    Dim nColIndex As Integer
    Dim lTotalTimeInSec As Long
    Dim lAvgTimeInSec As Long
    Dim dtDatePart As Date
    Dim dtColItem As Date
    
    
    lTotalTimeInSec = 0
    
    For nColIndex = 1 To oclCol.Count
    
        dtColItem = oclCol.Item(nColIndex)
        
        dtDatePart = GetDatePart(dtColItem)
    
        lTotalTimeInSec = lTotalTimeInSec + DateDiff("s", dtDatePart, dtColItem)
    
    Next nColIndex
    
    If oclCol.Count > 1 Then
    
        lAvgTimeInSec = lTotalTimeInSec / oclCol.Count
        
    Else
    
        lAvgTimeInSec = lTotalTimeInSec
    
    End If
    
    
    Dim dtTemp As Date
    
    dtTemp = GetDatePart(Now)
    
    dtTemp = DateAdd("s", lAvgTimeInSec, dtTemp)
    
    
    GetAvgTimeInCol = GetTimePart(dtTemp)

End Function

Private Function GetMaxTimeValInCol(oclCol As Collection) As Variant

    Dim nColIndex As Integer
    Dim vntLastColItem As Variant
    Dim vntReturn As Variant
    
    vntReturn = Empty
    
    If oclCol.Count <> 0 Then
    
        vntLastColItem = GetTimePart(oclCol.Item(1))
    
        For nColIndex = 2 To oclCol.Count
        
            If vntLastColItem < GetTimePart(oclCol.Item(nColIndex)) Then
            
               vntLastColItem = GetTimePart(oclCol.Item(nColIndex))
            
            End If
        
        Next nColIndex
        
    End If
    
    vntReturn = vntLastColItem
    
    GetMaxTimeValInCol = vntReturn

End Function


Private Function GetMinTimeValInCol(oclCol As Collection) As Variant

    Dim nColIndex As Integer
    Dim vntLastColItem As Variant
    Dim vntReturn As Variant
    
    vntReturn = Empty
    
    If oclCol.Count <> 0 Then
    
        vntLastColItem = GetTimePart(oclCol.Item(1))
    
        For nColIndex = 2 To oclCol.Count
        
            If vntLastColItem > GetTimePart(oclCol.Item(nColIndex)) Then
            
               vntLastColItem = GetTimePart(oclCol.Item(nColIndex))
            
            End If
        
        Next nColIndex
        
    End If
    
    vntReturn = vntLastColItem

    GetMinTimeValInCol = vntReturn

End Function

Public Function IsStatusReportReq() As Boolean

    Dim bReturn As Boolean
    
    bReturn = False

    'Check id status report is enforced
    If frmMain.mnuEnforceStatusReport.Checked Then
    
        'If shuting down after 6:30 PM
        If GetTimePart(Now) >= CDate("6:30 PM") Then
        
            If Not gbStatusReportWasGenerated Then
            
                If MsgBox("You haven't prepared today's status report. Do you want to prepare it now?", vbYesNo, "Prepare Status Report?") = vbYes Then
                
                    'Any non 0 value stops closing
                    bReturn = True
                    
                    'Call ShowStatusReportDialog
                    ShowStatusReportDialog
                    
                End If
            
            End If
        
        End If
    
    End If
    
    IsStatusReportReq = bReturn

End Function
