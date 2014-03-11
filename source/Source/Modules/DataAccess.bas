Attribute VB_Name = "modDataAccess"
Option Explicit


'
'==========================================================================================
' Routine Name : PopulateCombo
' Purpose      : Populates combo box from column in database with key set to ItemData in list box
' Parameters   : rcbo: Combo to be field in
'                vsTableName: Table name in which data column is there
'                vsPrimaryKeyCol: Name of column which is primary key
'                vsDataCol: Name of column from which data will be filled in
'                vsCriteriaCol,vsCriteriaValCol: Optionally you can specify WHERE condition for data to be filled - column name to searched and value to searched
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 27-Jul-1998 06:19 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function PopulateCombo(ByRef rcbo As ComboBox, ByVal vsSQL As String) As Boolean

    On Error GoTo ERR_PopulateCombo

    'Routine specific local vars here
    Dim raRows As Variant
    Dim sSQL As String
    Dim lRowCount As Long
    

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "modDataAccess.PopulateCombo"
    
    rcbo.Clear
    
    GetRowsInArr vsSQL, raRows, lRowCount
    
    If lRowCount <> 0 Then
        
        Call FillListFromArray(rcbo, raRows)
        
    End If
    

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    PopulateCombo = bSuccess

Exit Function

ERR_PopulateCombo:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


'==================================================
' Routine Name : GetRecordCount
' Purpose      : Gets the recordcount of the resultset
' Inputs       : Nil
' Assumes      : The resultset contains rows
' Returns      : True: if successful; False: if not
' Effects      : Nil
' Date written : 19/9/97
' Author       : Ritesh
' Revision History
' Date/Time     Person      Action
' 05/11/97      Vikram      The Abs value of Row count is returned. If the Rowcount exceeds 98, it starts returing rowcount with a negative (-) prefix.
'==================================================
Function GetRecordCount(rsResultset As rdoResultset) As Long
    On Error GoTo Err_GetRecordCount
    
    Dim sErrorLocation              ' Stores location of error
    Dim lCount As Long
    Dim vntBookmark As Variant
    Dim bIsBookmarkValid As Boolean
    

    sErrorLocation = "DataAccess.GetRecordCount"
    
    With rsResultset
    
        If .EOF And .BOF Then
            
            lCount = 0
            
        Else
        
            bIsBookmarkValid = (Not .BOF) Or (Not .EOF)
            
            If bIsBookmarkValid Then
                                
                vntBookmark = .Bookmark
            
            End If
            
            If Not .BOF Then .MoveFirst
            
            If Not .EOF Then .MoveLast
            
            lCount = .RowCount
            
            'If RowCount property failed then
            If lCount = -1 Then
            
               lCount = 0
               
               .MoveFirst
               
               Do While Not .EOF
               
                    lCount = lCount + 1
                    
                    .MoveNext
               
               Loop
            
            End If
            
            
            If bIsBookmarkValid Then
                                
                .Bookmark = vntBookmark
            
            End If
                    
        
        End If
            
    End With
    
    If lCount < 0 Then
    
        lCount = 0
    
    End If
    
    GetRecordCount = lCount
    
    ' move to the last row to get the row count
    'rsResultset.MoveLast
    
    ' return to the first row
    rsResultset.MoveFirst
                    
    'GetRecordCount = Abs(rsResultset.RowCount)
    
    Exit Function
            
Err_GetRecordCount:

    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If
    
    Exit Function

End Function


'==================================================
' Routine Name : IsRowExist
' Purpose      : Check whether a record with given value in the given field exists in the table
' Inputs       : Nil
' Assumes      : The last record is to be fetched from the VERSION table
' Returns      : True: if successful; False: if not
' Effects      : Nil
' Date written : 10/9/97
' Author       : Ritesh
' Revision History
' Date/Time     Person     Action
'==================================================

Function IsRowExist(rsResultset As rdoResultset) As Boolean
    On Error GoTo Err_IsRowExist
    Dim sErrorLocation              ' Stores location of error

    sErrorLocation = "DataAccess.IsRowExist"

    '  We have to check whether the record pointer is at the end of file or at the beginning of the file
    If Not rsResultset.EOF And Not rsResultset.BOF Then
        IsRowExist = True
    End If
        
    Exit Function
    
Err_IsRowExist:
    
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If
    
    Exit Function
End Function

Function GetID(ByVal vsTableName As String, ByVal vsPrimaryKeyCol As String, ByVal vsSearchCol As String, ByVal vsSearchFor As String, ByRef rlID As Long) As Boolean
    
    On Error GoTo ERR_GetID
    
    Dim sSQL As String
    Dim lRowCount As Long
    Dim raRows As Variant
    
    sSQL = "SELECT " & vsPrimaryKeyCol
    
    sSQL = sSQL & " FROM " & vsTableName
    
    sSQL = sSQL & " WHERE " & vsSearchCol & " = '" & vsSearchFor & "'"
    
    GetRowsInArr sSQL, raRows, lRowCount
    
    If lRowCount <> 0 Then
        
        rlID = raRows(0, 0)
        
        GetID = True
    
    Else
    
        GetID = False
    
    End If
    
Exit Function
ERR_GetID:

    GetID = False
    
    Exit Function
    
End Function


'
'==========================================================================================
' Routine Name : GetRowsInArr
' Purpose      : Queries the database with specified query and stores result in passed variant as 2 dimentional array
' Parameters   : vsSQL: SELECT query which returns rows
'                rvntRows: Result array of query is placed in this variant
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 27-Jul-1998 08:31 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function GetRowsInArr(ByVal vsSQL As String, ByRef rvntRows As Variant, ByRef rlRowCount As Long) As Boolean

    On Error GoTo ERR_GetRowsInArr

    'Routine specific local vars here
    Dim en As rdoEnvironment
    Dim cn As rdoConnection
    Dim rs As rdoResultset
    Dim sConnect As String

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "modDataAccess.GetRowsInArr"


    sConnect = "DSN=" & gsDSN & ";UID=;PWD="
    
    Set en = rdoEngine.rdoEnvironments(0)
    
    Set cn = en.OpenConnection("", rdDriverNoPrompt, False, sConnect)

    
    Set rs = cn.OpenResultset(vsSQL, rdOpenKeyset, rdConcurLock)

    rlRowCount = GetRecordCount(rs)
    
    If rlRowCount <> 0 Then
    
        rvntRows = rs.GetRows(rlRowCount)
        
    Else
    
        rvntRows = Empty
    
    End If
    

    
    rs.Close
    
    cn.Close
    
    en.Close


    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    GetRowsInArr = bSuccess

Exit Function

ERR_GetRowsInArr:

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
' Routine Name : FillListFromArray
' Purpose      : Fills given 2-D array with Item ID and Item name in to the list
' Parameters   : rctrl: Control that should be either List Box or Combo Box
'                vavntItems: Array having two column: Item ID and Item Name
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 08-Aug-1998 09:24 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function FillListFromArray(ByRef rctrl As Control, ByVal vavntItems As Variant) As Boolean

    On Error GoTo ERR_FillListFromArray

    'Routine specific local vars here
    Dim nRowIndex As Integer
    Dim nFirstColNum As Integer

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "modDataAccess.FillListFromArray"

    
    rctrl.Clear
    
    If Not IsArrayEmpty(vavntItems) Then
    
        nFirstColNum = LBound(vavntItems, 1)
    
        For nRowIndex = LBound(vavntItems, 2) To UBound(vavntItems, 2)
        
            rctrl.AddItem (vavntItems(nFirstColNum + 1, nRowIndex))
            
            rctrl.ItemData(rctrl.NewIndex) = vavntItems(nFirstColNum, nRowIndex)
        
        Next nRowIndex
        
    End If


    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    FillListFromArray = bSuccess

Exit Function

ERR_FillListFromArray:

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
' Routine Name : FillArrayFromList
' Purpose      : Fills ItemData and item in list from given 2-D array
' Parameters   : rctrl: Control that should be either List Box or Combo Box
'                vavntItems: Array having two column: Item ID and Item Name
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 08-Aug-1998 09:24 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function FillArrayFromList(ByVal vctrl As Control, ByRef ravntItems As Variant) As Boolean

    On Error GoTo ERR_FillArrayFromList

    'Routine specific local vars here
    Dim nRowIndex As Integer

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "modDataAccess.FillArrayFromList"

    
    If IsArray(ravntItems) Then
        
        Erase ravntItems
        
    End If
    
    
    If vctrl.ListCount <> 0 Then
    
        ReDim ravntItems(0 To 1, 0 To vctrl.ListCount - 1)
    
    End If
  
    For nRowIndex = 0 To vctrl.ListCount - 1
    
        ravntItems(0, nRowIndex) = vctrl.ItemData(nRowIndex)
        
        ravntItems(1, nRowIndex) = vctrl.List(nRowIndex)
    
    Next nRowIndex

    
    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    FillArrayFromList = bSuccess

Exit Function

ERR_FillArrayFromList:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


'==================================================
' Routine Name : GetNewKey
' Purpose      : Gets a new incremented key from the key factory table
'              : This routine uses a separate connection to be used for multiuser considerations
' Inputs       : vsTableName: The table whose incremented key it needs
' Assumes      : The table name entry is already there in the key factory
' Returns      : True: if successful; False: if not
' Effects      : Nil
' Date written : 20/9/97
' Author       : Ritesh
' Revision History
' Date/Time     Person     Action
'==================================================

Function GetNewKey(ByRef rConn As rdoConnection, ByVal vsTableName As String, ByRef rlkey As Long) As Boolean
    On Error GoTo Err_GetNewKey
    
    Dim bSuccess As Boolean             ' The flag which determines whether the function is successful
    
    Dim lkey As Long                    ' The new key value
    Dim sSQL As String                  ' The Sql statement
    Dim rsResultset As rdoResultset     ' The Record in the Key Factory
    
    Dim nCount As Integer               ' Maintains the count in case of the record being locked and
                                        ' it needs to repeat till the record is released
        
    
    Dim sErrorLocation              ' Stores location of error

    sErrorLocation = "DataAccess.GetNewKey"
    
    
    ' Make the select query to open the record for a particular table
    sSQL = "SELECT * FROM KEY_FACTORY WHERE TABLE_NAME = '" & vsTableName & "'"
    
    ' Start the transaction
    'rConn.BeginTrans
    
    ' Open the resultset
    Set rsResultset = rConn.OpenResultset(sSQL, rdOpenDynamic, rdConcurLock)
    
    ' set the resultset in edit mode
    ' If the editing failed because of other user having locked it implicitly
    ' keep retrying unless the other user releases the lock
    ' for a specific period of wait time
    ' If it is not released before that then exit with false value
    If IsRowExist(rsResultset) Then
        rsResultset.Edit
        
        ' get the last entered key in the key factory
        ' and then update the key value
        lkey = rsResultset("NEXT_KEY_VALUE")
        rsResultset("NEXT_KEY_VALUE") = lkey + 1
        
        bSuccess = True
    Else
        ' If no row is existing then go into Add mode
        rsResultset.AddNew
        
        'set the first entry
        lkey = 1
        rsResultset("NEXT_KEY_VALUE") = lkey + 1
        
        'set the first entry
        rsResultset("TABLE_NAME") = vsTableName
        bSuccess = True
        
    End If
        
    If bSuccess Then
        ' update the resultset
        rsResultset.Update
        
        ' return the new key value
        rlkey = lkey
        
    End If
    
    ' Close the recordset
    rsResultset.Close
    
    ' Commit transaction
    'rConn.CommitTrans
    
    GetNewKey = bSuccess
    
    Exit Function
    
Err_GetNewKey:

    Select Case gobjErrors(sErrorLocation)
        Case gnRESUME_NEXT
            Resume Next
    
    End Select
                    
    ' set the resultset to nothing
    Set rsResultset = Nothing
    
    'rConn.RollbackTrans
    
    Exit Function

End Function


Public Function AddARow(ByVal vsTableName As String, ByVal vavntColNames As Variant, ByVal vavntColValues As Variant, Optional ByVal vsGeneratedKeyCol As String = "", Optional ByRef rlKeyValue As Long) As Boolean

    Dim en As rdoEnvironment
    Dim cn As rdoConnection
    Dim rs As rdoResultset
    Dim sConnect As String
    Dim bSuccess As Boolean
    Dim nArrayIndex As Integer

    
    bSuccess = True

    sConnect = "DSN=" & gsDSN & ";UID=;PWD="
    
    Set en = rdoEngine.rdoEnvironments(0)
    
    Set cn = en.OpenConnection("", rdDriverNoPrompt, False, sConnect)
    
    
    If vsGeneratedKeyCol <> "" Then

        bSuccess = GetNewKey(cn, vsTableName, rlKeyValue)
            
    End If

    
    If bSuccess Then
    
        Set rs = cn.OpenResultset("SELECT * FROM " & vsTableName, rdOpenDynamic, rdConcurLock)
    
        rs.AddNew
        
            If vsGeneratedKeyCol <> "" Then
            
                rs(vsGeneratedKeyCol) = rlKeyValue
                
            End If
        
            If bSuccess Then
                    
                For nArrayIndex = LBound(vavntColNames) To UBound(vavntColNames)
                
                    rs(vavntColNames(nArrayIndex)) = vavntColValues(nArrayIndex)
                    
                Next nArrayIndex
                
            End If
            
        rs.Update
        
    End If
    
    rs.Close
    
    cn.Close
    
    en.Close
    
    AddARow = bSuccess
    
End Function


Public Function AddRowsWithConst(ByVal vsTableName As String, ByVal vsConstColName As String, ByVal vvntConstColValue As Variant, ByVal vavntColNames As Variant, ByVal vavntColValues As Variant, ByVal vavntColMapping As Variant, Optional ByVal vIsDeleteBeforeAdd As Boolean = False, Optional ByVal vsGeneratedKeyCol As String = "", Optional ByRef rlKeyValue As Long, Optional ByVal vnUseExistingPrimaryKeyCol As Integer = -1) As Boolean

    On Error GoTo Err_AddRowsWithConst
    
    
    Dim en As rdoEnvironment
    Dim cn As rdoConnection
    Dim rs As rdoResultset
    Dim sConnect As String
    Dim bSuccess As Boolean
    Dim nColNameIndex As Integer
    Dim nRowIndex As Integer
    
    bSuccess = True

    sConnect = "DSN=" & gsDSN & ";UID=;PWD="
    
    Set en = rdoEngine.rdoEnvironments(0)
    
    Set cn = en.OpenConnection("", rdDriverNoPrompt, False, sConnect)
    
    
    If vIsDeleteBeforeAdd Then

        cn.BeginTrans

        Call cn.Execute("DELETE FROM " & vsTableName & " WHERE " & vsConstColName & " = " & vvntConstColValue)

    End If
        
    If GetDimension(vavntColValues) = 2 Then
        
        Set rs = cn.OpenResultset("SELECT * FROM " & vsTableName, rdOpenDynamic, rdConcurLock)
        
        For nRowIndex = LBound(vavntColValues, 2) To UBound(vavntColValues, 2)
        
            'Should the existing Primary key col should be used
            If vnUseExistingPrimaryKeyCol <> -1 Then
            
                'Check if valid primary key exist in the column
                rlKeyValue = vavntColValues(vnUseExistingPrimaryKeyCol, nRowIndex)
                
                'If primary key is invalid then
                If rlKeyValue <= 0 Then
                
                    'Generate it
                    bSuccess = GetNewKey(cn, vsTableName, rlKeyValue)
 
                End If
                
            Else
            
                'Check if key needs to be generated
                If vsGeneratedKeyCol <> "" Then
    
                    bSuccess = GetNewKey(cn, vsTableName, rlKeyValue)
                
                End If
            
            End If
                            
            
            If bSuccess Then
            
                rs.AddNew
                
                    'If is not Auto numbered
                    If vsGeneratedKeyCol <> "" Then
                    
                        'Use the key generated
                        rs(vsGeneratedKeyCol) = rlKeyValue
                    
                    End If
                    
                    rs(vsConstColName) = vvntConstColValue
      
                    For nColNameIndex = LBound(vavntColNames) To UBound(vavntColNames)
                
                        rs(vavntColNames(nColNameIndex)) = vavntColValues(vavntColMapping(nColNameIndex), nRowIndex)
                    
                    Next nColNameIndex
                
                rs.Update
                
            End If
            
        Next nRowIndex
    
        rs.Close
        
    End If
    
    If vIsDeleteBeforeAdd Then
        
       cn.CommitTrans
        
    End If
    
    cn.Close
        
    en.Close

    AddRowsWithConst = bSuccess
    
Exit Function

Err_AddRowsWithConst:
    
    If vIsDeleteBeforeAdd Then

        cn.RollbackTrans

    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function ModifyARow(ByVal vsTableName As String, ByVal vsPrimaryKeyColName As String, ByVal vvntPrimaryKeyVal As Variant, ByVal vavntColNames As Variant, ByVal vavntColValues As Variant) As Boolean

    Dim en As rdoEnvironment
    Dim cn As rdoConnection
    Dim rs As rdoResultset
    Dim sConnect As String
    Dim bSuccess As Boolean
    Dim sSQL As String
    Dim nArrayIndex As Integer

    
    bSuccess = True

    sConnect = "DSN=" & gsDSN & ";UID=;PWD="
    
    Set en = rdoEngine.rdoEnvironments(0)
    
    Set cn = en.OpenConnection("", rdDriverNoPrompt, False, sConnect)
    
    
    If bSuccess Then
    
        sSQL = "SELECT *"
        sSQL = sSQL & " From " & vsTableName
        sSQL = sSQL & " WHERE (" & vsPrimaryKeyColName & " = " & vvntPrimaryKeyVal & ")"
        
        
        Set rs = cn.OpenResultset(sSQL, rdOpenDynamic, rdConcurLock)
    
        rs.Edit
                           
                For nArrayIndex = LBound(vavntColNames) To UBound(vavntColNames)
                
                    rs(vavntColNames(nArrayIndex)) = vavntColValues(nArrayIndex)
                    
                Next nArrayIndex
                    
        rs.Update
        
    End If
    
    rs.Close
    
    cn.Close
    
    en.Close
    
    ModifyARow = bSuccess
    
End Function


Public Function DeleteRows(ByVal vsTableName As String, ByVal vavntColNames As Variant, ByVal vavntColValues As Variant) As Boolean

    Dim en As rdoEnvironment
    Dim cn As rdoConnection
    Dim sConnect As String
    Dim bSuccess As Boolean
    Dim sSQL As String
    
    bSuccess = True

    sConnect = "DSN=" & gsDSN & ";UID=;PWD="
    
    Set en = rdoEngine.rdoEnvironments(0)
    
    Set cn = en.OpenConnection("", rdDriverNoPrompt, False, sConnect)
    
        sSQL = "DELETE"
        sSQL = sSQL & " From " & vsTableName
        sSQL = sSQL & " WHERE"
        
        Call BuildSQLWhereClause(sSQL, vavntColNames, vavntColValues)
        
        cn.Execute sSQL
                
    
    cn.Close
    
    en.Close
    
    DeleteRows = bSuccess
    
End Function



Function BuildSQLWhereClause(ByRef rstrSQL As String, ByVal vavntCriteriaColNames As Variant, ByVal vavntCriteriaColValues As Variant, Optional ByVal vsOperator As String = "AND") As String

    Dim nArrayIndex As Integer
        
    For nArrayIndex = LBound(vavntCriteriaColNames) To UBound(vavntCriteriaColNames)
    
        rstrSQL = rstrSQL & " (" & vavntCriteriaColNames(nArrayIndex) & " = " & vavntCriteriaColValues(nArrayIndex) & ")"
        
        If nArrayIndex < UBound(vavntCriteriaColNames) Then
        
            rstrSQL = rstrSQL & " " & vsOperator

        End If
                    
    Next nArrayIndex

End Function


Function GetColValue(ByVal vsSQL As String, ByRef vntColVal As Variant) As Boolean
    
    On Error GoTo ERR_GetColValue
    
    Dim lRowCount As Long
    Dim raRows As Variant
    
    
    Call GetRowsInArr(vsSQL, raRows, lRowCount)
    
    If lRowCount <> 0 Then
        
        vntColVal = raRows(0, 0)
        
        GetColValue = True
    
    Else
    
        vntColVal = Empty
        
        GetColValue = True
    
    End If
    
Exit Function
ERR_GetColValue:

    GetColValue = False
    
    vntColVal = Empty
    
    Exit Function
    
End Function

