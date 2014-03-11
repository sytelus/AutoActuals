Attribute VB_Name = "Utils"
Option Explicit

'==================================================
' Routine Name : SetMousePointer
' Purpose      : Sets the mouse cursor type
' Inputs       :
' Assumes      :
' Returns      : Old mouse pointer
' Effects      :
' Date written : 29/08/98
' Author       : Vikram Aryan
' Revision History
' Date      Person       Action
'=================================================='
' NOTE : Error handler will not be used
'=================================================='
Public Function SetMousePointer(ByVal nCursorType As Integer) As Integer
    Dim nOldMousePointer As Integer
    
    ' get the old mouse pointer
    nOldMousePointer = Screen.MousePointer
    
    ' set the mouse pointer
    Screen.MousePointer = nCursorType
    
    ' return the old mouse pointer
    SetMousePointer = nOldMousePointer
    
End Function '

Public Sub DoExit()
On Error GoTo err_exit
'  Unload frmMain  'give the name of main form in application
'  Exit Sub
  'End
err_exit:
  'End
End Sub

'
'==========================================================================================
' Routine Name : GetPathWithSlash
' Purpose      : Checks if slash already present at the end of path. If not append it.
' Parameters   : a string containing path
' Return       : Path with slash
' Effects      : None
' Assumes      : None
' Author       : shital
' Date         : 09-Apr-1998 02:43 PM
' Template     : Ver.11   Author: Shital Shah   Date: 07 Apr, 1998
' Revision History :
' Date          Person      Details.
'==========================================================================================
'

Public Function GetPathWithSlash(ByVal vsPath As String) As String

    On Error GoTo ERR_GetPathWithSlash

    'Routine specific local vars here

    'common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.GetPathWithSlash"


    If Trim$(vsPath) <> "" Then
    
        If Right$(vsPath, 1) <> "\" Then
            
            GetPathWithSlash = vsPath + "\"
            
        Else
        
            GetPathWithSlash = vsPath
        
        End If

    Else
        
        GetPathWithSlash = ""
        
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_GetPathWithSlash:

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
' Routine Name : GetWindowsPathWithSlash
' Purpose      : Get the Windows directory with slash
' Parameters   : None
' Return       : String containing path of Windows dir
' Effects      : None
' Assumes      : None
' Author       : shital
' Date         : 09-Apr-1998 02:50 PM
' Template     : Ver.11   Author: Shital Shah   Date: 07 Apr, 1998
' Revision History :
' Date          Person      Details.
'==========================================================================================
'

Public Function GetWindowsPathWithSlash() As String

    On Error GoTo ERR_GetWindowsPathWithSlash

    'Routine specific local vars here
    Dim Result As String
    Dim ResultLen As Integer

    'common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.GetWindowsPathWithSlash"

    
    Result = Space$(MAX_PATH)

    ResultLen = GetWindowsDirectory(Result, MAX_PATH)
    
    Result = Left$(Result, ResultLen)
    
    GetWindowsPathWithSlash = GetPathWithSlash(Result)


    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_GetWindowsPathWithSlash:

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
' Routine Name : ExecuteEXE
' Purpose      : Executes the specified EXE
' Parameters   : Name of the EXE file to run
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : shital
' Date         : 09-Apr-1998 03:00 PM
' Template     : Ver.11   Author: Shital Shah   Date: 07 Apr, 1998
' Revision History :
' Date          Person      Details.
'==========================================================================================
'

Public Function ExecuteEXE(ByVal vsEXEFileName As String, Optional ByVal vsArguments As String = "") As Boolean

    On Error GoTo ERR_ExecuteEXE

    'Routine specific local vars here

    'common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'by default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.ExecuteEXE"



    ExecuteEXE = WinExec(vsEXEFileName & " " & vsArguments, SW_SHOW) > 31



    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    ExecuteEXE = bSuccess

Exit Function

ERR_ExecuteEXE:

    'Call the global error handling routine to process the error, and check if execution should be continued

    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function

Public Function Check2Bool(ByVal vnCheckValue As Integer) As Boolean
    If vnCheckValue = 0 Then
        Check2Bool = False
    Else
        Check2Bool = True
    End If
End Function

Public Function GetLoginName(Optional ByVal vsOnNoLoginReturn As String = "<GetUserName API failed-Can not get Login name>") As String
On Error GoTo ERR_GetLoginName

Dim sResult As String
Dim nResultLen As Long

    nResultLen = 50
    
    sResult = Space$(nResultLen)
    
    If GetUserName(sResult, nResultLen) = 0 Then
      Err.Raise 9999, Description:=""
    End If
    
    
    GetLoginName = Left$(sResult, nResultLen - 1)

Exit Function
ERR_GetLoginName:

  GetLoginName = vsOnNoLoginReturn
  
  Exit Function

End Function


Public Sub ShowError(ByVal vsMsg As String, Optional ByVal vnNumber As Integer = -1, Optional ByVal vsLocation As String)
Dim sMsg As String

    sMsg = "Error "
    
    If Not (vnNumber = -1) Then sMsg = sMsg & vnNumber
    
    sMsg = sMsg & " : "
    
    sMsg = sMsg & vsMsg
    
    If Not IsMissing(vsLocation) Then
    
        If Not vsLocation = "" Then
      
            sMsg = sMsg & " at location " & vsLocation
            
        End If
        
    End If
        
   MsgBox sMsg
    
End Sub

Function IsFileExist(ByVal sFileName As String) As Boolean
    IsFileExist = Dir$(sFileName, vbNormal Or vbHidden Or vbSystem) <> ""
End Function
Public Function Capitalize(ByVal vsWord As String) As String

Dim sLeftMostChar As String

    Capitalize = vsWord
    
    If Capitalize <> "" Then
    
        sLeftMostChar = Left$(Capitalize, 1)
    
        If (sLeftMostChar >= "a") And (sLeftMostChar <= "z") Then
        
            Mid(Capitalize, 1) = Chr$(Asc(sLeftMostChar) - Asc("a") + Asc("A"))
        
        End If
    
    End If

End Function

Public Function MakeFormTopMost(frm As Form) As Boolean

    MakeFormTopMost = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Function
Public Function MakeFormNonTopMost(frm As Form) As Boolean

    MakeFormNonTopMost = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Function

Public Function GetMachineName() As String
On Error GoTo ERR_GetMachineName

Dim sResult As String
Dim nResultLen As Long

    nResultLen = 50
    
    sResult = Space$(nResultLen)
    
    If GetComputerName(sResult, nResultLen) = 0 Then
      Err.Raise 9999, Description:="GetComputerName API failed."
    End If
    
    
    GetMachineName = Left$(sResult, nResultLen)

Exit Function
ERR_GetMachineName:

  GetMachineName = "<Error occured while getting login name (" & Err.Number & "):" & Err.Description

End Function

Public Function OpenAnyFile(ByVal vsFileName As String, Optional ByVal vsParameters As String = "") As Boolean
    OpenAnyFile = ShellExecute(0, "open", vsFileName, vsParameters, "", SW_SHOW) > 32
End Function

Public Function GetWinList(ByRef raWinList As Variant, Optional ByVal bNoAddIfExist As Boolean = False) As Long

    On Error GoTo Err_GetWinList
    
    Dim lWinCount As Long
    Dim sWinCaption As String
    Dim hNextWnd As Long
    Dim nRet As Integer
    
    If Not bNoAddIfExist Then
        
        Erase raWinList
        
        lWinCount = 0
        
    Else
        
        If GetDimension(raWinList) <> 0 Then
            
            lWinCount = UBound(raWinList, 2)
            
        Else
        
            lWinCount = 0
        
        End If
    
    End If
    
    hNextWnd = FindWindowEx(0&, 0&, 0&, 0&)
    
    Do Until hNextWnd = 0
    
        If IsWindowVisible(hNextWnd) Then
    
            sWinCaption = Space$(256)
            
            nRet = GetWindowText(hNextWnd, sWinCaption, Len(sWinCaption))
            
            If nRet Then
            
                sWinCaption = Trim$(Left$(sWinCaption, nRet))
                
                If sWinCaption <> "" Then
        
                    lWinCount = lWinCount + 1
                    
                    If (Not bNoAddIfExist) Or ((bNoAddIfExist) And (Not ItemExistInArray(raWinList, 2, sWinCaption))) Then
                    
                        ReDim Preserve raWinList(1 To 2, 1 To lWinCount)
                        
                        raWinList(1, lWinCount) = hNextWnd
                        
                        raWinList(2, lWinCount) = sWinCaption
                        
                    End If
                    
                End If
            
            End If
        
        End If
        
        hNextWnd = GetWindow(hNextWnd, GW_HWNDNEXT)
        
    Loop
    
    GetWinList = lWinCount
    
Exit Function
Err_GetWinList:

    ReDim raWinList(1 To 2, 1 To 1)
    
    raWinList(2, 1) = "Error occured while getting the list of windows: " & Error$
    
    Exit Function
        
End Function

Public Function ItemExistInArray(ByVal vaList As Variant, ByVal vnColumn As Integer, ByVal vvntItem As Variant, Optional ByRef rnIndexFound As Variant) As Boolean

    Dim bResult As Boolean
    Dim i As Integer
    
    bResult = False
    
    'Set default value
    If Not IsMissing(rnIndexFound) Then
        
        rnIndexFound = -1
    
    End If
    
    If IsArray(vaList) Then
    
        If GetDimension(vaList) = 2 Then
    
            If (UBound(vaList, 1) - LBound(vaList, 1) + 1) = 2 Then
            
                For i = LBound(vaList, 2) To UBound(vaList, 2)
                
                    If vaList(vnColumn, i) = vvntItem Then
                    
                        bResult = True
                        
                        If Not IsMissing(rnIndexFound) Then
                            
                            rnIndexFound = i
                        
                        End If
                        
                        
                        Exit For
                    
                    End If
                
                Next i
                
            End If
            
        End If
    
    End If
    
    ItemExistInArray = bResult

End Function

Public Function GetDimension(ByVal varr As Variant) As Integer

    GetDimension = 0
    
    On Error Resume Next
    
    GetDimension = UBound(varr, 1) - LBound(varr, 1) + 1

End Function

Public Sub AddVerInCaption(frm As Form, Optional ByVal vsFormCaption As Variant)

    Dim sCaption As String
    
    If IsMissing(vsFormCaption) Then
    
        sCaption = frm.Caption
        
    Else
    
        sCaption = vsFormCaption
    
    End If
    
    frm.Caption = sCaption & "      " & "Ver." & App.Major & "." & App.Minor & "." & App.Revision

End Sub



'
'==========================================================================================
' Routine Name : FormattedTime
' Purpose      : Gives current time in either HH:MM formate or hh:mm AM/PM format as per system settings
' Parameters   : Time value to be formatted
' Return       : String for current time in hour & min as per system settings
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 27-Jul-1998 03:22 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function FormattedTime(Optional ByVal vdtTime As Date) As String

    On Error GoTo ERR_FormattedTime

    'Routine specific local vars here
    Dim dtTimeToBeFormatted As Date
    Dim sTimeInLongFormat As String

    'Common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.FormattedTime"

    
    If IsMissing(vdtTime) Then
    
        dtTimeToBeFormatted = Now
    
    Else
    
        dtTimeToBeFormatted = vdtTime
    
    End If

    'First get the time in the system format
    sTimeInLongFormat = Trim$(Format(dtTimeToBeFormatted, "Long Time"))
    
    'Is time contains AM/PM tag? - Long Time format is according to system settings but also contains seconds
    If UCase$(Right$(sTimeInLongFormat, 1)) = "M" Then
    
        'Then format using Medium Time format - this does not include seconds
        FormattedTime = Format(dtTimeToBeFormatted, "Medium Time")
    
    Else
    
        'Else format using Short Fomat
        FormattedTime = Format(dtTimeToBeFormatted, "Short Time")
    
    End If



    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_FormattedTime:

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
' Routine Name : SelectItemInCombo
' Purpose      : Searches ItemData value in specified combo. If found, that item is selected in combo
' Parameters   : rcbo: Combo box in which passed ItemData value is to be searched
'                vlItemData: This value is searched in combo, if found then that item is selected in combo
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 27-Jul-1998 07:38 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function SelectItemInCombo(ByRef rcbo As ComboBox, ByVal vlItemData As Long, Optional ByVal bSelectNoneIfNotFound As Boolean = False) As Boolean
    
    On Error GoTo ERR_SelectItemInCombo

    'Routine specific local vars here
    Dim nItemIndex As Integer
    Dim bIsItemFound As Boolean

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.SelectItemInCombo"


    bIsItemFound = False
    
    For nItemIndex = 0 To rcbo.ListCount - 1
    
        If rcbo.ItemData(nItemIndex) = vlItemData Then
        
            bIsItemFound = True
                    
            'Select that item and exit for loop
            rcbo.ListIndex = nItemIndex
            
            Exit For
        
        End If

    Next nItemIndex


    If Not bIsItemFound Then
    
       If bSelectNoneIfNotFound Then
       
            rcbo.ListIndex = -1
       
       Else
       
            If rcbo.ListCount <> 0 Then
            
                rcbo.ListIndex = 0
            
            End If
       
       End If
    
    End If
    
    
    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    SelectItemInCombo = bSuccess

Exit Function

ERR_SelectItemInCombo:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function

Function IsNumber(ByVal vvntVal As Variant) As Boolean
    
    If Not IsEmpty(vvntVal) Then
    
        IsNumber = IsNumeric(vvntVal)
    
    Else
    
        IsNumber = False
    
    End If
    
End Function

'Function returns boolean value indicating whether item in combo is new or not
Public Function IE4LikeCombo_Change(ByRef rcbo As ComboBox, ByRef rsPrevUnselectedText As String) As Boolean
    
    Dim sInText As String
    Dim lFoundIndex As Long
    Dim bIsNewItem As Boolean
    
    
    bIsNewItem = False
        
    sInText = rcbo.Text
          
    If (StrComp(rsPrevUnselectedText, sInText, 1) <> 0) And (rcbo.SelStart = Len(sInText)) Then
    
        'Find out the ListIndex for current prefix - If not found, -1 is returned
        lFoundIndex = SendMessage(rcbo.hwnd, CB_FINDSTRING, -1, ByVal CStr(sInText))
        
        'Save the current combo text as prev unselected
        rsPrevUnselectedText = sInText

        'Now set the found index item in combo
        If rcbo.ListIndex <> lFoundIndex Then
        
            rcbo.ListIndex = lFoundIndex
            
        End If
        
        bIsNewItem = (lFoundIndex = -1)

        'Now select the rest of the part of the text
        rcbo.SelStart = Len(rsPrevUnselectedText)
        rcbo.SelLength = Len(rcbo.Text)

    End If
       
    IE4LikeCombo_Change = bIsNewItem

End Function

Function IsItemExistInCombo(ByVal vcbo As ComboBox) As Boolean

    If Len(vcbo.Text) <> 0 Then
    
        IsItemExistInCombo = (SendMessage(vcbo.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(vcbo.Text)) <> -1)
        
    Else
    
        IsItemExistInCombo = False
    
    End If
    
End Function

Function IsItemExistInList(ByVal vlst As ListBox, ByVal vsItem As String) As Boolean

    IsItemExistInList = (SendMessage(vlst.hwnd, LB_FINDSTRINGEXACT, -1, ByVal CStr(vsItem)) <> -1)
    
End Function


Private Function BrowseForFolder(ByVal hWndOwner As Long, _
            ByVal sPrompt As String) As String

Dim iNull As Integer
Dim lpIDList As Long
Dim lResult As Long
Dim sPath As String
Dim udtBI As BrowseInfo

With udtBI
    .hWndOwner = hWndOwner
    .lpszTitle = sPrompt
    .ulFlags = BIF_RETURNONLYFSDIRS
End With

lpIDList = SHBrowseForFolder(udtBI)
If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    iNull = InStr(sPath, vbNullChar)
    If iNull Then
        sPath = Left$(sPath, iNull - 1)
    End If
End If

BrowseForFolder = sPath

End Function


'
'==========================================================================================
' Routine Name : MoveListToList
' Purpose      : Moves selected or all times from one list to another
' Parameters   : vIsOnlySelectedItem - If true only selected items in src list is moved to dest item, else all items are moved
'                rlstSrc - The source list box from where items are taken
'                rlstDest - The destination list box where items are copied
'                vbRemoveFromSource - If true and item is moved to destination, that item is removed from source
'                vbAllowDuplicates - If true, it is not checked whether item already exist in destination or not
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 14-Aug-1998 03:12 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function MoveListToList(ByVal vIsOnlySelectedItem As Boolean, ByRef rlstSrc As ListBox, ByRef rlstDest As ListBox, ByVal vbRemoveFromSource As Boolean, Optional ByVal vbAllowDuplicates As Boolean = False) As Boolean

    On Error GoTo ERR_MoveListToList

    'Routine specific local vars here
    Dim nItemIndex As Integer
    Dim lItemData As Long
    Dim sItem As String
    Dim bDoMove As Boolean

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.MoveListToList"

    nItemIndex = rlstSrc.ListCount - 1
    
    While nItemIndex >= 0
        
        bDoMove = True
        
        If vIsOnlySelectedItem Then
                    
            If Not rlstSrc.Selected(nItemIndex) Then
                
                bDoMove = False
            
            End If
            
        End If
            
        
        If bDoMove Then
        
            lItemData = rlstSrc.ItemData(nItemIndex)
            
            sItem = rlstSrc.List(nItemIndex)
            
            If Not vbAllowDuplicates Then
            
                If Not IsItemExistInList(rlstDest, sItem) Then
            
                    Call rlstDest.AddItem(sItem)
                    
                    rlstDest.ItemData(rlstDest.NewIndex) = lItemData
                    
                End If
                
            Else
            
               rlstDest.AddItem (sItem)
               
               rlstDest.ItemData(rlstDest.NewIndex) = lItemData
            
            End If
            
            If vbRemoveFromSource Then
            
                Call rlstSrc.RemoveItem(nItemIndex)
            
            End If
            
        End If
        
        nItemIndex = nItemIndex - 1
    
    Wend




    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    MoveListToList = bSuccess

Exit Function

ERR_MoveListToList:

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
' Routine Name : ConcateDateTime
' Purpose      : Takes passed date and time values and returns a Date type value
' Parameters   : vvntDate - String or Date type variable containing date to be concated
'                vvntTime - String or Date type variable containing time to be concated
' Return       : Date variable containing specified date and specified time
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 17-Aug-1998 06:56 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function ConcateDateTime(ByVal vvntDate As Variant, ByVal vvntTime As Variant) As Date

    On Error GoTo ERR_ConcateDateTime

    'Routine specific local vars here

    'Common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.ConcateDateTime"



    ConcateDateTime = DateAdd("n", DateDiff("n", TimeValue("0:00"), TimeValue(vvntTime)), DateValue(vvntDate))



    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_ConcateDateTime:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


Public Function AlternateStrIfNull(ByVal vvntVariable As Variant, ByVal vsAlternateValueOnNull As String) As String

    Dim bReturnAlternate As Boolean
    
    bReturnAlternate = True
    
    If VarType(vvntVariable) = vbObject Then
        
        If Not (vvntVariable Is Nothing) Then
        
            bReturnAlternate = False
            
        Else
        
            bReturnAlternate = True
            
        End If
        
    Else
    
        If Not (IsEmpty(vvntVariable)) Then
        
            If Not (IsNull(vvntVariable)) Then
            
                If (Len(CStr(vvntVariable)) <> 0) Then
                
                    bReturnAlternate = False
                
                End If
                
            End If
            
        End If
        
    End If
    
    
    If bReturnAlternate Then
    
        AlternateStrIfNull = vsAlternateValueOnNull
        
    Else
    
        AlternateStrIfNull = CStr(vvntVariable)
    
    End If

End Function

'
'==========================================================================================
' Routine Name : GetSelectedItemDataInCombo
' Purpose      : Get the ItemData property of selected item in combo. If ListIndex is -1, it searches whether Text proeprty is really not in Combo
' Parameters   : vcbo: Combo box for which ItemData is needed
' Return       : ItemData proeprty of of item in Text property of combo
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 09-Sep-1998 03:10 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function GetSelectedItemDataInCombo(ByVal vcbo As ComboBox) As Long

    On Error GoTo ERR_GetSelectedItemDataInCombo

    'Routine specific local vars here
    Dim nSelectedItemIndex As Integer

    'Common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.GetSelectedItemDataInCombo"


    If vcbo.ListIndex = -1 Then
    
        nSelectedItemIndex = (SendMessage(vcbo.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(vcbo.Text)))
        
    Else
    
        nSelectedItemIndex = vcbo.ListIndex
        
    End If

    If nSelectedItemIndex <> -1 Then
    
        GetSelectedItemDataInCombo = vcbo.ItemData(nSelectedItemIndex)
        
    Else
    
        Err.Raise vbError + 999, "Utils.GetSelectedItemDataInCombo", "Can not find item in Text property in combo"
        
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_GetSelectedItemDataInCombo:

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
' Routine Name : GetSelectedItemInCombo
' Purpose      : Get the index of selected item in combo. If ListIndex is -1, it searches whether Text proeprty is really not in Combo
' Parameters   : vcbo: Combo box for which ItemData is needed
' Return       : index of item in Text property of combo
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 09-Sep-1998 03:23 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function GetSelectedItemInCombo(ByVal vcbo As ComboBox) As Integer

    On Error GoTo ERR_GetSelectedItemInCombo

    'Routine specific local vars here

    'Common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.GetSelectedItemInCombo"


    If vcbo.ListIndex = -1 Then
    
        GetSelectedItemInCombo = (SendMessage(vcbo.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(vcbo.Text)))
        
    Else
    
        GetSelectedItemInCombo = vcbo.ListIndex
        
    End If





    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_GetSelectedItemInCombo:

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
' Routine Name : SelectItemInCombo
' Purpose      : Selects the item specified by it's index. If that item does not exist then selects second specified item. If second item also does not exist then selects nothing
' Parameters   : vcbo: Combo box in which selection to be made
'                vnFirstIndexToTry: First index to try for selecting item
'                vbSecondIndexToTry: If item for first index does not exist then select item for this index
' Return       : Returns index of selected item
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 09-Sep-1998 03:44 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function NoFailSelectItemInCombo(ByVal vcbo As ComboBox, ByVal vnFirstIndexToTry As Integer, Optional ByVal vbSecondIndexToTry As Integer = 0, Optional ByVal vboolSelectFirstOnFail As Boolean = True) As Integer

    On Error GoTo ERR_SelectItemInCombo

    'Routine specific local vars here

    'Common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.SelectItemInCombo"


    If vnFirstIndexToTry > (vcbo.ListCount - 1) Then
    
        If vbSecondIndexToTry > (vcbo.ListCount - 1) Then
        
            If vboolSelectFirstOnFail Then
            
                If vcbo.ListCount <> 0 Then
                
                    vcbo.ListIndex = 0
                    
                Else
                
                    vcbo.ListIndex = -1
                
                End If
                
            Else
                
                vcbo.ListIndex = -1
                
            End If
            
        Else
        
            vcbo.ListIndex = vbSecondIndexToTry
        
        End If
    
    Else
    
        vcbo.ListIndex = vnFirstIndexToTry
    
    End If
    
    
    NoFailSelectItemInCombo = vcbo.ListIndex
    

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_SelectItemInCombo:

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
' Routine Name : MakeExclusiveLists
' Purpose      : This functions deletes all elements in target list box which are there in source box also,
' Parameters   : vlstSource: List box from which items are checked whether it is in target list box also.
'                rlstTarget: List box from which those items to be deleted which are also in source list box.
' Return       : Sucess or not.
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 11-Sep-1998 12:41 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function MakeExclusiveLists(ByVal vlstSource As ListBox, ByRef rlstTarget As ListBox) As Boolean

    On Error GoTo ERR_MakeExclusiveLists

    'Routine specific local vars here
    Dim nSourceListIndex As Integer
    Dim nIndexInTarget As Integer
    Dim sSourceItem As String

    'Common variables
    Dim bSuccess As Boolean                 'Return true if success
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'By default assume everything gone fine
    bSuccess = True

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.MakeExclusiveLists"


    For nSourceListIndex = 0 To vlstSource.ListCount - 1
    
        sSourceItem = vlstSource.List(nSourceListIndex)
                
        nIndexInTarget = GetItemIndexInList(rlstTarget, sSourceItem)
        
        If nIndexInTarget <> -1 Then
        
            rlstTarget.RemoveItem nIndexInTarget
        
        End If
    
    Next nSourceListIndex




    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Return success status of function
    MakeExclusiveLists = bSuccess

Exit Function

ERR_MakeExclusiveLists:

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
' Routine Name : GetItemIndexInCombo
' Purpose      : Retrieves the index of the specified item string in combo
' Parameters   : vcbo: Combo in which search to be made
'                vsItemToSearch: Item for which index is to be found in combo
'                nReturnIfNoFound: Value to be returned if item is not found in combo
' Return       : Index of the item in combo
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 11-Sep-1998 12:51 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function GetItemIndexInCombo(ByVal vcbo As ComboBox, ByVal vsItemToSearch As String, Optional ByVal vnReturnIfNoFound As Integer = -1) As Integer

    On Error GoTo ERR_GetItemIndexInCombo

    'Routine specific local vars here

    'Common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.GetItemIndexInCombo"



    GetItemIndexInCombo = SendMessage(vcbo.hwnd, CB_FINDSTRINGEXACT, vnReturnIfNoFound, ByVal CStr(vsItemToSearch))



    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_GetItemIndexInCombo:

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
' Routine Name : GetItemIndexInList
' Purpose      : Retrieves the index of the specified item string in combo
' Parameters   : vlst: Combo in which search to be made
'                vsItemToSearch: Item for which index is to be found in combo
'                nReturnIfNoFound: Value to be returned if item is not found in combo
' Return       : Index of the item in combo
' Effects      : None
' Assumes      : None
' Author       : Shital
' Date         : 11-Sep-1998 12:51 PM
' Template     : Ver.15   Author: Shital Shah   Date: 06 July, 1998
' Date          Person      Details.
'==========================================================================================
'

Public Function GetItemIndexInList(ByVal vlst As ListBox, ByVal vsItemToSearch As String, Optional ByVal vnReturnIfNoFound As Integer = -1) As Integer

    On Error GoTo ERR_GetItemIndexInList

    'Routine specific local vars here

    'Common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'Set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.GetItemIndexInList"



    GetItemIndexInList = SendMessage(vlst.hwnd, LB_FINDSTRINGEXACT, vnReturnIfNoFound, ByVal CStr(vsItemToSearch))



    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

Exit Function

ERR_GetItemIndexInList:

    'Call the global error handling routine to process the error, and check if execution should be continued
    If gobjErrors.GlobalErrorHandler(sErrorLocation) = gnRESUME_NEXT Then
        Resume Next
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

    'Explicitly exit to prevent errors due to non resumed errors
    Exit Function

End Function


Public Sub ReRaisErr()
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function IsArrayEmpty(ByVal vavntArr As Variant) As Boolean

    On Error GoTo ERR_IsArrayEmpty
    
    Dim nLowerBoundVal As Integer
    
    'This will generate the error if Variant does not containes valid array
    nLowerBoundVal = LBound(vavntArr, 1)
    
    IsArrayEmpty = False
    

Exit Function
ERR_IsArrayEmpty:

    'Do not resume next in this event handler
    Select Case Err.Number
    
        Case 13, 9          '13:Type Mismatch, 9:Subscript Out of range
            
            IsArrayEmpty = True
            
        Case Else
        
            ReRaisErr
    
    End Select
    
    Exit Function

End Function


Public Function ReturnNullIfBlank(ByVal vvntVar As Variant, Optional vboolCovertToDate As Boolean = False) As Variant

    If Not IsEmpty(vvntVar) Then
    
        If Len(CStr(vvntVar)) = 0 Then
        
            ReturnNullIfBlank = Null
            
        Else
        
            If Not vboolCovertToDate Then
            
                ReturnNullIfBlank = vvntVar
                
            Else
            
                ReturnNullIfBlank = CDate(vvntVar)
                
            End If
        
        End If
    
    Else
    
        ReturnNullIfBlank = Null
    
    End If

End Function


Public Function BoolToCheck(ByVal vboolVal As Boolean, Optional ByVal vboolGreyCheck As Boolean = False) As Integer

    If Not vboolGreyCheck Then

        If vboolVal Then
        
            BoolToCheck = vbChecked
        
        Else
        
            BoolToCheck = vbUnchecked
        
        End If
        
    Else
    
        BoolToCheck = vbGrayed
    
    End If

End Function

Public Function CheckToBool(ByVal vnCheckBoxValue As Integer, Optional ByRef rbCheckBoxGrayed As Variant) As Boolean

    If Not IsMissing(rbCheckBoxGrayed) Then
    
        rbCheckBoxGrayed = False
    
    End If

    Select Case vnCheckBoxValue
    
        Case vbChecked
        
            CheckToBool = True
        
        Case vbUnchecked
        
            CheckToBool = False
        
        Case vbGrayed
        
            If Not IsMissing(rbCheckBoxGrayed) Then
            
                rbCheckBoxGrayed = True
            
            End If
        
    
    End Select


End Function

Sub MakeAControlReadOnly(ctrlControl As Control, Optional ByVal vboolEnable As Boolean = False)

    On Error GoTo Err_MakeAControlReadOnly
    
    If TypeOf ctrlControl Is Label Then
        'Do nothing
        
    ElseIf TypeOf ctrlControl Is CommandButton Then
        ' Disable the command button
        EnableControl ctrlControl, vboolEnable
        
'    ElseIf TypeOf ctrlControl Is TDBGrid Then
'        ' Make the Grid control read only
'        Call MakeGridReadOnly(ctrlControl)
'
'    ElseIf TypeOf ctrlControl Is SSDateCombo Then
'        ' set the allow editing property to false
'        EnableControl ctrlControl, False
'
'        ' Set the Back color to the system menu bar color
'        ctrlControl.BackColor = vbMenuBar

    ElseIf TypeOf ctrlControl Is OptionButton Or TypeOf ctrlControl Is CheckBox Then
        ' enable/disable the control
        EnableControl ctrlControl, vboolEnable
        
        ' Set the Back color to the system menu bar color
        Call SetBackColorForEnable(ctrlControl, vboolEnable)

    ElseIf TypeOf ctrlControl Is ListBox Or TypeOf ctrlControl Is ComboBox Then
        ' enable/disable the control
        EnableControl ctrlControl, vboolEnable
        
        ' Set the Back color to the system menu bar color
        Call SetBackColorForEnable(ctrlControl, vboolEnable)


    Else
        
        'Control.ReadOnly = True
        ctrlControl.Locked = Not vboolEnable
        
        ' Set the Back color to the system menu bar color
        Call SetBackColorForEnable(ctrlControl, vboolEnable)
    
    End If
        
Exit Sub

Err_MakeAControlReadOnly:
    
    ' if a control does not support Readonly property then ignore such error
    If Err = gnERR_OBJECT_DOES_NOT_SUPPORT_THIS_PROPERTY Then
        'ctrlControl.Enabled = False
        Resume Next
    End If
        
End Sub

Public Sub EnableControl(ctlControl As Control, bEnableFLag As Boolean)
    ' if the control is not already Enabled/Disabled then Enable/Disable it
    If Not ctlControl.Enabled = bEnableFLag Then
        ctlControl.Enabled = bEnableFLag
    End If
End Sub

Public Sub MakeControlsReadOnly(frmForm As Form, Optional ByVal ctlOnlyInThisContainer As Variant, Optional ByVal vboolEnable As Boolean = False)
                        
    On Error GoTo Err_MakeControlReadOnly
    
    Dim ctrlCurrent As Control
        
        ' make each of the control read only
        For Each ctrlCurrent In frmForm.Controls
        
            If IsMissing(ctlOnlyInThisContainer) Then
            
                ' Make the current control readonly
                MakeAControlReadOnly ctrlCurrent, vboolEnable
                
            Else
                
                'Check if control is in required container
                If ctrlCurrent.Container Is ctlOnlyInThisContainer Then
                
                    ' Make the current control readonly
                    MakeAControlReadOnly ctrlCurrent, vboolEnable
                
                End If
                
            End If

        Next
    
    Exit Sub
    
Err_MakeControlReadOnly:
    
    ' if a control does not support Readonly property then ignore such error
    If Err = gnERR_OBJECT_DOES_NOT_SUPPORT_THIS_PROPERTY Then
        ctrlCurrent.Enabled = vboolEnable
        Resume Next
    End If
        
End Sub


Public Sub SetProperComboWidth(ByRef rcboCombo As ComboBox)

    Dim lMaxTextWidth As Long
    Dim nComboItemIndex As Integer
    Dim lTextWidth As Long
    Dim dConvFactorFromFormUnit As Double
    Dim lNull As Long
    Dim lNewDropWidth As Long
    
    lMaxTextWidth = 0
    
    
    For nComboItemIndex = 0 To rcboCombo.ListCount - 1
    
        lTextWidth = rcboCombo.Parent.TextWidth(rcboCombo.List(nComboItemIndex))
        
        If lTextWidth > lMaxTextWidth Then
        
            lMaxTextWidth = lTextWidth
        
        End If
        
    Next nComboItemIndex
    
    Select Case rcboCombo.Parent.ScaleMode
    
        Case vbTwips
            dConvFactorFromFormUnit = Screen.TwipsPerPixelX
        
        Case vbPixels
            dConvFactorFromFormUnit = 1
            
        Case Else
            MsgBox "Utils.SetProperComboWidth function not defined for non twips units in form's ScaleMode: form is " & rcboCombo.Parent.Name
            
            dConvFactorFromFormUnit = 0

    
    End Select
    
    If dConvFactorFromFormUnit <> 0 Then
    
        If lMaxTextWidth > (rcboCombo.Width - 230) Then
        
            lNull = 0
            
            '100 is to companset for scroll bar width
            lNewDropWidth = (lMaxTextWidth / dConvFactorFromFormUnit) + 50
            
            Call SendMessage(rcboCombo.hwnd, CB_SETDROPPEDWIDTH, lNewDropWidth, lNull)
        
        End If
    
    End If
    

End Sub

Public Function MinToHrMin(ByVal vlMinutes As Variant, Optional ByVal vsHrTag As String = "hr", Optional ByVal vsMinTag As String = "min", Optional ByVal vsErrOnNonNumericData As String = "? Min", Optional ByVal vlFuzzyApprox As Long = 0) As String

    Dim sReturn As String
    
    '
    'Example: vlFuzzyApprox = 5
    '

    If IsNumber(vlMinutes) Then
    
        'If minutes is between 55 to 65
        If (vlMinutes >= (60 - vlFuzzyApprox)) And (vlMinutes <= (60 + vlFuzzyApprox)) Then
            
            'Treat as 1 Hr
            sReturn = "1 " & vsHrTag
        
        'If munutes is less then 55
        ElseIf (vlMinutes < (60 - vlFuzzyApprox)) Then
        
            'Get approximation to nearest tens
            sReturn = ApproxToTens(vlMinutes, vlFuzzyApprox) & " " & vsMinTag
            
        Else
        
            Dim lHours As Long
            Dim lMinutes As Long
            
            lHours = (vlMinutes \ 60)
            lMinutes = (vlMinutes - (lHours * 60))
            
            'If remainant minutes is greater then 55
            If lMinutes >= (60 - vlFuzzyApprox) Then
                
                'Treat as whole hour
                lHours = lHours + 1
            
            End If
            
            sReturn = lHours & " " & vsHrTag
            
            
            'If remainant minutes is lessthen 5 or greater then 55
            If (lMinutes <= vlFuzzyApprox) Or (lMinutes >= (60 - vlFuzzyApprox)) Then
            
                'Don't add str for min
                
            Else
                
                'Get nearest Tens
                sReturn = sReturn & " " & ApproxToTens(lMinutes, vlFuzzyApprox) & " " & vsMinTag
                
            End If
        
        End If
    
    Else
        
        If vsErrOnNonNumericData = "? Min" Then
            
            sReturn = "? " & vsMinTag
            
        Else
        
            sReturn = vsErrOnNonNumericData
        
        End If
    
    End If
    
    MinToHrMin = sReturn

End Function

Public Function ApproxToTens(ByVal vlNumber As Long, ByVal vlFuzzyLimit As Long) As Long

    Dim lUpperTen As Long
    Dim lLowerTen As Long
    Dim lReturn As Long

    lLowerTen = (vlNumber \ 10) * 10
    
    lUpperTen = lLowerTen + 10
        
    If (vlNumber - vlFuzzyLimit) < lLowerTen Then
    
        lReturn = lLowerTen
        
        If lReturn = 0 Then
        
            lReturn = vlFuzzyLimit
        
        End If
        
    ElseIf (vlNumber + vlFuzzyLimit) > lUpperTen Then
    
        lReturn = lUpperTen
        
    Else
    
        lReturn = vlNumber
    
    End If
    
    ApproxToTens = lReturn

End Function


Public Sub SetBackColorForEnable(ctrlControl As Control, ByVal vboolEnable As Boolean)

    If Not vboolEnable Then
        
        ctrlControl.BackColor = vbMenuBar
        
    Else
    
        If TypeOf ctrlControl Is TextBox Then
            
            ctrlControl.BackColor = vbWindowBackground
            
        Else
        
            ctrlControl.BackColor = vbButtonFace
            
        End If
    
    End If

End Sub

Public Sub ClearCollection(ByRef roclColl As Collection)

    Dim i As Integer
    
    For i = roclColl.Count To 1 Step -1
    
        roclColl.Remove i
    
    Next i

End Sub

Public Function GetDatePart(ByVal vntDate As Variant) As Variant

    If IsDate(vntDate) Then
    
        GetDatePart = DateSerial(DatePart("yyyy", vntDate), DatePart("m", vntDate), DatePart("d", vntDate))
        
    Else
    
        GetDatePart = Empty
        
    End If
 
End Function

Public Function GetTimePart(ByVal vntDate As Variant) As Variant

    If IsDate(vntDate) Then
    
        GetTimePart = TimeSerial(DatePart("h", vntDate), DatePart("n", vntDate), DatePart("s", vntDate))
        
    Else
    
        GetTimePart = Empty
        
    End If
 
End Function

Public Function GetFormattedDatePart(ByVal vdtDate As Date) As String

    GetFormattedDatePart = Format$(vdtDate, "dd mmm, yyyy")

End Function

Public Function GetColItem(oclCol As Collection, ByVal vsKey As String) As Variant

    On Error Resume Next
    
    Dim vntItem As Variant
    
    'Try to get collection's item
    vntItem = oclCol.Item(vsKey)
    
    If Err.Number <> 0 Then
    
        GetColItem = Empty
    
    Else
    
        GetColItem = vntItem
    
    End If

End Function
