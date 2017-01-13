Option Explicit 
 
'Wrapper Class for simplifacation access to sap gui created by Jan Becka 09/08/2013 
 
'Last Update 13|01|17
'dependency on class clsCollection 
'dependency for loops moved to ATk functions 
 
'for excel window handling 
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpszClass As String, ByVal lpszWindow As String) As Long 
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long 
 
'system waiting function 
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 
 
Private SleepMS As Long 'this property store value of miliseconds to wait 
Private SLEEP_FAST& 'this property store value of miliseconds 
 
'Public test As SAPGUIControl 
Public session As Object 'Sap session object 
 
'SAP Table Variables 
Private strTablePath 'Save SAP table path for fuctions 
Private lng_table_visible_rows 'Count table visible rows 
Private used_tabs As clsCollection 'main collection to store used_columns collection 
Private used_columns As clsCollection 'returns column index in the table when column name is entered 
 
'SAP List Variables 
Private strListKey$ 
Private strListAddress$ 
Private lngListHeaderRow& 
Private lngListCurrentRow& 
Private used_lists As clsCollection 
Private used_list_header As clsCollection 
 
Private lngCurrentMainWindow& 
 
Private bool_AutoClose As Boolean 
 
'Issue Checking Variable 
'Private black_list As clsCollection 'to check if item is on black list 
'Public issue_log As clsCollection 'for logging of all issues call by actin InputIssues 
 
'Private cc_BlackListedRecordColumns As clsCollection 
'Private cc_BlackListedRecordItemColumns As clsCollection 
Private cc_BlackList As clsCollection 
'Private cc_IssuesCheck As clsCollection 
Public cc_IssueLog As clsCollection 
 
 
'NEW DATA ENGINE 
 
Private cc_RunParameters As clsCollection 'to store parameters same for whole run 
 
Private cc_CurrentCursors As clsCollection 
'Private cc_DataValueOfTables As clsCollection 'to store variant arrays with data 
'Private cc_DataHeaderOfTables As clsCollection 'to store header indexes for data arrays 
 
'Private cc_HeadColumnIndex As clsCollection ' 
'Private cc_ItemColumnIndex As clsCollection 
 
'Private cc_HeadRowIndex As clsCollection 
'Private cc_ItemRowIndex As clsCollection 
 
'Private cc_TableNames As clsCollection 
 
'Private cc_TableRowIndexes As clsCollection 
 
'Private bool_HeaderLoaded As Boolean 
'Private bool_FirstItem As Boolean 
 
Private str_CurrentKeyColumnName As String 
'Private str_CurrentKeyColumnValue As String 
 
'END NEW DATA ENGINE 
 
'Records Variable 
'Private colRecordDetialsIndexes As clsCollection 
'Private colItemDetailsIndexes As clsCollection 
'Private colRecords As clsCollection 'all records collection 
'Private colRecord As clsCollection 'collection containing all items 
'Private lngRecordCursor& 
'Private lngRecordItemCursor& 
 
Private lng_CurrentPhase& 
 
'Private colRecordItems As clsCollection 
 
Private lng_added_entry_row& 'row of added entry 
 
Private str_active_tab$ 'stores value of active tab 
 
'Private sht_log As Worksheet 'sheet where is log first row for header 
 
Private t_RunLog As ListObject 'item log 
 
Public Enum e_ItemLoopType 
    e_UniqueFilter 
    e_Unique 
    e_Items 
End Enum 
 
Public Enum main_button_type 'Enum for top row buttons 
    Back_3 = 3 
    Exit_15 = 15 
    cancel_12 = 12 
    Save_11 = 11 
End Enum 
 
Public Enum value_type 'Enum for top row buttons 
    saptext = 1 
    sapKey = 2 
    sapValue = 3 
    sapDisplayed = 4 
    sapCtextKey = 5 
    sapCtextDesc = 6 
    sapBool = 7 
    sapTooltip = 8 
    sapIconName = 9 
End Enum 
 
Public Enum e_PressKey 
    Enter_0 = 0 
    F1_1 = 1 
    F2_2 = 2 
    F3_3 = 3 
    F4_4 = 4 
    F5_5 = 5 
    F6_6 = 6 
    F7_7 = 7 
    F8_8 = 8 
    F9_9 = 9 
    F10_10 = 10 
    CtrlS_11 = 11 
    F12_12 = 12 
    ShiftF1_13 = 13 
    ShiftF2_14 = 14 
    ShiftF3_15 = 15 
    ShiftF4_16 = 16 
    ShiftF5_17 = 17 
    ShiftF6_18 = 18 
    ShiftF7_19 = 19 
    ShiftF8_20 = 20 
    ShiftF9_21 = 21 
    ShiftCtrl0_22 = 22 
    ShiftF11_23 = 23 
    ShiftF12_24 = 24 
    CtrlF1_25 = 25 
    CtrlF2_26 = 26 
    CtrlF3_27 = 27 
    CtrlF4_28 = 28 
    CtrlF5_29 = 29 
    CtrlF6_30 = 30 
    CtrlF7_31 = 31 
    CtrlF8_32 = 32 
    CtrlF9_33 = 33 
    CtrlF10_34 = 34 
    CtrlF11_35 = 35 
    CtrlF12_36 = 36 
    CtrlShiftF1_37 = 37 
    CtrlShiftF2_38 = 38 
    CtrlShiftF3_39 = 39 
    CtrlShiftF4_40 = 40 
    CtrlShiftF5_41 = 41 
    CtrlShiftF6_42 = 42 
    CtrlShiftF7_43 = 43 
    CtrlShiftF8_44 = 44 
    CtrlShiftF9_45 = 45 
    CtrlShiftF10_46 = 46 
    CtrlShiftF11_47 = 47 
    CtrlShiftF12_48 = 48 
    CtrlE_70 = 70 
    CtrlF_71 = 71 
    CtrlSlash_72 = 72 
    CtrlBackSlash_73 = 73 
    CtrlN_74 = 74 
    CtrlO_75 = 75 
    CtrlX_76 = 76 
    CtrlC_77 = 77 
    CtrlV_78 = 78 
    CtrlZ_79 = 79 
    CtrlPageUp_80 = 80 
    PageUp_81 = 81 
    PageDown_82 = 82 
    CtrlPageDown_83 = 83 
    CtrlG_84 = 84 
    CtrlR_85 = 85 
    CtrlP_86 = 86 
End Enum 
 
Public Enum log_flag 'Enum for top row buttons 
    Ok = 1 
    info = 2 
    warning = 3 
    ISSUE = 4 
End Enum 
 
Public Enum text_condition 'Enum for top row buttons 
    notEqual = 5 
End Enum 
 
Public Enum input_Format 
    TextFormat = 0 
    DateFormat = 1 
    DecimalRound1 = 2 
    DecimalRound2 = 3 
    DecimalRound3 = 4 
    DecimalRound4 = 5 
End Enum 
 
Public Enum e_DialogSelectionType 
    IncludeSelection = 0 
    ExcludeSelection = 1 
End Enum 
 
 
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long 
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long 
Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long 
Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long 
 
Private Type UUID 'GUID 
  Data1 As Long 
  Data2 As Integer 
  Data3 As Integer 
  Data4(7) As Byte 
End Type 
 
 
Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}" 
Const OBJID_NATIVEOM As Long = &HFFFFFFF0 
 
'Sub GetAllWorkbookWindowNames() 
Private Sub z_GetWorkbook(ByRef wbk_Ref As Workbook, Optional ByVal str_WorkbookName As String = "Worksheet in Basis (1)") 
    On Error GoTo MyErrorHandler 
 
    Dim hWndMain As Long 
    hWndMain = FindWindowEx(0&, 0&, "XLMAIN", vbNullString) 
 
    Do While hWndMain <> 0 
        z_GetWbkWindows hWndMain, wbk_Ref, str_WorkbookName 
        hWndMain = FindWindowEx(0&, hWndMain, "XLMAIN", vbNullString) 
    Loop 
 
    Exit Sub 
 
MyErrorHandler: 
    MsgBox "GetAllWorkbookWindowNames" & vbCrLf & vbCrLf & "Err = " & Err.Number & vbCrLf & "Description: " & Err.Description 
End Sub 
 
Private Sub z_GetWbkWindows(ByVal hWndMain As Long, ByRef wbk_Ref As Workbook, ByVal str_WorkbookName As String) 
    On Error GoTo MyErrorHandler 
 
    Dim hWndDesk As Long 
    hWndDesk = FindWindowEx(hWndMain, 0&, "XLDESK", vbNullString) 
 
    If hWndDesk <> 0 Then 
        Dim hWnd As Long 
        hWnd = FindWindowEx(hWndDesk, 0, vbNullString, vbNullString) 
 
        Dim strText As String 
        Dim lngRet As Long 
        Do While hWnd <> 0 
            strText = String$(100, Chr$(0)) 
            lngRet = GetClassName(hWnd, strText, 100) 
 
            If Left$(strText, lngRet) = "EXCEL7" Then 
                z_GetExcelObjectFromHwnd hWnd, wbk_Ref, str_WorkbookName 
                Exit Sub 
            End If 
 
            hWnd = FindWindowEx(hWndDesk, hWnd, vbNullString, vbNullString) 
            Loop 
 
        On Error Resume Next 
    End If 
 
    Exit Sub 
 
MyErrorHandler: 
    MsgBox "z_GetWbkWindows" & vbCrLf & vbCrLf & "Err = " & Err.Number & vbCrLf & "Description: " & Err.Description 
End Sub 
 
Public Function z_GetExcelObjectFromHwnd(ByVal hWnd As Long, ByRef wbk_Ref As Workbook, ByVal str_WorkbookName As String) As Boolean 
    On Error GoTo MyErrorHandler 
 
    Dim fOk As Boolean 
    fOk = False 
 
    Dim iid As UUID 
    Call IIDFromString(StrPtr(IID_IDispatch), iid) 
 
    Dim obj As Object 
    If AccessibleObjectFromWindow(hWnd, OBJID_NATIVEOM, iid, obj) = 0 Then 'S_OK 
        Dim objApp As Excel.Application 
        Set objApp = obj.Application 
        
        Dim wbk_Cursor As Workbook 
         
        For Each wbk_Cursor In objApp.Workbooks 
            If wbk_Cursor.Name = str_WorkbookName Then 
                Set wbk_Ref = wbk_Cursor 
                Exit For 
            End If 
        Next 
 
        fOk = True 
    End If 
 
    z_GetExcelObjectFromHwnd = fOk 
 
    Exit Function 
 
MyErrorHandler: 
    MsgBox "z_GetExcelObjectFromHwnd" & vbCrLf & vbCrLf & "Err = " & Err.Number & vbCrLf & "Description: " & Err.Description 
End Function 
 
 
Private Sub Class_Initialize() 
    SLEEP_FAST = 350 
    str_active_tab = "empty" 'to avoid empty path 
    Set used_tabs = New clsCollection 
    Set used_lists = New clsCollection 
 
    bool_AutoClose = True 
 
End Sub 
 
Private Sub Class_Terminate() 
 
    'Close Session opened by macro 
        If bool_AutoClose = True Then 
            If Not session Is Nothing Then 
                If session.Parent.Children.Count > 1 Then 'close session when its more the one session opened 
                    session.findById("wnd[0]/tbar[0]/okcd").Text = "/i" 
                    Confirm 
                End If 
            End If 
        End If 
         
    'if there were no loging during macro 
        If cc_IssueLog Is Nothing Then Exit Sub 
 
    'Log items 
        If cc_IssueLog.Count > 0 Then 
            LogToRange 
        End If 
 
End Sub 
 
Public Sub DialogVariant(ByVal str_VariantName As Variant, Optional ByVal var_VariantShortCut As Long = ShiftF5_17) 
 
    'Open Dialog for Variant 
    If TypeName(var_VariantShortCut) = "String" Then 
        ItemPress CStr(var_VariantShortCut) 
    Else 
        PressKey CLng(var_VariantShortCut) 
    End If 
     
     
    If SapObjectExist("wnd[1]/usr/txtV-LOW") Then 
     
        ItemEnter , str_VariantName, "wnd[1]/usr/txtV-LOW" 
        ItemEnter , "", "wnd[1]/usr/ctxtENVIR-LOW" 
        ItemEnter , "", "wnd[1]/usr/txtENAME-LOW" 
        ItemEnter , "", "wnd[1]/usr/txtAENAME-LOW" 
        ItemEnter , "", "wnd[1]/usr/txtMLANGU-LOW" 
     
        ItemPress "wnd[1]/tbar[0]/btn[8]" 
 
        If SapObjectExist("wnd[1]") Then 
            ConfirmPopUps 
            GoTo NOT_FOUND: 
        End If 
 
 
    Else 
     
        'press find button 
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").pressToolbarButton "&FIND" 
     
        'fill find figure 
        ItemEnter , str_VariantName, "wnd[2]/usr/txtGS_SEARCH-VALUE" 'find string 
         
        'search the item 
        ItemPress "wnd[2]/tbar[0]/btn[0]" 
         
        If Statusbar = "No Hit Found" Then 
            PressKey F12_12 
            PressKey F12_12 
            GoTo NOT_FOUND 
        End If 
         
        'close 
        PressKey F12_12 
         
        'select highlighted row 
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow 
     
        'confirm selection 
        ItemPress "wnd[1]/tbar[0]/btn[2]" 
         
    End If 
 
 
Exit Sub 
 
NOT_FOUND: 
     MsgBox str_VariantName & " variant not found" 
ISSUE: 
   
End Sub 
 
Public Function DialogFilterByList(Optional ByVal obj_List As Variant, Optional lng_SelectionType As e_DialogSelectionType = IncludeSelection, Optional bool_ClearOtherCriteria = False, Optional ByVal str_Address As String) 
 
'GO TO CONTRACT VIEW 
 
'Open throug dialog 
On Error GoTo Err: 
    session.findById(str_Address).showContextMenu 
    session.findById(Left(str_Address, 10)).selectContextMenuItemByPosition "1" 
 
 'Exclude inculde tab 
     Select Case lng_SelectionType 
         Case 0: ItemSelect ("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA") 'select tab 
         Case 1: ItemSelect ("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV") 'select tab 
     End Select 
          
 'Clean up 'if not specified list then only clear Field 
    If bool_ClearOtherCriteria Or IsMissing(obj_List) Then 
        PressKey ShiftF4_16 
    End If 
     
 'Copy to clipboard 
    Select Case TypeName(obj_List) 
        Case "String", "Long", "Integer" 
         
            Dim str_TableAddress As String 
             
            str_TableAddress = "wnd[1]/usr/tabsTAB_STRIP/_XTABLE_/ssubSCREEN_HEADER:SAPLALDB:30_XSCREENNR_/tblSAPLALDBSINGLE_XEND_" 
             
            If lng_SelectionType = ExcludeSelection Then 
                str_TableAddress = Replace(str_TableAddress, "_XTABLE_", "tabpNOSV") 
                str_TableAddress = Replace(str_TableAddress, "_XSCREENNR_", "30") 
                str_TableAddress = Replace(str_TableAddress, "_XEND_", "_E") 
            Else 
                str_TableAddress = Replace(str_TableAddress, "_XTABLE_", "tabpSIVA") 
                str_TableAddress = Replace(str_TableAddress, "_XSCREENNR_", "10") 
                str_TableAddress = Replace(str_TableAddress, "_XEND_", "") 
            End If 
 
            TableInitiate str_TableAddress 
            TableAdd Array(TableCell(, "Single value", obj_List)) 
         
        Case "ListObject" 
            obj_List.DataBodyRange.Copy 
            ItemPress "wnd[1]/tbar[0]/btn[24]" 'Paste Clipboard 
             
        Case "Range" 
            obj_List.Copy 
            ItemPress "wnd[1]/tbar[0]/btn[24]" 'Paste Clipboard 
             
     End Select 
 
 'Paste & Close 
     ItemPress "wnd[1]/tbar[0]/btn[8]" 'Close Window 
 
    Exit Function 
Err: 
    Err.Raise Err.Number 
    Resume 
     
 
End Function 
 
Public Function DialogFilterByRange(ByVal obj_ListMin As Variant, ByVal obj_ListMax As Variant, Optional lng_SelectionType As e_DialogSelectionType = IncludeSelection, Optional bool_ClearOtherCriteria = False, Optional ByVal str_Address As String) 
     
    'GO TO CONTRACT VIEW 
 
'Open throug dialog 
On Error GoTo Err: 
    session.findById(str_Address).showContextMenu 
    session.findById(Left(str_Address, 10)).selectContextMenuItemByPosition "1" 
     
 'Exclude inculde tab 
     Select Case lng_SelectionType 
         Case 0: ItemSelect ("wnd[1]/usr/tabsTAB_STRIP/tabpINTL") 'select tab 
         Case 1: ItemSelect ("wnd[1]/usr/tabsTAB_STRIP/tabpNOINT") 'select tab 
     End Select 
          
 'Clean up 
    If bool_ClearOtherCriteria Then 
        PressKey ShiftF4_16 
    End If 
          
          
 
 'Copy to clipboard 
    Select Case TypeName(obj_ListMin) 
        Case "String", "Long", "Integer" 
         
            Dim str_TableAddress As String 
             
            str_TableAddress = "wnd[1]/usr/tabsTAB_STRIP/_XTABLE_/ssubSCREEN_HEADER:SAPLALDB:30_XSCREENNR_/tblSAPLALDBINTERVAL_XEND_" 
             
            If lng_SelectionType = ExcludeSelection Then 
                str_TableAddress = Replace(str_TableAddress, "_XTABLE_", "tabpNOINT") 
                str_TableAddress = Replace(str_TableAddress, "_XSCREENNR_", "40") 
                str_TableAddress = Replace(str_TableAddress, "_XEND_", "_E") 
            Else 
                str_TableAddress = Replace(str_TableAddress, "_XTABLE_", "tabpINTL") 
                str_TableAddress = Replace(str_TableAddress, "_XSCREENNR_", "20") 
                str_TableAddress = Replace(str_TableAddress, "_XEND_", "") 
            End If 
 
         
         
            TableInitiate str_TableAddress 
            Call TableAdd(Array( _ 
                TableCell(, "Lower limit", obj_ListMin), _ 
                TableCell(, "Upper limit", obj_ListMax))) 
         
        Case "ListObject" 
            obj_ListMin.DataBodyRange.Copy 
            ItemPress "wnd[1]/tbar[0]/btn[24]" 'Paste Clipboard 
             
        Case "Range" 
            obj_ListMin.Copy 
            ItemPress "wnd[1]/tbar[0]/btn[24]" 'Paste Clipboard 
             
     End Select 
 
 'Paste & Close 
     ItemPress "wnd[1]/tbar[0]/btn[8]" 'Close Window 
 
    Exit Function 
Err: 
    Err.Raise Err.Number 
    Resume 
       
End Function 
 
Public Function SapConnectToInstance(Optional ByRef str_InstanceName$ = "ERP") As Boolean 
    Dim SapGuiAuto As Object 
    Dim SAP_Applic As Object 
    Dim Connection As Object 
    Dim WScript As Object 
     
    On Error GoTo Err: 
     
    'set default function exit status as false mean fail 
    SapConnectToInstance = False 
     
    If Environ("USERNAME") = "beckajan" And str_InstanceName = "ERP" Then str_InstanceName = "ERA eERP R/3 - Acc." 
 
    'Establish SAP connection 
    On Error Resume Next 
       Set SapGuiAuto = GetObject("SAPGUI") 
    On Error GoTo Err: 
    If SapGuiAuto Is Nothing Then 
        MsgBox "please start SAPLogon", vbExclamation 
        Exit Function 
    End If 
    On Error Resume Next 
    Set SAP_Applic = SapGuiAuto.GetScriptingEngine 
    On Error GoTo Err: 
    If SAP_Applic Is Nothing Then 
        MsgBox "scripting disabled", vbCritical 
        Exit Function 
    End If 
 
     
    'Find ERP or ERA Connection 
    If Connection Is Nothing Then 
         
        Dim lng_ConnectionCount& 
        Dim lng_ConnectionCursor& 
        Dim bool_ConnectionFound As Boolean 
         
        lng_ConnectionCount = CLng(SAP_Applic.Children.Count) 
        lng_ConnectionCursor = 0 
         
        'Loop throng connections to found if there is right function 
        Do Until CBool(lng_ConnectionCursor > (lng_ConnectionCount - 1)) 
            If CBool(SAP_Applic.Children(CLng(lng_ConnectionCursor)).Description Like "*" & str_InstanceName & "*") Then 
                bool_ConnectionFound = True 
                Exit Do 
            End If 
            lng_ConnectionCursor = lng_ConnectionCursor + 1 
        Loop 
         
        'Connection Found 
        If bool_ConnectionFound Then 
            Set Connection = SAP_Applic.Children(CLng(lng_ConnectionCursor)) 
        Else 
            MsgBox "no " & str_InstanceName & " instances found ", vbExclamation 
            Exit Function 
        End If 
         
    End If 
 
    'Check Empty if there is not Empty Session 
    If session Is Nothing Then 
         
        Dim lng_SessionsCount& 
        Dim lng_SessionCursor& 
        Dim bool_EmptySessionFound& 
         
        lng_SessionsCount = CLng(Connection.Children.Count) 
        lng_SessionCursor = 0 
              
        'find if any empty session is not open 
        Do Until CBool(lng_SessionCursor > (lng_SessionsCount - 1)) 
             
            If CBool(Connection.Children(CLng(lng_SessionCursor)).findById("wnd[0]").Text Like "SAP Easy Access*") Then 
                bool_EmptySessionFound = True 
                Exit Do 
            End If 
            lng_SessionCursor = lng_SessionCursor + 1 
        Loop  'Then Exit For 'empty session found leave stop looking for next one 
 
        If bool_EmptySessionFound Then 
            Set session = Connection.Children(CLng(lng_SessionCursor))  'set this session as default 
        Else 
            Connection.Children(CLng(lng_SessionsCount) - 1).createsession 'Create new session for macro purpose 
     
            'wait until new window will appear 
            Do 
                DoEvents: Sleep 1000 
            Loop While lng_SessionsCount = CLng(Connection.Children.Count) 
 
            'set this window as new 
            Set session = Connection.Children(CLng(Connection.Children.Count) - 1) 'set this session as default 
            DoEvents: Sleep 300 
             
        End If 
    End If 
     
    session.findById("wnd[0]").maximize 
     
    lngCurrentMainWindow = 0 
     
    SapConnectToInstance = True 
     
Exit Function 
 
Err: 
 
    If Err.Number = 614 Then 
        MsgBox "Your SAP Application Don't have enabled SAPGUI Macros please request this via IT Self-Service Portal" 
        Exit Function 
    End If 
     
    If Err.Number > 0 Then Err.Raise Err.Number 
     
End Function 
'Here Starts Collection Properties 
Public Sub SapSetSleepTimeInMS(ByRef time As Long) 
    SLEEP_FAST = time 
End Sub 
 
Public Sub PressKey(lng_Keys As e_PressKey, Optional ByVal lng_SendToWindow) 
 
    If IsMissing(lng_SendToWindow) Then 
        lng_SendToWindow = lngCurrentMainWindow 
    Else 
        Select Case TypeName(lng_SendToWindow) 
            Case "String": lng_SendToWindow = CLng(Mid(lng_SendToWindow, 4, 1)) 
            Case Else 
        End Select 
    End If 
 
    Call session.findById("wnd[" & lng_SendToWindow & "]").sendVKey(lng_Keys) 
 
End Sub 
 
Public Sub Confirm(Optional lng_repeated& = 1) 
    Dim lng_Index& 'index for repeating 
 
    For lng_Index = 1 To lng_repeated 
        Sleep (SLEEP_FAST) 
        session.findById("wnd[0]").sendVKey 0 
    Next 
     
End Sub 
 
Public Sub SapDontAutoClose() 
    bool_AutoClose = False 
End Sub 
 
Private Sub z_SetMainWindowLevel(ByRef strAddress$) 
 
    lngCurrentMainWindow = CLng(Mid(strAddress, InStr(strAddress, "[") + 1, (InStr(strAddress, "]") - (InStr(strAddress, "[") + 1)))) 
 
End Sub 
 
Public Sub Wait(ByRef lngMiliseconds&) 
    Sleep (lngMiliseconds) 
End Sub 
 
Public Sub ConfirmWithPopUps(Optional ByRef lngKeyCode& = 0, Optional ByVal str_ButtonCaption As String = "") 
    ConfirmPopUps CLng(lngKeyCode), True, str_ButtonCaption 
End Sub 
 
Public Sub ConfirmAll(Optional ByRef lngKeyCode& = 0, Optional ByRef boolConfirmBefore As Boolean = False, Optional ByVal str_ButtonCaption As String = "") 
     
    Confirm 
 
    Dim str_LastWarning As String 
 
    Do While Statusbar <> "" 
        If str_LastWarning = Statusbar And str_LastWarning <> "" Then Exit Sub 
        Call LogAdd(str_CurrentKeyColumnName, , Statusbar, ISSUE) 
 
        Confirm 
        str_LastWarning = Statusbar 
    Loop 
 
    ConfirmPopUps CLng(lngKeyCode), boolConfirmBefore, str_ButtonCaption 
 
    'statusabar confirmations 
 
    'Dim str_LastWarning As String 
 
    Do While Statusbar <> "" 
        If str_LastWarning = Statusbar And str_LastWarning <> "" Then Exit Sub 
        Call LogAdd(str_CurrentKeyColumnName, , Statusbar, ISSUE) 
        Confirm 
    Loop 
     
End Sub 
 
Public Sub ConfirmPopUps(Optional ByRef lngKeyCode& = 0, Optional ByRef boolConfirmBefore As Boolean = False, Optional ByVal str_ButtonCaption As String = "") 
 
    'default is 0 means enter 
 
    Dim boolPopUpNotShown As Boolean: boolPopUpNotShown = False 'set boolean which defines if error were appeared 
 
    If Not SapObjectExist("wnd[" & lngCurrentMainWindow + 1 & "]") Then Exit Sub 
 
    If str_ButtonCaption <> "" Then 
        Dim obj_Cursor As Object 
     
        For Each obj_Cursor In session.findById("wnd[" & lngCurrentMainWindow + 1 & "]").Children(1).Children 
            If obj_Cursor.Type = "GuiButton" Then 
                If obj_Cursor.Text = str_ButtonCaption Then 
                    ItemPress (obj_Cursor.ID) 
                    Exit Sub 
                End If 
            End If 
        Next 
    End If 
 
    'confirm just when its neccesary 
    If boolConfirmBefore Then Confirm 'call confirm action 
     
    'for each entry row check need to confirm popup confirmation 
    On Error GoTo POPUP_END 
    Do Until boolPopUpNotShown 
        session.findById("wnd[" & lngCurrentMainWindow + 1 & "]").sendVKey lngKeyCode 'if exit not e 
    Loop 
     
    Exit Sub 
 
POPUP_END: 
    If Err.Number = 619 Then boolPopUpNotShown = True: On Error GoTo 0: Resume Next 
     
End Sub 
 
Private Sub z_TableCheckPage(ByRef lng_Row&, Optional ByRef str_temp_table_path, Optional ByRef lng_temp_table_visible_rows) 
    
    Dim lng_table_scrollbar& 'save value current of scrollbar position 
    If IsMissing(lng_temp_table_visible_rows) Then lng_temp_table_visible_rows = lng_table_visible_rows 
    If IsMissing(str_temp_table_path) Then str_temp_table_path = strTablePath 
    
    With session 
 
        'load scrollbar first row position 
        lng_table_scrollbar = .findById(str_temp_table_path).verticalScrollbar.Position 
           
        'check if absolute row is inside scroll bar view 
        If (lng_Row >= lng_table_scrollbar) And (lng_Row < (lng_table_scrollbar + lng_temp_table_visible_rows)) Then 
             lng_Row = lng_Row - lng_table_scrollbar 'calculate to relative to view 
        Else 
            .findById(str_temp_table_path).verticalScrollbar.Position = lng_Row 
            ConfirmWithPopUps 
             lng_Row = 0 'set row to 0 
        End If 
 
    End With 
End Sub 
 
 
Public Function SapObjectExist(ByRef strAddress$) As Boolean 
    'Simple check if sap item exist 
 
    Dim objTest As Object 
 
    On Error Resume Next 
        Set objTest = session.findById(strAddress) 
        SapObjectExist = Not CBool(objTest Is Nothing) 
    On Error GoTo 0 
 
    Set objTest = Nothing 
 
End Function 
 
Public Sub TableInitiate(ByRef strTablePathEntry$) 
 
    Dim lng_admin_table_rows& 'rows admin setting table 
    Dim lng_admin_table_visible_rows& 'rows admin setting table 
    Dim lng_admin_table_cursor& 'current admin table cursor 
    Dim lng_admin_table_page& 'admin table page 
    Dim lng_admin_cursor_item_column& 'stores temporary item under cursors row 
    Dim lng_admin_table_columns_count& 'save value of column in table 
     
    Dim arr_settings_table() 'array which is storing order of items 
    Dim lng_settings_table_cursor& 'cursor for loop in settings table 
    Dim lng_visible_column_counter& 'counts only visible columns 
     
    Dim str_under_cursor_name 'name of field under cursor 
 
'check tab 
    z_HandleTab strTablePathEntry 
     
'reset entry rows everytime table initiate is called 
    lng_added_entry_row = -1 
 
'set window main level 
    z_SetMainWindowLevel strTablePathEntry 
 
'check if its aleready initiated same table skip process 
  If strTablePath = strTablePathEntry Then Exit Sub 
   
'save table path to temporary 
    strTablePath = strTablePathEntry 'set public table path 
     
'count rows to recieve last row 
    lng_table_visible_rows = session.findById(strTablePathEntry).visibleRowCount 
 
'check if collection is not exist already 
    If used_tabs.KeyExists(strTablePath) Then 
        Set used_columns = used_tabs.Item(strTablePath) 
    Else 
 
        'poplate name collection 
        Set used_columns = New clsCollection 
 
          
    On Error GoTo Errors: 
        With session 
            lng_admin_table_columns_count = .findById(strTablePathEntry).Columns.Count 
            lng_admin_table_cursor = -1 
             
            Do 
                lng_admin_table_cursor = lng_admin_table_cursor + 1 
                 
                If Not (used_columns.KeyExists(.findById(strTablePathEntry).Columns(lng_admin_table_cursor).Title)) Then 
                    'in some cases colums have same names and descriptions then it takes just first one 
                    'there is no possibility how to recognize which one is which one means they should be unimportatnt 
                    used_columns.Add _ 
                                    lng_admin_table_cursor, _ 
                                    Trim(CStr(.findById(strTablePathEntry).Columns(lng_admin_table_cursor).Title)) 
                 
                    'Debug.Print Trim(CStr(.findById(strTablePathEntry).Columns(lng_admin_table_cursor).Title)) 
                Else 
                    'Debug.Print .findById(strTablePathEntry).Columns(lng_admin_table_cursor).tooltip & "-" & .findById(strTablePathEntry).Columns(lng_admin_table_cursor).Title 
                End If 
 
            Loop Until lng_admin_table_columns_count = lng_admin_table_cursor 
             
        End With 
LoopEnd: 
     
    On Error GoTo 0 
 
        used_tabs.Add used_columns, strTablePath 'store collection if there will be used in next time 
           
    End If 
Exit Sub 
Errors: 
Select Case Err.Number 
    Case 613: Resume LoopEnd 
End Select 
    
End Sub 
 
Public Function TableRowsCount() As Long 
'count number of rows if parameters are entered 
    TableRowsCount = lng_table_visible_rows 
End Function 
 
Public Sub TableEnterLine(ByRef lngRowIndex&, ByRef strColumnName$, ByRef anyValue) 
    session.findById(strTablePath).getcell(lngRowIndex, used_columns.Item(strColumnName)).Text = anyValue 
End Sub 
 
Public Sub TableEnter(ByRef lng_action As value_type, ByVal lngRowIndex&, ByRef strColumnName$, ByRef anyValue) 
    z_TableCheckPage lngRowIndex  'check if row is currently vissible in the table 
     
    If lngRowIndex = -1 Then lng_added_entry_row = 0 'if row is -1 means then there are no rows yet 
     
    With session.findById(strTablePath).getcell(lngRowIndex, used_columns.Item(strColumnName)) 
     
        Select Case lng_action 
            Case 1:  .Text = anyValue 
            Case 2:  .Key = anyValue 
            Case 3:  .Value = anyValue 
            Case 7:  .Selected = CBool(anyValue) 
            Case Else 
        End Select 
    End With 
     
    DoEvents 
    Wait SleepMS 
End Sub 
 
Public Function TableRead(ByRef lng_action As value_type, ByVal lngRowIndex&, ByRef strColumnName$) As String 
    z_TableCheckPage lngRowIndex  'check if row is currently vissible in the table 
   
    With session.findById(strTablePath).getcell(lngRowIndex, used_columns.Item(strColumnName)) 
         
        Select Case lng_action 
             Case 1: TableRead = .Text 
             Case 2: TableRead = .Key 
             Case 3: TableRead = .Value 
             Case 4: TableRead = .DisplayedText 
             Case 7: TableRead = .Selected 
             Case 8: TableRead = .Tooltip 
             Case 9: TableRead = .IconName 
             Case Else 
        End Select 
    End With 
     
End Function 
 
Public Sub TableItemEdit(ByVal lngRowIndex&, ByVal strColumnName$) 
    z_TableCheckPage lngRowIndex  'check if row is currently vissible in the table 
     
    Me.session.findById(strTablePath).getcell(lngRowIndex, used_columns.Item(strColumnName)).SetFocus 
    Me.session.findById("wnd[" & lngCurrentMainWindow & "]").sendVKey 2 'double click ekvivalent 
     
End Sub 
 
Public Function TableUpdateLine(ByRef lngLine&, ParamArray rowItems()) As Long 
     
    Dim lngCursor& 
     
    For lngCursor = LBound(rowItems) To UBound(rowItems) Step 3 
     
        TableEnter CLng(rowItems(lngCursor)), lngLine, CStr(rowItems(lngCursor + 1)), CStr(rowItems(lngCursor + 2)) 
     
    Next 
 
End Function 
 
Public Function TableFind(ByRef lng_action As value_type, ByRef strColumnName$, Optional ByRef anyValue, Optional ByVal lng_table_absolute_cursor& = -1) As Long 
 
On Error GoTo Err 
 
    Dim str_empty_symbol$ 'setup empty cell symbol 
    Dim str_cell_value$ ' contains value of the cell 
   ' Dim lng_table_absolute_cursor& 'absolute postion in table 
    Dim lngViewMax& 'if vertical scrollbar maximum 0 then switch to page view 
    Dim lngActualScrollBarPostion 'temporary value to store current scrollbar position 
    Dim bool_Found As Boolean 'if value was found 
      
    Dim lng_WatchDog& 
      
    'scroll from start moved to entry parameters 
    'lng_table_absolute_cursor = -1 
    bool_Found = False 
     
    If lng_action = sapKey Then str_empty_symbol = " " Else str_empty_symbol = "" 'key has differnt empty value 
    If IsMissing(anyValue) Then anyValue = str_empty_symbol 
     
     
    With session.findById(strTablePath).verticalScrollbar 
     
        'define last view 
        If .Maximum = 0 Then 'check if there are more pages 
            lngViewMax = .pagesize 
        Else 
            lngViewMax = .Maximum 
        End If 
         
        lngActualScrollBarPostion = .Position 
         
    End With 
     
    'reset positon 
    session.findById(strTablePath).verticalScrollbar.Position = 0 'reset position 
    ConfirmPopUps 
    On Error Resume Next 
     session.findById(strTablePath).verticalScrollbar.Position = 0 
    On Error GoTo Err 
     
     
    'looping through the view to find empty cell 
    On Error GoTo Err_LastItem 
    Do 
        lng_table_absolute_cursor = lng_table_absolute_cursor + 1 
       ' Debug.Print lng_table_absolute_cursor 
        str_cell_value = TableRead(lng_action, lng_table_absolute_cursor, strColumnName) 
    Loop While str_cell_value <> anyValue And lng_table_absolute_cursor < lngViewMax 
 
Err_LastItem: 
     
    If Err.Number = 613 Then 'check if only one record 
        lng_table_absolute_cursor = lng_table_absolute_cursor - 1 
        On Error GoTo Err 
        Resume Err_LastItem 
    End If 
 
 
    'check if loop ends with no data or record was on last row 
    If anyValue = TableRead(lng_action, lng_table_absolute_cursor, strColumnName) Then 
        TableFind = lng_table_absolute_cursor 
    Else 
        TableFind = -1 
    End If 
 
    'check if scroll bar position is not over view maximum 
    If lng_table_absolute_cursor > lngViewMax Then 
        session.findById(strTablePath).verticalScrollbar.Position = lngViewMax 
    Else 
        session.findById(strTablePath).verticalScrollbar.Position = lng_table_absolute_cursor 
    End If 
 
 
Err: 
 
    Debug.Print Err.Number; Err.Description 
     
End Function 
 
Private Function z_FormatInput(ByRef str_Input$, ByRef lng_FormatCheck As input_Format) As String 
 
    Select Case lng_FormatCheck 
        Case input_Format.TextFormat 
            z_FormatInput = str_Input 
        Case input_Format.DateFormat 
            z_FormatInput = Format(str_Input, "ddmmyyyy") 
        Case input_Format.DecimalRound1 
            z_FormatInput = Replace(Replace(CStr(Round(CDbl(str_Input), 1)), Application.International(xlThousandsSeparator), ""), Application.International(xlDecimalSeparator), ",") 
        Case input_Format.DecimalRound2 
            z_FormatInput = Replace(Replace(CStr(Round(CDbl(str_Input), 2)), Application.International(xlThousandsSeparator), ""), Application.International(xlDecimalSeparator), ",") 
        Case input_Format.DecimalRound3 
            z_FormatInput = Replace(Replace(CStr(Round(CDbl(str_Input), 3)), Application.International(xlThousandsSeparator), ""), Application.International(xlDecimalSeparator), ",") 
        Case input_Format.DecimalRound4 
            z_FormatInput = Replace(Replace(CStr(Round(CDbl(str_Input), 4)), Application.International(xlThousandsSeparator), ""), Application.International(xlDecimalSeparator), ",") 
    End Select 
 
End Function 
 
Public Function TableCell(Optional ByVal lng_Type As value_type = saptext, Optional ByVal str_Column As String, Optional ByVal str_Value As String) As Variant 
 
    Dim arr_Temp(0 To 2) 
 
    arr_Temp(0) = lng_Type 
    arr_Temp(1) = str_Column 
    arr_Temp(2) = str_Value 
     
    TableCell = arr_Temp 
 
End Function 
 
Public Function TableAdd(ByVal value_input As Variant) As Long 
    
    Dim lng_column_cursor& 
    
 
    lng_column_cursor = LBound(value_input) 
'QUICK ADD CHECK 
    'to save search performance search for new row just with the first check 
    'first conditon just if its first time data are added 
    'second condition checking every time when is cursor on last of view rows to repeat search procedure table find 
    If lng_added_entry_row <> -1 And TableLastRecordInPageView = False Then 
        lng_added_entry_row = lng_added_entry_row + 1 'Add new line 
        GoTo NEW_LINE_INPUT 'skip reset lookup and go strait to adding item 
    End If 
     
On Error GoTo Err: 
    session.findById(strTablePath).SetFocus 
     
    Err.Raise 617 'raise error exception skip send key action 
     
    session.findById("wnd[0]").sendVKey 20 'send to window Shift+F8 means insert new record 
    lng_added_entry_row = session.findById(strTablePath).currentrow + session.findById(strTablePath).verticalScrollbar.Position 'returns relative row 
     
NEW_LINE_INPUT: 
     
'Check line of item 
    If lng_added_entry_row = -1 Then 
        lng_added_entry_row = session.findById(strTablePath).currentrow + session.findById(strTablePath).verticalScrollbar.Position 
    End If 
 
'Check if first input item is clear 
     If TableRead(CLng(value_input(lng_column_cursor)(0)), lng_added_entry_row, CStr(value_input(lng_column_cursor)(1))) <> "" Then 
        lng_added_entry_row = lng_added_entry_row + 1 
    End If 
 
    If TableRead(CLng(value_input(lng_column_cursor)(0)), lng_added_entry_row, CStr(value_input(lng_column_cursor)(1))) <> "" Then 
        lng_added_entry_row = lng_added_entry_row + 1 
        On Error GoTo Err 
        Err.Raise 617 
    End If 
 
'Enter new Items 
    For lng_column_cursor = LBound(value_input) To UBound(value_input) 
        TableEnter CLng(value_input(lng_column_cursor)(0)), lng_added_entry_row, CStr(value_input(lng_column_cursor)(1)), CStr(value_input(lng_column_cursor)(2)) 'add item to emtpy row 
    Next 
     
'Return Table Row 
    TableAdd = lng_added_entry_row 'return row of new entry 
     
    Exit Function 
 
Err: 
    Select Case Err.Number 
        Case 617, 613 'virtual key (Shift+F8) is not enabled 
            'TableFind function without last third value argument means find first empty row 
            lng_added_entry_row = TableFind(CLng(value_input(lng_column_cursor)(0)), CStr(value_input(lng_column_cursor)(1))) 
            Resume NEW_LINE_INPUT 
        Case Is > 0 
           Debug.Print Err.Description; " "; Err.Number: Debug.Assert Err.Number = 0: Resume 
    End Select 
     
End Function 
 
Public Function TableLastRecordInPageView() As Boolean 
    'Debug.Print session.findById(strTablePath) 
      TableLastRecordInPageView = ((lng_added_entry_row + 1) = (session.findById(strTablePath).verticalScrollbar.Position + session.findById(strTablePath).verticalScrollbar.pagesize)) 
End Function 
 
Public Sub TableRecordRemoved() 
    lng_added_entry_row = lng_added_entry_row + 1 
End Sub 
 
Public Sub TableScrollToRow(ByRef lngScrollRows&) 
    session.findById(strPath).verticalScrollbar.Position lngScrollRows 
End Sub 
Public Sub TableSelect(Optional ByVal lngRowIndex, Optional ByVal lngColumnIndexOrName) 
     
     
    If Not IsMissing(lngRowIndex) Then 
        z_TableCheckPage CLng(lngRowIndex)  'check if row is currently vissible in the table 
    End If 
     
    If lngRowIndex = -1 Then lng_added_entry_row = 0 'if row is -1 means then there are no rows yet 
     
    'SELECT column 
    If IsMissing(lngRowIndex) And Not IsMissing(lngColumnIndexOrName) Then 
         
    End If 
     
    'SELECT Row 
    If Not IsMissing(lngRowIndex) And IsMissing(lngColumnIndexOrName) Then 
        session.findById(strTablePath).getAbsoluteRow(lngRowIndex).Selected = True 
    End If 
     
    'SELECT Cell 
    If Not IsMissing(lngRowIndex) And Not IsMissing(lngColumnIndexOrName) Then 
        If IsNumeric(lngColumnIndexOrName) Then 
            session.findById(strTablePath).getcell(lngRowIndex, CLng(lngColumnIndexOrName)).SetFocus 
        Else 
            session.findById(strTablePath).getcell(lngRowIndex, used_columns.Item(lngColumnIndexOrName)).SetFocus 
        End If 
    End If 
     
    DoEvents 
    Wait SleepMS 
End Sub 
Public Sub ListInitiate(ByRef strListAddressEntry$, ByRef strListKeyEntry$, ByRef arrHeaderNames) 
 
'check tab 
    z_HandleTab strListAddressEntry 
   
'save table path to temporary 
    strListKey = strListKeyEntry 'set public table path 
    strListAddress = strListAddressEntry 
 
'check if collection is not exist already 
    If used_lists.KeyExists(strListKeyEntry) Then 
        Set used_list_header = used_lists.Item(strListKeyEntry) 
         
        lngListCurrentRow = 1 
        lngListHeaderRow = used_list_header.ReadDetail(1) 
        strListAddress = used_list_header.ReadDetail(0) 
    Else 
     
    'poplate name collection with header names 
        Dim colTempHeaderList As clsCollection 
         
        Set colTempHeaderList = New clsCollection 
        colTempHeaderList.FillFromArray arrHeaderNames 
 
    'create new list header collection 
        Set used_list_header = New clsCollection 
         
        With session 
         
            Dim lngObjCursor& 
            Dim lngListHeaderColumn& 
             
            lngListHeaderRow = -1 'initiate empty header row 
             
            With .findById(strListAddressEntry) 
             
                For lngObjCursor = 0 To .Children.Count 
                 
                    With .Children(CLng(lngObjCursor)) 
                 
                        If colTempHeaderList.KeyExists(CStr(.Text)) Then 
         
                        Debug.Print CStr(.Text) 
         
                        'get coordinations from 
                            z_GetCoordsFromAddress lngListHeaderColumn, lngListHeaderRow, CStr(.ID) 
                         
                        'add item to tabel header collection 
                            used_list_header.Add lngListHeaderColumn, CStr(.Text) 
                             
                        'check if all headers are found 
                            colTempHeaderList.Remove (.Text) 
                            If colTempHeaderList.Count = 0 Then GoTo AllHeadersFound 
                             
                        End If 
                     
                    End With 
     
                Next 
             
            End With 
             
            If colTempHeaderList.Count <> 0 Then MsgBox "Not All Header Names Found": Exit Sub 
             
AllHeadersFound: 
        End With 
 
        lngListCurrentRow = 1 'setup first row of the list 
        used_list_header.AddDetails strListAddressEntry, lngListHeaderRow, strListKey 'save details about current list 
        used_lists.Add used_list_header, strListKey 'store collection if there will be used in next time 
           
    End If 
 
End Sub 
 
Private Sub z_GetCoordsFromAddress(ByRef lngA, ByRef lngB, ByVal strAdress$) 
     
    Dim arrCoords() As String 
     
    strAdress = Mid(strAdress, InStrRev(strAdress, "[") + 1) 
    arrCoords = Split(strAdress, ",") 
     
    lngA = CLng(arrCoords(0)) 
    lngB = CLng(Mid(arrCoords(1), 1, Len(arrCoords(1)) - 1)) 
 
End Sub 
 
Public Sub ListRecordFocus(ByRef strHeaderName$, Optional ByRef lngRow) 
     
    If Right(strListAddress, 1) <> "/" Then strListAddress = strListAddress & "/" 
     
    If IsMissing(lngRow) Then 
        Call session.findById(strListAddress & "lbl[" & used_list_header.Item(strHeaderName) & "," & lngListCurrentRow + (lngListHeaderRow + 1) - session.findById(strListAddress).verticalScrollbar.Position & "]").SetFocus 
    Else 
        If lngRow = "header" Then '-1 means header 
           Call session.findById(strListAddress & "lbl[" & used_list_header.Item(strHeaderName) & "," & lngListHeaderRow & "]").SetFocus 
        Else 
        End If 
    End If 
End Sub 
 
 
Public Sub ListRecordSelect(ByRef strHeaderName$, Optional ByRef lngRow) 
     
    If Right(strListAddress, 1) <> "/" Then strListAddress = strListAddress & "/" 
     
    If IsMissing(lngRow) Then 
        EditItem strListAddress & "lbl[" & used_list_header.Item(strHeaderName) & "," & lngListCurrentRow + (lngListHeaderRow + 1) - session.findById(strListAddress).verticalScrollbar.Position & "]" 
    Else 
        If lngRow = "header" Then '-1 means header 
            EditItem strListAddress & "lbl[" & used_list_header.Item(strHeaderName) & "," & lngListHeaderRow & "]" 
        Else 
        End If 
    End If 
End Sub 
 
 
Public Function ListRecordCheck(ByRef strHeaderName$, Optional ByRef lngRow) As String 
         
    If Right(strListAddress, 1) <> "/" Then strListAddress = strListAddress & "/" 
         
    If IsMissing(lngRow) Then 
        Call Enter(sapBool, True, strListAddress & "chk[" & used_list_header.Item(strHeaderName) & "," & lngListCurrentRow + (lngListHeaderRow + 1) - session.findById(strListAddress).verticalScrollbar.Position & "]") 
    Else 
        Call Enter(sapBool, True, strListAddress & "chk[" & used_list_header.Item(strHeaderName) & "," & lngRow & "]") 
    End If 
 
End Function 
 
Public Function ListRecordRead(ByRef strHeaderName$, Optional ByRef lngRow) As String 
     
    If Right(strListAddress, 1) <> "/" Then strListAddress = strListAddress & "/" 
      
    If IsMissing(lngRow) Then 
        ListRecordRead = Read(saptext, strListAddress & "lbl[" & used_list_header.Item(strHeaderName) & "," & lngListCurrentRow + (lngListHeaderRow + 1) - session.findById(strListAddress).verticalScrollbar.Position & "]") 
    Else 
        ListRecordRead = Read(saptext, strListAddress & "lbl[" & used_list_header.Item(strHeaderName) & "," & lngRow & "]") 
    End If 
 
End Function 
Public Sub ListInitiateAdvanced(ByRef strListAddressEntry$, ByRef strListKeyEntry$, ByRef arr_HeaderNames, ByRef lng_HeaderRowInput&, ByRef arr_HeaderIndexes) 
 
'check tab 
    z_HandleTab strListAddressEntry 
       
'save table path to temporary 
    strListKey = strListKeyEntry 'set public table path 
    strListAddress = strListAddressEntry 
'count rows to recieve last row 
 
'check if collection is not exist already 
    If used_lists.KeyExists(strListKeyEntry) Then 
        Set used_list_header = used_lists.Item(strListKeyEntry) 
         
        lngListCurrentRow = 1 
        lngListHeaderRow = used_list_header.ReadDetail(1) 
        strListAddress = used_list_header.ReadDetail(0) 
    Else 
     
    'create new list header collection 
        Set used_list_header = New clsCollection 
         
        With session 
         
            Dim lngObjCursor& 
            Dim lngListHeaderColumn& 
             
            lngListHeaderRow = lng_HeaderRowInput 'initiate empty header row 
             
            Dim obj_List As Object 
            Set obj_List = .findById(strListAddressEntry) 
             
            Dim i_Cursor& 
             
            For i_Cursor = 0 To UBound(arr_HeaderNames) 
                used_list_header.Add arr_HeaderIndexes(i_Cursor), arr_HeaderNames(i_Cursor) 
            Next 
             
             
AllHeadersFound: 
        End With 
 
        lngListCurrentRow = 1 'setup first row of the list 
        used_list_header.AddDetails strListAddressEntry, lngListHeaderRow, strListKey 'save details about current list 
        used_lists.Add used_list_header, strListKey 'store collection if there will be used in next time 
           
    End If 
 
End Sub 
 
Public Function ListRecordNext(Optional lngCursorPosition) As Boolean 
 
    Dim lngEoFWatchDog& 
    Dim lngTempCursorPosition& 
 
    'if cursor is not defined take previous 
    If IsMissing(lngCursorPosition) Then 
        lngTempCursorPosition = lngListCurrentRow 
    Else 
        lngTempCursorPosition = lngCursorPosition 
    End If 
  
    ListRecordNext = True 
  
NextRow: 
  
    If Not z_ScrollbarAction(CLng(lngTempCursorPosition), strListAddress) Then 
        ListRecordNext = False 
        Exit Function 
    Else 
     
On Error GoTo DummyLine 
        If vbString = VarType(session.findById(strListAddress & "lbl[" & used_list_header.Item(1) & "," & lngTempCursorPosition + 1 + (lngListHeaderRow + 1) - session.findById(strListAddress).verticalScrollbar.Position & "]").ID) Then 
            lngListCurrentRow = CLng(lngTempCursorPosition) + 1 
            lngEoFWatchDog = 0 'reset watch dog 
        End If 
    End If 
         
    'increase row 
    
DummyLine: 
    Select Case Err.Number 
        Case 619 
            If lngEoFWatchDog = 10 Then 'if there error is 10 lines in row means that we are on end of file 
                ListRecordNext = False 
                Exit Function 
            End If 
            lngEoFWatchDog = lngEoFWatchDog + 1 
            lngTempCursorPosition = lngTempCursorPosition + 1: Resume NextRow 
        Case 0: 'no error 
        Case Else: On Error GoTo 0: Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext: Resume 
    End Select 
 
End Function 
 
Private Function z_ScrollbarAction(ByVal lng_RowCursorPosition As Long, ByRef strAddress$) As Boolean 
 
    z_ScrollbarAction = True 
 
    With session.findById(strAddress).verticalScrollbar 
     
        Dim lng_PageSize& 
        Dim lng_MinRowCurrentPage& 
        Dim lng_MinRowLastPage& 
        Dim lng_MaxRowCurrentPage& 
        Dim lng_MaxRowLastPage& 
         
        lng_PageSize = .pagesize 
        lng_MinRowCurrentPage = .Position 
        lng_MinRowLastPage = .Maximum 
        lng_MaxRowCurrentPage = lng_MinRowCurrentPage + lng_PageSize 
        lng_MaxRowLastPage = lng_MinRowLastPage + lng_PageSize 
     
     
        If lng_PageSize = 0 Then Exit Function 'mean no scroll bar in the table 
     
        'last row of total 
        Select Case lng_RowCursorPosition 
         
            'in the current view not need to scroll 
            Case lng_MinRowCurrentPage To lng_MaxRowCurrentPage 
                z_ScrollbarAction = True 
             
            'out of table view return false 
            Case Is > lng_MaxRowLastPage, Is < 0 
                z_ScrollbarAction = False 
                 
            'In Last Table Line dont scroll 
            Case Is = lng_MaxRowLastPage 
                z_ScrollbarAction = False 
 
                         
            'last row of current view, move scrollbar 
            Case Is = lng_MaxRowCurrentPage 
                Select Case lng_MaxRowCurrentPage 
                    Case Is > lng_MinRowLastPage 
                        .Position = lng_MinRowLastPage 
                    Case Is < lng_MinRowLastPage 
                        .Position = lng_MaxRowCurrentPage '+1 ? 
                End Select 
                 
            'not in view move new to page to the checked row 
            Case Is > lng_MaxRowCurrentPage, Is < lng_MinRowCurrentPage 
                .Position = lng_RowCursorPosition 
 
        End Select 
     
    End With 
 
End Function 
 
 
Public Function ListToArray() As Variant 
 
Dim arrTemp() 
Dim arrCoordTemp 
 
ReDim arrTemp(session.findById(strListAddress).Children.Count - 1, 4) 
With session.findById(strListAddress) 
    For i = 0 To (.Children.Count - 1) 
     
        Dim strId As String 
        Dim arrResult() As String 
 
             
            On Error Resume Next 
             
            With .Children(CLng(i)) 
                strId = .ID 
                arrTemp(i, 0) = .Text 
            End With 
                 
                arrTemp(i, 1) = i 
                 
                arrTemp(i, 2) = strId 
     
     
                Call z_GetCoordsFromAddress(arrTemp(i, 3), arrTemp(i, 4), strId) 
 
    Next 
    End With 
 
    ListToArray = arrTemp 
 
End Function 
 
 
Public Function GoToTransaction(ByRef str_transaction_name$) As Boolean 
'check for top level 
    If Not session.findById("wnd[0]").Text Like "*SAP Easy Access" Then 
        Select Case MsgBox("Close current actions in SAP and run macro?", vbExclamation Or vbOKCancel) 
            Case vbOK: GoToTransaction = True 
                GoToMainPage 
            Case vbCancel: GoToTransaction = False 
                Exit Function 
        End Select 
    End If 
     
    session.findById("wnd[0]/tbar[0]/okcd").Text = str_transaction_name 
    Confirm 
     
End Function 
 
Public Sub PressMainButton(ByRef button As main_button_type) 
    session.findById("wnd[0]/tbar[0]/btn[" & button & "]").press 
     
    Wait 300 
End Sub 
 
Public Sub GoToMainPage() 
                    
    On Error GoTo ERROR_HANDLER 
     
    Dim i&, z&, lngWatchdog& 
    For i = 1 To session.Children.Count 
        session.Children(CLng(i)).sendVKey 12 'press escape 
        session.Children(CLng(i + 1)).findById("usr/btnSPOP-OPTION1").press 
    Next 
     
    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press  'if exit note 
     
    Do While Not session.findById("wnd[0]").Text Like "*SAP Easy Access" And lngWatchdog < 15 
        PressMainButton Exit_15 
         
        'Dim test 
         
        Select Case session.findById("wnd[1]").Text 
            Case "Exit Order Processing": ItemPress ("wnd[1]/usr/btnSPOP-OPTION2") 
            Case Else: ItemPress ("wnd[1]/usr/btnSPOP-OPTION1") 
        End Select 
         
        lngWatchdog = lngWatchdog + 1 
    Loop 
         
    'reset main window level 
    lngCurrentMainWindow = 0 
         
ERROR_HANDLER: 
    If Err.Number = 619 Then Resume Next 
    If Err.Number = 614 Then Resume Next 
 
End Sub 
 
Public Function Statusbar(Optional ByRef lng_word_in_sbar, Optional ByRef str_sbar_condition) As String 
 
    Dim str_status_bar_text 'save statusbar text as temprary 
 
    Statusbar = "" 
    str_status_bar_text = session.findById("wnd[0]/sbar").Text 
     
    If IsMissing(lng_word_in_sbar) Then 
        If IsMissing(str_sbar_condition) Then 
            Statusbar = str_status_bar_text 
        Else 
            If str_status_bar_text Like str_sbar_condition Then 
                Statusbar = str_sbar_condition 
            End If 
        End If 
    Else 
        Dim arrString() As String 
        arrString = Split(str_status_bar_text, " ") 
         
        If IsMissing(str_sbar_condition) Then 
            Statusbar = CStr(arrString(lng_word_in_sbar)) 
        Else 
            If str_status_bar_text Like str_sbar_condition Then 
                Statusbar = CStr(arrString(lng_word_in_sbar)) 
            End If 
        End If 
    End If 
End Function 
 
Public Function StatusbarContains(ParamArray str_sbar_condition()) As Boolean 
     
    Dim i_Cursor As Long 
    Dim str_Statusbar As String 
 
     
    str_Statusbar = session.findById("wnd[0]/sbar").Text 
     
    For i_Cursor = LBound(str_sbar_condition) To UBound(str_sbar_condition) 
        If str_Statusbar Like CStr(str_sbar_condition(i_Cursor)) Then 
            StatusbarContains = True 
            Exit For 
        End If 
    Next 
     
End Function 
 
 
Private Sub z_HandleTab(ByRef str_address_name$) 
    Dim lng_TabWordLenght& 
 
 
 
 
 
 
    If Not str_address_name Like "*/tab*/*" Then Exit Sub 'skip if link to contain tab 
     
    Dim arr_Address 
     
    Dim lng_Cursor As Long 
     
    arr_Address = Split(str_address_name, "/") 
     
    Dim str_TabAddres As String 
     
    For lng_Cursor = LBound(arr_Address) To UBound(arr_Address) 
         
        str_TabAddres = str_TabAddres & "/" & arr_Address(lng_Cursor) 
        If arr_Address(lng_Cursor) Like "tab*" Then 
            On Error Resume Next 
            Me.session.findById(Mid(str_TabAddres, 2)).Select  '    ItemSelect Mid(str_TabAddres, 2) 
            On Error GoTo 0 
        End If 
         
    Next 
 
    'If Not str_address_name Like str_active_tab & "*" Then 'session is on different tab 
    '    str_active_tab = Mid(str_address_name, 1, InStr(str_address_name, "/tabp") + lng_TabWordLenght)  'save tab 
    '    ItemSelect str_active_tab 'switch to tab 
 
    'End If 
 
End Sub 
 
Public Sub DebugPrintDetails(ByRef str_Address$) 
    With session.findById(str_Address) 
        Debug.Print .Name; " "; .Type; " "; .ID 
    End With 
End Sub 
 
Public Sub ComboByKey(ByRef str_combo_key$, ByRef str_combo_address$) 
    z_HandleTab str_combo_address 
    session.findById(str_combo_address).Key = str_combo_key 
End Sub 
 
Public Function ComboListToArray(ByRef str_combo_address$) As Variant 
    Dim arr_combo_list() 'temporary array for storing combo values 
    Dim lng_combo_index& 'index for parsing collection 
 
    z_HandleTab str_combo_address 'switch tab check 
     
    With session.findById(str_combo_address) 
        ReDim Preserve arr_combo_list(1, .entries.Count) 'define array 
         
        For lng_combo_index = 0 To .entries.Count - 1 
            arr_combo_list(0, lng_combo_index) = .entries.Item(CVar(lng_combo_index)).Key 'save combo key to array 0 
            arr_combo_list(1, lng_combo_index) = .entries.Item(CVar(lng_combo_index)).Value 'save combo value to array 1 
         
        Next 
    End With 
 
    ComboListToArray = arr_combo_list 'return list back to array 
 
End Function 
 
Public Function DialogLabelMultiList(ByRef lngAction&, ByRef lngIndexWindow&, ByRef boolLastItem As Boolean, ByRef strValue$, ByRef strIssueText$) As Boolean 
    Dim lngViewRows& 
    Dim lngDialogRowsCount& 
    Dim lngScrollBarPosition& 
    Dim lngHeaderRows& 'number of rows taken from header 
    Dim lngConfirmIndex& 'last field should be confirmed from button with different index 
     
     
    'Predifined Corrections 
    'last button 
    If boolLastItem Then 
        lngConfirmIndex = 2 
    Else 
        lngConfirmIndex = 6 
    End If 
     
    'header two items special label correction 
    If lngIndexWindow Mod 2 <> 0 Then 
        lngHeaderRows = (((lngIndexWindow - 1) * 3)) + ((lngIndexWindow + 1) / 2) 
    Else 
        lngHeaderRows = (((lngIndexWindow - 1) * 3)) + (lngIndexWindow / 2) 
    End If 
    
    'scroll bar handling 
    With session.findById("wnd[" & lngIndexWindow & "]/usr/").verticalScrollbar 
        lngViewRows = .pagesize 
        lngDialogRowsCount = .Maximum 
        lngScrollBarPosition = .Position 
         
On Error GoTo Err 
 
        Dim i As Long 
 
        Do 
         
            For i = lngHeaderRows To (session.findById("wnd[" & lngIndexWindow & "]/usr/").Children.Count - 1) Step 3 
                 
                If strValue = session.findById("wnd[" & lngIndexWindow & "]/usr/").Children(CLng(i)).Text Then 
                    session.findById("wnd[" & lngIndexWindow & "]/usr/").Children(CLng(i + 2)).SetFocus 
                    session.findById("wnd[" & lngIndexWindow & "]/tbar[0]/btn[" & lngConfirmIndex & "]").press 
                    GoTo FOUND 
                End If 
                 
NEXT_ITEM: 
            Next 
                 
            'scroll bar 
            .Position = lngViewRows + .Position 
 
        Loop While (lngViewRows + .Position) <> (.Maximum + lngViewRows) 
    End With 
    'not found returns 
NOT_FOUND: 
    DialogLabelMultiList = False 
     
    session.findById("wnd[" & lngIndexWindow & "]/usr/").verticalScrollbar.Position = 0 'reset window 
    session.findById("wnd[" & lngIndexWindow & "]/usr/").Children(CLng(lngHeaderRows + 2)).SetFocus 'select first item 
     
    LogAdd "", "", strIssueText & " - " & CStr(strValue & "- not found in list selected instead " & session.findById("wnd[" & lngIndexWindow & "]/usr/").Children(CLng(lngHeaderRows + 2)).Text & " " & session.findById("wnd[" & lngIndexWindow & "]/usr/").Children(CLng(lngHeaderRows)).Text), info 
     
    session.findById("wnd[" & lngIndexWindow & "]/tbar[0]/btn[" & lngConfirmIndex & "]").press 'confirm 
 
Exit Function 
 
FOUND: 
    DialogLabelMultiList = True 
    Exit Function 
 
 
Err: 
    Select Case Err.Number 
        Case 614 
            Resume NOT_FOUND 
        Case Else 
            Debug.Assert Err.Number <> 0 
    End Select 
     
End Function 
 
Public Function DialogLabelSimpleList(ByRef lngValueSwitch&, ByRef strValue$, ByRef strAddressOrNumber) As Boolean 
 
    Dim lngViewRows& 
    Dim lngDialogRowsCount& 
    Dim lngScrollBarPosition& 
    Dim lngHeaderRows& 'number of rows taken from header 
    Dim lngConfirmIndex& 'last field should be confirmed from button with different index 
    Dim lngIndexWindow& 
     
    'check what type of is entry value 
    Select Case VarType(strAddressOrNumber) 
        Case vbString:   lngIndexWindow = CStr(Mid(strAddressOrNumber, 6, 1)) + 1 'entered item address 
        Case vbLong, vbInteger:  lngIndexWindow = strAddressOrNumber 'entered just level of window 
    End Select 
    
    If SapObjectExist("wnd[" & lngIndexWindow & "]") Then 'if file dialog not yet exist 
        If VarType(strAddressOrNumber) = vbString Then ItemFocus CStr(strAddressOrNumber) 'if item address focus item 
        session.findById("wnd[" & lngIndexWindow - 1 & "]").sendVKey 4 'raise Dialog 
    End If 
     
    With session.findById("wnd[" & lngIndexWindow & "]/usr/").verticalScrollbar 
        lngViewRows = .pagesize 
        lngDialogRowsCount = .Maximum 
        lngScrollBarPosition = .Position 
         
        If lngViewRows = 0 Then lngViewRows = 25 
        .Position = 0 
    End With 
 
On Error GoTo Err 
 
    Do 
        Dim i As Long 
     
     
        For i = 0 To (session.findById("wnd[" & lngIndexWindow & "]/usr/").Children.Count - 1) Step 2 
             
            If strValue = session.findById("wnd[" & lngIndexWindow & "]/usr/").Children(CLng(i + lngValueSwitch)).Text Then 
                session.findById("wnd[" & lngIndexWindow & "]/usr/").Children(CLng(i)).SetFocus 
                session.findById("wnd[" & lngIndexWindow & "]/tbar[0]/btn[0]").press 
                GoTo FOUND 
            End If 
             
NEXT_ITEM: 
        Next 
             
        With session.findById("wnd[" & lngIndexWindow & "]/usr/").verticalScrollbar 
            .Position = lngViewRows + .Position 
            lngDialogRowsCount = .Maximum 
        End With 
         
    Loop While True 
 
FOUND: 
    DialogLabelSimpleList = True 
    Exit Function 
 
NOT_FOUND: 
    DialogLabelSimpleList = False 
 
 
Err: 
    Select Case Err.Number 
        Case 614 
            Resume NOT_FOUND 
        Case Else 
            Debug.Assert Err.Number <> 0 
    End Select 
 
End Function 
 
Public Function DialogTextCondition(ByRef lngTextCondition As text_condition, ByRef strAddress$) 
 
    Dim strWindowIndex$ 
 
    Call EditItem(strAddress) 
     
    strWindowIndex = CLng(Mid(strAddress, 5, 1)) + 1 
     
    With session.findById("wnd[" & strWindowIndex & "]/usr/cntlOPTION_CONTAINER/shellcont/shell") 
        .setCurrentCell lngTextCondition, "TEXT" 
        .selectedRows = CStr(lngTextCondition) 
        .doubleClickCurrentCell 
    End With 
End Function 
 
Public Sub ExcelMinimize() 
 
    Dim RetVal$ 
    RetVal = ShowWindow(FindWindow("XLMAIN", vbNullString), 2) 
 
End Sub 
 
Public Sub ExcelMaximize() 
 
    Dim hWnd As Long 
 
    Dim RetVal$ 
    RetVal = ShowWindow(FindWindow("XLMAIN", vbNullString), 1) 
 
End Sub 
 
Private Sub z_InitiateLog(Optional ByRef t_Log As ListObject, Optional ByVal lng_CurrentPhase As Long = 0) 
     
On Error GoTo Err 
     
    Dim rng_Cursor As Range 
     
 
    Set cc_BlackList = New clsCollection 
    Set cc_IssueLog = New clsCollection 
 
    Dim arr_Log 
    Dim lng_Cursor As Long 
 
    If t_Log Is Nothing Then Exit Sub 'if log is not specified Quit 
     
    Set t_RunLog = t_Log 
     
    arr_Log = t_Log.Range.Value ' shtLog.Range("A1").CurrentRegion.Value 
 
    Const RecordId As Integer = 1 
    Const RecordValue As Integer = 2 
    Const ItemId As Integer = 3 
    Const ItemValue As Integer = 4 
    Const Message As Integer = 5 
    Const Status As Integer = 6 
    Const TimeStamp As Integer = 7 
    Const Phase As Integer = 8 
   ' End Enum 
         
    Dim arr_Row 
     
    For lng_Cursor = LBound(arr_Log) + 1 To UBound(arr_Log) 
 
        arr_Row = WorksheetFunction.Index(arr_Log, lng_Cursor) 
         
        If IsEmpty(arr_Row(RecordId)) Then GoTo NEXT_LOOP:  'Exit Sub 'skip if there are no records 
         
        If arr_Row(Status) = "OK" And arr_Row(Phase) = lng_CurrentPhase Then   ' rng_Cursor.Value = "OK" And rng_Cursor.Offset(, 2) = lng_CurrentPhase Then 
 
            'skip blacklisted columns 
            'if record column is empty then behave same as issue column 
            If CStr(arr_Row(RecordId)) <> "" Then 
                If Not cc_BlackList.KeyExists(CStr(arr_Row(RecordValue))) Then 
                    cc_BlackList.Add CStr(arr_Row(RecordValue)), CStr(arr_Row(RecordValue)) 
                End If 
            End If 
             
        End If 
         
        'Dont Reload EE - Error Items for the run 
        If Not (arr_Row(Status) = "EE" And arr_Row(Phase) = lng_CurrentPhase) Then 
            cc_IssueLog.Add arr_Row 
        End If 
        
NEXT_LOOP: 
 
    Next 
     
Exit Sub 
Err: 
    Err.Raise Err.Number 
    Resume 
         
     
End Sub 
 
Public Sub LogAdd(Optional ByVal str_itemId$, Optional ByVal str_itemValue$, Optional ByRef str_result, Optional lng_LogActionFlag As log_flag = info) 
 
    Dim str_RecordValue$ 
    'Dim str_itemValue$ 
    Dim str_LogMessageType$ 
    Dim str_recordId As String 
 
 
    'Message Code 
    Select Case lng_LogActionFlag 
        Case info: str_LogMessageType = "II" 
        Case Ok: str_LogMessageType = "OK" 
        Case warning: str_LogMessageType = "WW" 
        Case ISSUE: str_LogMessageType = "EE" 
    End Select 
 
    'Record ID 
    str_recordId = str_CurrentKeyColumnName 
    If Not cc_CurrentCursors Is Nothing Then 
        If cc_CurrentCursors.KeyExists("Main") Then 
        str_RecordValue = cc_CurrentCursors.Item("Main").Value 
 
            'Black list flag to skip record from next processing run 
            If lng_LogActionFlag = Ok Then 
                If Not cc_BlackList.KeyExists(str_RecordValue) Then 'check if not exist 
                    cc_BlackList.Add str_RecordValue, str_RecordValue 
                End If 
            End If 
                 
        End If 
    End If 
 
 
 
         
    If cc_IssueLog Is Nothing Then Set cc_IssueLog = New clsCollection 
'FINAL LOG LINE 
 
    Dim arrLogLine(1 To 8) 
 
    arrLogLine(1) = str_recordId 'Array(str_recordId, str_RecordValue, str_itemId, str_itemValue, str_result, str_LogMessageType, CDbl(Now), lng_CurrentPhase) 
    arrLogLine(2) = str_RecordValue 
    arrLogLine(3) = str_itemId 
    arrLogLine(4) = str_itemValue 
    arrLogLine(5) = str_result 
    arrLogLine(6) = str_LogMessageType 
    arrLogLine(7) = CDbl(Now) 
    arrLogLine(8) = lng_CurrentPhase 
 
    cc_IssueLog.Add (arrLogLine) 
 
End Sub 
 
Public Function LogAsArray(ByRef arr_IssueLog) As Boolean 
     
    LogAsArray = False 
 
    'exit skip if there are no items in the issue log 
    If cc_IssueLog Is Nothing Then Exit Function 
    If cc_IssueLog.Count = 0 Then Exit Function 
         
            
    arr_IssueLog = cc_IssueLog.ItemsInArray 
     
 
    LogAsArray = True 
 
End Function 
 
Public Sub LogToRange(Optional ByRef rng_reference As Range = Nothing) 
 
    Dim arr_item 
 
    'check if issue log have items 
    If LogAsArray(arr_item) Then 
 
        If rng_reference Is Nothing Then 
            If Not t_RunLog Is Nothing Then 
                t_RunLog.Range.Offset(1).Resize(UBound(arr_item, 1), UBound(arr_item, 2)) = arr_item 
            End If 
        Else 
            rng_reference.Resize(UBound(arr_item, 1) + 1, UBound(arr_item, 2)) = arr_item 
        End If 
    End If 
 
End Sub 
 
 
Public Function ExcelSetRunTables( _ 
    Optional ByVal t_ParameterTable As Object, _ 
    Optional ByVal t_Log As Object, _ 
    Optional ByVal lng_LogRun As Long = 0) 
 
On Error GoTo Err: 
 
    Dim arr_Parameters 
    Dim lng_Cursor As Long 
     
'PARAMETER SETUP 
    
    If Not t_ParameterTable Is Nothing Then 
     
        arr_Parameters = t_ParameterTable.DataBodyRange.Value 
 
        If TypeName(t_ParameterTable) <> "ListObject" Then Debug.Print "error" 
 
        Set cc_RunParameters = New clsCollection 
         
        For lng_Cursor = LBound(arr_Parameters) To UBound(arr_Parameters) 
            Call cc_RunParameters.Add(arr_Parameters(lng_Cursor, 2), arr_Parameters(lng_Cursor, 1)) 
        Next 
 
    End If 
     
'HEADER RECORD INITIATION 
'LOG INITIATION 
 
    lng_CurrentPhase = lng_LogRun 
 
    'Initiate Log Moved Here 
    Call z_InitiateLog(t_Log, lng_LogRun) 
 
Exit Function 
Err: 
    Debug.Assert False 
    Resume 
         
End Function 
 
 
Public Function ExcelLoadItemData( _ 
    Optional ByVal t_ItemTable As Object, _ 
    Optional ByVal str_ItemGroupByColumnName As String = "", _ 
    Optional ByVal str_ItemGroupName As String = "defaultItems") 
 
     'Dim rng_Cursor As Range 
 
    If Not bool_HeaderLoaded Then 
        MsgBox "Can't load data without initiating Header Exitting" 
    End If 
 
 
    Dim arr_Parameters 
    Dim lng_Cursor As Long 
     
'DATA NAMES 
 
    Dim arr_GroupByColumn 
    Dim arr_KeyGroupByColumn 
     
    Dim str_TableName As String 
     
    str_TableName = t_ItemTable.Name 
 
    Dim bool_GroupByItems As Boolean 
     
    bool_GroupByItems = str_ItemGroupByColumnName <> "" 
 
     
'SAVE DATA AND HEADER 
    Call z_SetHeaderAndValues(WorksheetFunction.Index(t_ItemTable.HeaderRowRange.Value, 0), t_ItemTable.DataBodyRange.Value, str_TableName, str_ItemGroupName) 
 
'VALIDATE KEYCOLUMN IS PRESENT 
    If Not cc_DataHeaderOfTables.Item(str_TableName).KeyExists(str_CurrentKeyColumnName) Then 
        MsgBox str_CurrentKeyColumnName & "Header column is not found": Exit Function 
    End If 
 
    'VALIDATE KEYCOLUMN IS PRESENT 
    If bool_GroupByItems Then 
        If Not cc_DataHeaderOfTables.Item(str_TableName).KeyExists(str_ItemGroupByColumnName) Then 
            MsgBox str_ItemGroupByColumnName & "Header column is not found": Exit Function 
        End If 
    End If 
        
 
'GROUP BY 
    arr_KeyGroupByColumn = WorksheetFunction.Transpose(WorksheetFunction.Index(cc_DataValueOfTables.Item(str_TableName), 0, cc_DataHeaderOfTables.Item(str_TableName).Item(str_CurrentKeyColumnName))) 
     
    If bool_GroupByItems Then 
        arr_GroupByColumn = WorksheetFunction.Transpose(WorksheetFunction.Index(cc_DataValueOfTables.Item(str_TableName), 0, cc_DataHeaderOfTables.Item(str_TableName).Item(str_ItemGroupByColumnName))) 
    End If 
     
     
    Dim str_CurrentRowUniqueName  As String 
  
    For lng_Cursor = LBound(arr_KeyGroupByColumn) To UBound(arr_KeyGroupByColumn) 
 
        str_CurrentRowUniqueName = str_TableName & "_" & CStr(arr_KeyGroupByColumn(lng_Cursor)) & "_" & str_ItemGroupName 
     
        If Not cc_TableRowIndexes.KeyExists(str_CurrentRowUniqueName) Then 
             
            Dim cc_RowsIndexesTemp As clsCollection 
            Set cc_RowsIndexesTemp = New clsCollection 
            Call cc_TableRowIndexes.Add(cc_RowsIndexesTemp, str_CurrentRowUniqueName) 
             
        End If 
         
        'ITEM ROW LIST UPDATE 
        With cc_TableRowIndexes.Item(str_CurrentRowUniqueName) 
         
            If str_ItemGroupByColumnName = "" Then 
                Call .Add(lng_Cursor) 'only add row 
            Else 
                If Not .KeyExists(arr_GroupByColumn(lng_Cursor)) Then 
                    Call .Add(lng_Cursor, arr_GroupByColumn(lng_Cursor)) 'add row with particular key to prevent duplications 
                End If 
            End If 
             
        End With 
         
    Next 
     
End Function 
 
Public Function z_SetHeaderAndValues(arr_HeaderRow, arr_DataField, str_TableName, str_ItemName) 
 
    If str_ItemName = "header" Then 'IF HEADER RESET 
        Set cc_TableNames = New clsCollection 
    End If 
     
    If cc_TableNames.KeyExists(str_TableName) Then Exit Function 
 
'GET TABLE NAME 
 
    Call cc_TableNames.Add(str_TableName, str_ItemName) 
  
'DATA SAVE 
 
     
    Set cc_DataValueOfTables = New clsCollection 
     
    Call cc_DataValueOfTables.Add(arr_DataField, str_TableName) 
        
        
'HEADER INDEXES 
 
    Dim cc_HeadColumnIndex As clsCollection 
     
    Set cc_HeadColumnIndex = New clsCollection 
 
    Dim lng_Cursor As Long 
 
    For lng_Cursor = LBound(arr_HeaderRow) To UBound(arr_HeaderRow) 
        Call cc_HeadColumnIndex.Add(lng_Cursor, arr_HeaderRow(lng_Cursor)) 
    Next 
 
    Set cc_DataHeaderOfTables = New clsCollection 
     
    Call cc_DataHeaderOfTables.Add(cc_HeadColumnIndex, str_TableName) 
     
End Function 
 
 
 
Public Function RecordNext(ByVal t_MainTable As ListObject, ByVal str_KeyColumn As String) As Boolean 
 
    'Dim i& 'black list cursor 
 
    Dim bool_NextCursor As Boolean 
    Dim rng_Cursor As Range 
 
NEXT_ITEM: 
 
    bool_NextCursor = shtATk.TableLoop(t_MainTable, rng_Cursor, str_KeyColumn, e_UniquesWithFilter, e_ClearCurrentFilters) 
 
    str_CurrentKeyColumnName = str_KeyColumn 
 
    'CURSORS CHECK 
    If cc_CurrentCursors Is Nothing Then Set cc_CurrentCursors = New clsCollection 
     
    'CLEAR CURRENT CURSOR IF EXITST 
    If cc_CurrentCursors.KeyExists("Main") Then Call cc_CurrentCursors.Remove("Main") 
     
    'LAST RUN CLEANUP 
    If bool_NextCursor Then 
     
    'FIRST RUN 
        If cc_BlackList.KeyExists(CStr(rng_Cursor.Value)) Then GoTo NEXT_ITEM: 
        'If RecordBlackListCheck Then GoTo NEXT_ITEM: 
 
        Call cc_CurrentCursors.Add(rng_Cursor, "Main") 'add actual cursor to next 
   
    End If 
     
    RecordNext = bool_NextCursor 
     
End Function 
 
 
 
Public Function Record(ByRef str_ColumnName$, Optional ByRef lng_FormatReturned As input_Format = input_Format.TextFormat) As String 
 
    Record = z_FormatInput(CStr(shtATk.TableLoopCursor(cc_CurrentCursors.Item("Main"), str_ColumnName)), lng_FormatReturned) 
 
End Function 
 
Public Function Parameter(str_ParameterName As String, Optional ByRef lng_FormatReturned As input_Format = input_Format.TextFormat) As String 
 
    Parameter = z_FormatInput(cc_RunParameters.Item(str_ParameterName), lng_FormatReturned) 
 
End Function 
 
 
Public Function RecordBlackListCheck() As Boolean 
 
    Dim bool_BlackListed As Boolean 
 
    Dim i& 
 
    bool_BlackListed = False 
 
    If cc_BlackListedRecordColumns.Count > 0 Then 
        For i = 1 To cc_BlackListedRecordColumns.Count 
            If cc_BlackList.KeyExists(Record(cc_BlackListedRecordColumns.Item(i))) Then bool_BlackListed = True 
        Next 
    End If 
 
    RecordBlackListCheck = bool_BlackListed 
 
End Function 
 
 
Public Function RecordItemNext(ByVal t_MainTable As ListObject, ByVal str_KeyColumn As String, Optional lng_LoopType As e_ItemLoopType = e_Items, Optional ByVal str_ItemGroupName As String = "defaultItems") As Boolean 
 
    Dim bool_NextCursor As Boolean 
    Dim rng_Cursor As Range 
 
NEXT_ITEM: 
 
    bool_NextCursor = shtATk.TableLoop(t_MainTable, rng_Cursor, str_KeyColumn, lng_LoopType) 
   
    If cc_CurrentCursors Is Nothing Then Set cc_CurrentCursors = New clsCollection 
 
'SAVE CURRENT CURSOR 
    If cc_CurrentCursors.KeyExists(str_ItemGroupName) Then Call cc_CurrentCursors.Remove(str_ItemGroupName) 
 
    Call cc_CurrentCursors.Add(rng_Cursor, str_ItemGroupName) 'add actual cursor to next 
     
    RecordItemNext = bool_NextCursor 
 
    'RecordItemNext = cc_TableRowIndexes.Item(cc_TableNames.Item(str_ItemGroupName) & "_" & str_CurrentKeyColumnValue & "_" & str_ItemGroupName).MoveNext 
 
End Function 
 
 
Public Function RecordItem(ByRef strItemDetailName$, Optional ByRef lng_FormatReturned As input_Format = input_Format.TextFormat, Optional ByRef str_ItemGroupName = "defaultItems") As String 
   
    RecordItem = z_FormatInput(CStr(shtATk.TableLoopCursor(cc_CurrentCursors.Item(str_ItemGroupName), strItemDetailName)), lng_FormatReturned) 
 
     
End Function 
 
 
Public Function ClipboardFillWithString(ByRef str_InputString As String) As Boolean 
    'modified chip pearsons macro 
 
    Dim obj_Clipboard As Object 
     
    On Error GoTo ErrH: 
    If obj_Clipboard Is Nothing Then 
        Set obj_Clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'late bind clipboard address 
    End If 
     
    With obj_Clipboard 
        .SetText str_InputString 
        .PutInClipboard 
    End With 
     
    fbool_PutInClipboard = True 
    Exit Function 
ErrH: 
    fbool_PutInClipboard = False 
    Exit Function 
End Function 
 
Function ClipboardToTextFile(ByRef str_FileName$, Optional ByRef bool_SkipHeader As Boolean, Optional ByRef ColumnDelimeter = "|", Optional ByRef arr_ColumnNames) As Boolean 
 
ClipboardToTextFile = False 
 
On Error GoTo Err: 
 
Dim Clip As Object 
Set Clip = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 
    
Dim i_Rows& 
Dim i_Cells& 
 
Dim arr_Rows() As String 
Dim arr_Cells() As String 
 
    Dim lng_HeaderRowCorrection& 
    Dim lng_HeaderRow& 
     
    'check if header should be skipped 
    If bool_SkipHeader Then 
        lng_HeaderRowCorrection = 1 
    Else 
        lng_HeaderRowCorrection = 0 
    End If 
     
    'Get RAW array from clipboard and remove object 
    Clip.GetFromClipboard 
    arr_Rows = Split(Clip.GetText, vbNewLine) 
    Set Clip = Nothing 
     
    Dim lng_TableRowsCount& 
    Dim lng_TableColumnsCount& 
     
     
    Dim lng_RowsStart& 
    Dim lng_RowEnd& 
     
    lng_RowsStart = LBound(arr_Rows) 
    lng_RowEnd = UBound(arr_Rows) 
     
    'FILL ROWS 
    For i_Rows = lng_RowsStart To lng_RowEnd 
        If arr_Rows(i_Rows) Like "|*" Then 
             
            If lng_HeaderRow = 0 Then 
                lng_HeaderRow = i_Rows 'first row is header row 
                lng_TableColumnsCount = UBound(Split(Mid(arr_Rows(i_Rows), 2, Len(arr_Rows(i_Rows)) - 1), "|")) 'get number of columns 
                Exit For 
            End If 
        End If 
    Next 
         
    'Just count rows 
    For i_Rows = lng_RowsStart To lng_RowEnd 
        If arr_Rows(i_Rows) Like "|*" Then lng_TableRowsCount = lng_TableRowsCount + 1 
    Next i_Rows 
     
    'check if there are some rows 
    If lng_TableRowsCount = 0 Then Exit Function 
 
 
    'SETUP HEADER ORDER 
    Dim arr_ColumnSourcePosition() 
    'Dim arr_ColumnResultPosition() 
    Dim i_HeaderRow& 
     
    If Not IsMissing(arr_ColumnNames) Then 
        ReDim arr_ColumnSourcePosition(UBound(arr_ColumnNames)) 
        'ReDim arr_ColumnResultPosition(UBound(arr_ColumnNames)) 
         
        Dim cc_HeaderPositions As clsCollection 
        Set cc_HeaderPositions = New clsCollection 
         
         
        arr_Cells = Split(Mid(arr_Rows(lng_HeaderRow), 2, Len(arr_Rows(lng_HeaderRow)) - 1), "|") 
         
        'Save columns positions 
        For i_Cells = LBound(arr_Cells) To UBound(arr_Cells) 
            Call cc_HeaderPositions.Add(i_Cells, Trim(arr_Cells(i_Cells))) 
        Next 
             
        'Link positions with with output layout 
        For i_Cells = LBound(arr_ColumnNames) To UBound(arr_ColumnNames) 
            If cc_HeaderPositions.KeyExists(arr_ColumnNames(i_Cells)) Then 
                arr_ColumnSourcePosition(i_Cells) = cc_HeaderPositions.Item(arr_ColumnNames(i_Cells)) 
            Else 
                arr_ColumnSourcePosition(i_Cells) = -1 
            End If 
        Next 
         
        lng_TableColumnsCount = UBound(arr_ColumnNames) + 1 
         
    End If 
 
    'CREATE TEXT FILE 
    Dim outputTextFile As Object 'Could be declared as a TextStream 
    Dim fso As Object 
    Set fso = CreateObject("Scripting.FileSystemObject") 
    Dim i As Long, j As Long, file2Write As Long 
    Dim ubound1 As Long, ubound2 As Long 
    Dim recordsetDataArray As Variant 
    Dim strText As String 
    Set outputTextFile = fso.CreateTextFile(ThisWorkbook.Path & "\" & str_FileName, True) 
 
    'SETUP TABLE 
    ReDim arr_Result(lng_TableRowsCount - 1 - lng_HeaderRowCorrection, lng_TableColumnsCount - 1) 
 
 
    If bool_SkipHeader Then 
        lng_RowsStart = lng_HeaderRow + 1 
    Else 
        lng_RowsStart = LBound(arr_Rows) 
    End If 
 
'Pepare formating for Data 
    If Not IsMissing(arr_ColumnNames) Then 
        Dim lng_ColumnArrStart& 
        Dim lng_ColumnArrEnd& 
         
        lng_ColumnArrEnd = UBound(arr_ColumnSourcePosition) 
        lng_ColumnArrStart = LBound(arr_ColumnSourcePosition) 
     
        Dim arr_TempRow() 
        ReDim arr_TempRow(UBound(arr_ColumnSourcePosition)) 
    End If 
 
    'create temporary array 
    Dim arr_ResultArray() 
    Dim lng_CurrentLine& 
     
    ReDim arr_ResultArray(0) 
    lng_CurrentLine = 0 
     
    'FILL ResultArray 
    For i_Rows = lng_RowsStart To lng_RowEnd 
             
        If arr_Rows(i_Rows) Like "|*" Then 
         
            ReDim Preserve arr_ResultArray(lng_CurrentLine) 
            
            If IsMissing(arr_ColumnNames) Then 'to save some performance are array desision above inner array (cells)loops 
               arr_ResultArray(lng_CurrentLine) = arr_Rows(i_Rows) 
            Else 
                arr_Cells = Split(Mid(arr_Rows(i_Rows), 2, Len(arr_Rows(i_Rows)) - 1), "|") 
         
                For i_Cells = lng_ColumnArrStart To lng_ColumnArrEnd 
                    If arr_ColumnSourcePosition(i_Cells) <> -1 Then 
                        arr_TempRow(i_Cells) = Trim(arr_Cells(arr_ColumnSourcePosition(i_Cells))) 
                    Else 
                        arr_TempRow(i_Cells) = vbNullString 
                    End If 
                Next i_Cells 
                 
                arr_ResultArray(lng_CurrentLine) = Join(arr_TempRow, "|") 
            End If 
             
            lng_CurrentLine = lng_CurrentLine + 1 
             
        End If 
         
    Next i_Rows 
     
     
    ReDim Preserve arr_ResultArray(lng_CurrentLine - 1) 'remove last empty line 
         
    With outputTextFile 
         
        .Write Join(arr_ResultArray, vbNewLine) 
        .Close 
     
    End With 
     
    ClipboardToTextFile = True 
 
Exit Function 
 
Err: 
If Err.Number <> 0 Then 
    Debug.Assert False 
    Resume 
End If 
 
End Function 
 
Public Function ClipboardToArray(ByRef arr_Result, Optional ByRef bool_SkipHeader As Boolean, Optional ByRef NewLineDelimeter = vbNewLine, Optional ByRef ColumnDelimeter = "|", Optional ByRef arr_ColumnNames) As Boolean 
 
ClipboardToArray = False 
 
On Error GoTo Err: 
 
Dim Clip As Object 
Set Clip = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 
    
Dim i_Rows& 
Dim i_Cells& 
 
Dim arr_Rows() As String 
Dim arr_Cells() As String 
 
    Dim lng_HeaderRowCorrection& 
    Dim lng_HeaderRow& 
     
    'check if header should be skipped 
    If bool_SkipHeader Then 
        lng_HeaderRowCorrection = 1 
    Else 
        lng_HeaderRowCorrection = 0 
    End If 
     
    'Get RAW array from clipboard and remove object 
    Clip.GetFromClipboard 
    arr_Rows = Split(Clip.GetText, vbNewLine) 
    Set Clip = Nothing 
     
    Dim lng_TableRowsCount& 
    Dim lng_TableColumnsCount& 
     
    'FILL ROWS 
    For i_Rows = LBound(arr_Rows) To UBound(arr_Rows) 
         
        If arr_Rows(i_Rows) Like "|*" Then 
            lng_TableRowsCount = lng_TableRowsCount + 1 
            If lng_HeaderRow = 0 Then lng_HeaderRow = i_Rows 'first row is header row 
            If lng_TableColumnsCount = 0 Then lng_TableColumnsCount = UBound(Split(Mid(arr_Rows(i_Rows), 2, Len(arr_Rows(i_Rows)) - 1), "|")) 'get number of columns 
        End If 
         
    Next i_Rows 
 
    'check if there are some rows 
    If lng_TableRowsCount = 0 Then Exit Function 
 
 
    'SETUP HEADER ORDER 
    Dim arr_ColumnSourcePosition() 
    'Dim arr_ColumnResultPosition() 
    Dim i_HeaderRow& 
     
    If Not IsMissing(arr_ColumnNames) Then 
        ReDim arr_ColumnSourcePosition(UBound(arr_ColumnNames)) 
        'ReDim arr_ColumnResultPosition(UBound(arr_ColumnNames)) 
         
        Dim cc_HeaderPositions As clsCollection 
        Set cc_HeaderPositions = New clsCollection 
         
         
        arr_Cells = Split(Mid(arr_Rows(lng_HeaderRow), 2, Len(arr_Rows(lng_HeaderRow)) - 1), "|") 
         
        'Save columns positions 
        For i_Cells = LBound(arr_Cells) To UBound(arr_Cells) 
            Call cc_HeaderPositions.Add(i_Cells, Trim(arr_Cells(i_Cells))) 
        Next 
             
        'Link positions with with output layout 
        For i_Cells = LBound(arr_ColumnNames) To UBound(arr_ColumnNames) 
            If cc_HeaderPositions.KeyExists(arr_ColumnNames(i_Cells)) Then 
                arr_ColumnSourcePosition(i_Cells) = cc_HeaderPositions.Item(arr_ColumnNames(i_Cells)) 
            Else 
                arr_ColumnSourcePosition(i_Cells) = -1 
            End If 
        Next 
         
        lng_TableColumnsCount = UBound(arr_ColumnNames) + 1 
         
    End If 
 
 
 
    'SETUP TABLE 
    ReDim arr_Result(lng_TableRowsCount - 1 - lng_HeaderRowCorrection, lng_TableColumnsCount - 1) 
    Dim lng_RowsCount& 
 
    lng_RowsCount = 0 - lng_HeaderRowCorrection 
 
    'FILL ResultArray 
    For i_Rows = LBound(arr_Rows) To UBound(arr_Rows) 
             
        If arr_Rows(i_Rows) Like "|*" Then 
         
            If bool_SkipHeader Then: If i_Rows = lng_HeaderRow Then GoTo NEXT_ROW: 'skip header row 
             
            arr_Cells = Split(Mid(arr_Rows(i_Rows), 2, Len(arr_Rows(i_Rows)) - 1), "|") 
     
            If IsMissing(arr_ColumnNames) Then 'to save some performance are array desision above inner array (cells)loops 
                For i_Cells = LBound(arr_Result, 2) To UBound(arr_Result, 2) 
                    arr_Result(lng_RowsCount, i_Cells) = Trim(arr_Cells(i_Cells)) 
                Next i_Cells 
            Else 
                For i_Cells = LBound(arr_Result, 2) To UBound(arr_Result, 2) 
                    If arr_ColumnSourcePosition(i_Cells) <> -1 Then 
                        arr_Result(lng_RowsCount, i_Cells) = Trim(arr_Cells(arr_ColumnSourcePosition(i_Cells))) 
                    End If 
                Next i_Cells 
            End If 
             
NEXT_ROW: 
 
            lng_RowsCount = lng_RowsCount + 1 
             
        End If 
         
    Next i_Rows 
     
    ClipboardToArray = True 
 
Exit Function 
 
Err: 
If Err.Number <> 0 Then 
    Debug.Assert False 
    Resume 
End If 
 
End Function 
 
Public Function ClipboardFillViaDownload() As Boolean 
 
    'Procedure from SE16 not tested anywhere else yet 
 
    SelectItem "wnd[0]/mbar/menu[1]/menu[5]" 'select download 
    SelectItem "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]" 
    PressItem "wnd[1]/tbar[0]/btn[0]" 
 
End Function 
Public Function SapCsvCleanFix(ByRef str_CsvPath As String) As Boolean 
     
    On Error GoTo Err: 
     
    Dim lng_ActiveColumn& 
    Dim i_LineCounter& 
    Dim i_HeaderCursor 
    Dim i_DataCursor 
     
    Dim arr_CsvSource 
    Dim arr_TempResult 
    Dim arr_ActiveColumns() 
    Dim arr_TempHeader 
    Dim arr_TempData 
     
    ReDim arr_ActiveColumns(1, 0) 
     
    With CreateObject("Scripting.FileSystemObject") 
     
        With .OpenTextFile(str_CsvPath, 1) 
 
            arr_CsvSource = Split(.ReadAll, vbCrLf) 
 
            Dim i_Cursor& 
            Dim str_HeaderLine As String 
 
            i_LineCounter = LBound(arr_CsvSource) 
     
            ReDim arr_TempResult(0 To UBound(arr_CsvSource)) 
     
            For i_Cursor = LBound(arr_CsvSource) To UBound(arr_CsvSource) 
                If InStr(1, arr_CsvSource(i_Cursor), "|") > 0 Then 
                 
                    If i_LineCounter > 0 And str_HeaderLine = arr_CsvSource(i_Cursor) Then GoTo NEXT_LINE:  'header Check 
                     
                    Dim arr_Line 
                    Dim c_Cursor As Long 
                    arr_Line = Split(arr_CsvSource(i_Cursor), "|") 
                     
                    For c_Cursor = LBound(arr_Line) To UBound(arr_Line) 
                        arr_Line(c_Cursor) = Replace(Trim(arr_Line(c_Cursor)), """", """""""""") 
                    Next 
                     
                    arr_TempResult(i_LineCounter) = """=""""" & Join(arr_Line, """""""|""=""""") & """""""" 
                     
                    str_HeaderLine = arr_CsvSource(i_Cursor) 
                    i_LineCounter = i_LineCounter + 1 
                         
NEXT_LINE: 
                End If 
            Next 
             
            ReDim Preserve arr_TempResult(1 To i_LineCounter) 
             
        End With 
     
    End With 
 
    Dim file_Output As Object 
     
    With CreateObject("Scripting.FileSystemObject") 
     
        'write file 
        Set file_Output = .CreateTextFile(str_CsvPath, True) 
     
        Call file_Output.WriteLine("sep=|" & vbCrLf & Join(arr_TempResult, vbCrLf)) 
        Call file_Output.Close 
 
    End With 
 
    'CLEARing 
    Erase arr_TempResult 
 
 
 
Err: 
    'Resume 
     
End Function 
 
Public Function SapCsvToArray( _ 
                    ByRef arr_Result, _ 
                    ByRef str_CsvPath As String, _ 
                    Optional ByVal arr_HeaderMask) As Boolean 
     
    On Error GoTo Err: 
     
    Dim lng_ActiveColumn& 
    Dim i_LineCounter& 
    Dim i_HeaderCursor 
    Dim i_DataCursor 
     
    Dim arr_CsvSource 
    Dim arr_TempResult 
    Dim arr_ActiveColumns() 
    Dim arr_TempHeader 
    Dim arr_TempData 
     
    ReDim arr_ActiveColumns(1, 0) 
     
    With CreateObject("Scripting.FileSystemObject") 
     
        With .OpenTextFile(str_CsvPath, 1) 
 
            arr_CsvSource = Split(.ReadAll, vbCrLf) 
 
            'If IsMissing(arr_HeaderMask) Then 
            '    ReDim arr_TempResult(LBound(arr_HeaderMask) To UBound(arr_HeaderMask), LBound(arr_CsvSource) To UBound(arr_CsvSource)) 
            'End If 
             
            Dim i_Cursor& 
            i_LineCounter = LBound(arr_CsvSource) 
     
            For i_Cursor = LBound(arr_CsvSource) To UBound(arr_CsvSource) 
                 
                If InStr(1, arr_CsvSource(i_Cursor), "|") > 0 Then 
                    
                    If i_LineCounter = 0 Then 
                        arr_TempHeader = Split(arr_CsvSource(i_Cursor), "|") 
                         
                    'MAPPING HEADER to see which column should be copied 
                         
                        For i_HeaderCursor = LBound(arr_HeaderMask) To UBound(arr_HeaderMask) 
                            lng_ActiveColumn = -1 
                            'On Error Resume Next 
                            If z_Match(arr_TempHeader, lng_ActiveColumn, arr_HeaderMask(i_HeaderCursor)) Then 
                                 
                                arr_ActiveColumns(0, UBound(arr_ActiveColumns, 2)) = i_HeaderCursor 
                                arr_ActiveColumns(1, UBound(arr_ActiveColumns, 2)) = lng_ActiveColumn 
                                 
                                ReDim Preserve arr_ActiveColumns(1, UBound(arr_ActiveColumns, 2) + 1) 
                            End If 
                        Next 
                    End If 
 
                    'MOVE DATA TO RESULTS ARRAY 
   
                    If i_LineCounter > 0 Then 
                        arr_TempData = Split(arr_CsvSource(i_Cursor), "|") 
                         
                        For i_DataCursor = LBound(arr_ActiveColumns, 2) To UBound(arr_ActiveColumns, 2) - 1 
                            arr_TempResult(arr_ActiveColumns(0, i_DataCursor), i_LineCounter - 1) = Trim(arr_TempData(arr_ActiveColumns(1, i_DataCursor))) 
                        Next 
                    End If 
                    i_LineCounter = i_LineCounter + 1 
                     
                End If 
            Next 
        End With 
     
    End With 
 
    ReDim Preserve arr_TempResult(LBound(arr_HeaderMask) To UBound(arr_HeaderMask), LBound(arr_CsvSource) To i_LineCounter - 2) 
 
    'TRANSPOSING TO FINAL ARRAY 
    Dim x&, y& 
    Dim arr_Final() 
     
    ReDim arr_Final(LBound(arr_TempResult, 2) To UBound(arr_TempResult, 2), LBound(arr_TempResult) To UBound(arr_TempResult)) 
     
    For x = LBound(arr_TempResult) To UBound(arr_TempResult) 
        For y = LBound(arr_TempResult, 2) To UBound(arr_TempResult, 2) 
            arr_Final(y, x) = arr_TempResult(x, y) 
        Next 
    Next 
     
    arr_Result = arr_Final 
     
    'CLEARing 
    Erase arr_Final 
    Erase arr_TempResult 
    Erase arr_TempData 
 
 
Err: 
    'Resume 
     
End Function 
 
Public Sub ItemFocus(ByRef strItemAddress$) 
     If Not strItemAddress Like "*tab*/" Then z_HandleTab strItemAddress  'different tab check 
    Me.session.findById(strItemAddress).SetFocus 
End Sub 
Public Sub ItemSelect(ByRef strItemAddress$) 
    If Not strItemAddress Like "*/tab*/" Then z_HandleTab strItemAddress  'different tab check 
    Me.session.findById(strItemAddress).Select 
End Sub 
Public Sub ItemPress(ByRef strItemAddress$) 
    If Not strItemAddress Like "*/tab*/" Then z_HandleTab strItemAddress 
    Me.session.findById(strItemAddress).press 
    DoEvents 
    Wait 100 
End Sub 
Public Function ItemExist(ByRef strItemAddress$) 
    ItemExist = SapObjectExist(strItemAddress) 
End Function 
 
Public Sub ItemEdit(ByVal strAddress$) 
    session.findById(strAddress).SetFocus 'SETFOCUS 
    session.findById(Mid(strAddress, 1, 6)).sendVKey 2 
End Sub 
 
Public Function ItemEnter(Optional lng_action As value_type = saptext, Optional anyValue, Optional strPath As String = "") As String 
    
   z_HandleTab strPath 'different tab check 
     
    With session.findById(strPath) 
     
        Select Case lng_action 
            Case 1:  .Text = anyValue 
            Case 2:  .Key = anyValue 
            Case 3:  .Value = anyValue 
            Case 5:  .SetFocus: DialogLabelSimpleList 1, 0, CStr(anyValue) 
            Case 6:  .SetFocus: DialogLabelSimpleList 1, 1, CStr(anyValue) 
            Case 7:  .Selected = CBool(anyValue) 
        End Select 
    End With 
     
    Wait SleepMS 
     
End Function 
 
Public Function ItemRead(Optional lng_action As value_type = saptext, Optional strPath As String = "") As String 
    z_HandleTab strPath 'different tab check 
     
    With session.findById(strPath) 
     
        Select Case lng_action 
            Case 1: ItemRead = .Text 
            Case 2: ItemRead = .Key 
            Case 3: ItemRead = .Value 
            Case 7: ItemRead = .Selected 
        End Select 
    End With 
     
End Function 
 
 
'OLD COMPATIBILITY LAYER 
 
Private Function z_Match(ByRef arr_Source As Variant, ByRef int_ArrayIndex As Long, ByRef var_Value As Variant) 
    Dim i_Cursor As Long 
    Dim bool_Found As Boolean 
 
    If TypeName(var_Value) = "String" Then 
 
        For i_Cursor = LBound(arr_Source) To UBound(arr_Source) 
            If Trim(CStr(arr_Source(i_Cursor))) = var_Value Then bool_Found = True: Exit For 
        Next 
     
    Else 
        For i_Cursor = LBound(arr_Source) To UBound(arr_Source) 
            If arr_Source(i_Cursor) = var_Value Then bool_Found = True: Exit For 
        Next 
     
    End If 
 
    If bool_Found Then 
        z_Match = True 
        int_ArrayIndex = i_Cursor 
    Else 
        int_ArrayIndex = 0 
    End If 
 
End Function 
 
'Public Function InitiateRecords(ByRef sht_source As Worksheet, ByRef lngHeaderRow&, ByRef strKeyName$, ByRef arrDetails, Optional ByRef arrItemDetails, Optional ByVal lng_LogRun As Long = 0) As Boolean 
'    MsgBox "InitiateRecords ... Removed" 
'End Function 
'Public Sub FocusItem(ByRef strItemAddress$) 
'    MsgBox "FocusItem ... Removed" 
'End Sub 
'Public Sub SelectItem(ByRef strItemAddress$) 
'    MsgBox "SelectItem ... Removed" 
'End Sub 
'Public Sub PressItem(ByRef strItemAddress$) 
'    MsgBox "PressItem ... Removed" 
'End Sub 
'Public Function EditItem(ByVal strAddress$) 
'    MsgBox "EditItem ... Removed" 
'End Function 
'Public Function Read(ByRef lng_action As value_type, ByRef strPath$) As String 
'     MsgBox "Read ... Removed" 
'End Function 
'Public Function Enter(ByRef lng_action As value_type, ByRef anyValue, ByRef strPath$) 
'     MsgBox "Enter ... Removed" 
'End Function 
'Public Function BlackListCheck(ByRef str_record_name) As Boolean 
'    MsgBox "BlackListCheck ... Removed" 
'End Function 
 
 
 
Public Function SapChewVBS(ByVal str_Path As String) As String 
    On Error GoTo Err 
 
    'strcode = "" 'empty code 
 
    Dim lines() As String 
 
    'ChDir ThisWorkbook.Path 
    If str_Path = "skip" Then 
        ReDim lines(0) 
        lines(0) = "" 
    Else 
        Dim hf As Integer: hf = FreeFile 
        Dim i As Long 
         
        Const ForReading = 1, ForWriting = 2 
        Dim fso, MyFile, FileName 
     
        Set fso = CreateObject("Scripting.FileSystemObject") 
     
        ' Open the file for output. 
        Set MyFile = fso.OpenTextFile(str_Path, ForReading) 
         
        ' Read from the file. 
        If MyFile.AtEndOfStream Then 
            lines = Array("") 
        Else 
            lines = Split(MyFile.ReadAll, vbNewLine) 
        End If 
 
    End If 
     
    Dim c As String 'Code variable 
     
     
    c = c & vbCr & "Sub XSUBROUTINENAME(Optional ByRef t_Parameters As ListObject, Optional ByRef t_Input As ListObject, Optional ByRef t_Log As ListObject)" 
    c = c & vbCr & "" 
    c = c & vbCr & "On Error GoTo ERR" 
    c = c & vbCr & "" 
    c = c & vbCr & "'GENERATION START" 
    c = c & vbCr & "Dim SAP as clsSap" 
    c = c & vbCr & "set SAP = New clsSap" 
    c = c & vbCr & "" 
    c = c & vbCr & "with SAP" 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & "'INITIATE PARAMETERS AND LOG" 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & "Call .ExcelSetRunTables(t_Parameters, t_Log) " 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & "If Not .SapConnectToInstance(.Parameter(""System""))  then Exit Sub" 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & ".ExcelMinimize" 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & "'MAIN LOOP STARTS HERE" 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & "'Do While .RecordNext(t_Input, ""KeyColumn"") 'Record Loop Start" 
  
    c = c & vbCr & "" 
 
    For i = 15 To UBound(lines) 
 
        Select Case True 
            Case UCase(lines(i)) Like UCase("*.caretPosition = *"), lines(i) = ""  'SKIP part 
             
            Case UCase(lines(i)) Like UCase("*.setFocus") 
               c = c & vbCr & "'" & lines(i) & "'SETFOCUS" 
                 
            Case UCase(lines(i)) Like UCase("*(?wnd?0?/*/tabp?????).select") 
               c = c & vbCr & "" 
               c = c & vbCr & "'" & lines(i) & "'TABSELECT" 
               c = c & vbCr & "" 
                 
            Case UCase(lines(i)) Like UCase("*(?wnd?0?/tbar?0?/okcd?).text = *") 
               c = c & vbCr & "" 
               c = c & vbCr & vbTab & ".GoToTransaction " & """" & z_fromQuotes(lines(i), 3) & """" 
               c = c & vbCr & "" 
               'c = c & vbcr & vbTab & ".ExcelMinimize" 
                 
            Case UCase(lines(i)) Like UCase("*(?wnd?0?/tbar?0?/btn?3??).press") 
               c = c & vbCr & "" 
               c = c & vbCr & vbTab & ".PressMainButton Back_3" 
               c = c & vbCr & "" 
                 
            Case UCase(lines(i)) Like UCase("*(?wnd?0?/tbar?0?/btn?11??).press") 
               c = c & vbCr & "" 
               c = c & vbCr & vbTab & ".PressMainButton Save_11" 
               c = c & vbCr & "" 
             
            Case UCase(lines(i)) Like UCase("*(?wnd?0?/tbar?0?/btn?15??).press") 
               c = c & vbCr & "" 
               c = c & vbCr & vbTab & ".PressMainButton Exit_15" 
               c = c & vbCr & "" 
             
            Case UCase(lines(i)) Like UCase("*(?wnd?0??).sendVKey 0") 
               c = c & vbCr & "" 
               c = c & vbCr & vbTab & ".ConfirmAll" 
               c = c & vbCr & "" 
             
            Case UCase(lines(i)) Like UCase("*(?wnd?1??).sendVKey 0") 
               c = c & vbCr & "" 
               c = c & vbCr & vbTab & ".ConfirmWithPopUps" 
               c = c & vbCr & "" 
                 
            Case UCase(lines(i)) Like UCase("*.verticalScrollbar.position = *") 
                If Not ((UCase(lines(i + 1)) Like "*.verticalScrollbar.position = *") And (UCase(lines(i - 1)) Like "*.verticalScrollbar.position = *")) Then _ 
                   c = c & vbCr & vbTab & "." & lines(i) & "'SCROLLING" 
                 
            Case UCase(lines(i)) Like UCase("*(?wnd?0?/*/tbl*).*") 
                c = c & vbCr & vbTab & "." & lines(i) & "'TABLE Operation" 
                c = c & vbCr & vbTab & "." & lines(i) & "'Call .TableAdd(Array(.TableCell(saptext,""SAPColumnName"",.RecordItem(""Name""))))" 
                 
            Case UCase(lines(i)) Like UCase("*(?wnd?#?/*).key =*") 
                c = c & vbCr & vbTab & ".ItemEnter sapkey, """ & z_fromQuotes(lines(i), 3) & """, _" 
                c = c & vbCr & vbTab & vbTab & """" & z_fromQuotes(lines(i), 1) & """" 
                c = c & vbCr & "" 
 
            Case UCase(lines(i)) Like UCase("*(?wnd?#?/*).text =*") 
                 
                c = c & vbCr & vbTab & ".ItemEnter , """ & z_fromQuotes(lines(i), 3) & """, _" 
                c = c & vbCr & vbTab & vbTab & """" & z_fromQuotes(lines(i), 1) & """" 
                c = c & vbCr & "" 
 
            Case UCase(lines(i)) Like UCase("*(?wnd?#?/*).Press") 
             
                c = c & vbCr & vbTab & ".ItemPress """ & z_fromQuotes(lines(i), 1) & """" 
                c = c & vbCr & "" 
                 
            Case UCase(lines(i)) Like UCase("*(?wnd?#?/*).Select") 
             
                c = c & vbCr & vbTab & ".ItemSelect """ & z_fromQuotes(lines(i), 1) & """" 
                c = c & vbCr & "" 
                 
                 
            Case Else 
                c = c & vbCr & vbTab & "." & lines(i) 
        End Select 
    Next 
        
    c = c & vbCr & "" 
    c = c & vbCr & "LOOP_END:" 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & "'.LogAdd , , ""Ok"", ok" 
    c = c & vbCr & vbTab & ".LogToRange" 
    c = c & vbCr & vbTab & ".GoToMainPage" 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & "'Loop 'end of the Record Loop" 
    c = c & vbCr & "" 
    c = c & vbCr & vbTab & ".ExcelMaximize" 
    c = c & vbCr & "" 
    c = c & vbCr & "Exit Sub" 
    c = c & vbCr & "ERR:" 
    c = c & vbCr & vbTab & ".LogAdd , , ERR.Number & "" "" & ERR.Description, issue" 
    c = c & vbCr & vbTab & ".LogToRange" 
    c = c & vbCr & vbTab & "ERR.Raise ERR.Number" 
    c = c & vbCr & vbTab & "Resume" 
    c = c & vbCr & "End With" 
    c = c & vbCr & "End Sub" 
 
 
    SapChewVBS = c 'return code for function 
 
    Exit Function 
Err: 
 
    Err.Raise Err.Number 
    Resume 
 
End Function 
 
Private Function z_fromQuotes(ByRef strLine$, ByRef lngIndex&) As String 
 
    Dim arrTemp() As String 
     
    arrTemp = Split(strLine, """") 
     
    z_fromQuotes = arrTemp(lngIndex) 
 
End Function 
 
Public Function DialogExtractCSV( _ 
    ByVal str_ExtractName As String, _ 
    Optional ByVal str_ExtractPath As String = "", _ 
    Optional ByVal var_TriggerAction As Variant = CtrlShiftF9_45) 
 
 
        'SAVE AS CSV 
     
        'Open Dialog for Variant 
        If TypeName(var_TriggerAction) = "String" Then 
            ItemPress CStr(var_TriggerAction) 
        Else 
            PressKey CLng(var_TriggerAction) 
        End If 
         
     
        ItemSelect "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]" 
        ItemPress "wnd[1]/tbar[0]/btn[0]" 
     
        If str_ExtractPath = "\" Then str_ExtractPath = ThisWorkbook.Path & "\" 
     
        On Error Resume Next 
        Kill (str_ExtractPath & str_ExtractName) 
        On Error GoTo Err 
     
        .ItemEnter saptext, str_ExtractPath, "wnd[1]/usr/ctxtDY_PATH" 
        .ItemEnter saptext, str_ExtractName, "wnd[1]/usr/ctxtDY_FILENAME" 
        .PressKey CtrlS_11 
         
        .SapCsvCleanFix str_ExtractPath & str_ExtractName 
 
 
End Function 
 
Public Function DialogExtractXXL( _ 
    ByVal str_ExtractName As String, _ 
    Optional ByVal str_ExtractPath As String = "", _ 
    Optional ByVal var_TriggerAction As Variant = "wnd[0]/tbar[1]/btn[31]", _ 
    Optional DefaultName As String = "") 
 
    If str_ExtractPath = "" Then str_ExtractPath = ThisWorkbook.Path 
 
    'Open Dialog for Variant 
    If TypeName(var_TriggerAction) = "String" Then 
         
        If CStr(var_TriggerAction) Like "*menu*" Then 
            ItemSelect CStr(var_TriggerAction) 
        Else 
            ItemPress CStr(var_TriggerAction) 
        End If 
         
    Else 
        PressKey CLng(var_TriggerAction) 
    End If 
 
    ItemSelect "wnd[1]/usr/radRB_OTHERS" 
     
    ItemEnter sapKey, "08", "wnd[1]/usr/cmbG_LISTBOX" 
 
    ItemPress "wnd[0]/tbar[0]/btn[0]" 
 
    Confirm 
     
    ItemSelect "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]" 
     
    Confirm 
    Confirm 
     
    Dim wbk_workbook As Workbook 
     
    Call z_GetWorkbook(wbk_workbook, "Worksheet in Basis (1)") 
     
    Call wbk_workbook.SaveAs(str_ExtractPath & "\" & str_ExtractName) 
     
    Call wbk_workbook.Close(False) 
     
    Confirm 
 
End Function 
 

