Option Explicit 
 
Public Enum e_FilterType 
    e_ShowOnlyCriteriaItems = 0 
    e_HideCriteriaItems = -1 
End Enum 
 
Public Enum e_ResetFilter 
   e_ClearCurrentFilters = -1 
   e_LeaveCurrentFilters = 0 
End Enum 
 
Public Enum OutlookFolders 
    e_Inbox = 6 
    e_Outbox = 4 
    e_Drafts = 16 
    e_Deleted = 3 
    e_Calendar = 9 
End Enum 
 
Public Enum e_ClosingAction 
    e_Save 
    e_DontSave 
    e_Delete 
End Enum 
 
Public Enum e_OutlookItemTypes 
    e_Email = 43 
    e_Meeting = 26 
    e_Task = 48 
End Enum 
 
Public Enum e_OutlookFlag 
    e_Complete = 1 
    e_Marked = 2 
    e_NoFlag = 0 
End Enum 
 
Public Enum e_WordPasteAs 
    e_Original = 1 
    e_Picture = 13 
End Enum 
  
Public Enum FieldType 
    e_General 
    e_Text 
    e_Skip 
    e_DateMDY 
    e_DateDMY 
    e_DateYMD 
End Enum 
 
Public Enum e_FileNameAction 
    e_Prefix 
    e_Suffix 
    e_Replace 
End Enum 
 
Public Enum e_PickerType 
    e_FilePicker 
    e_FolderPicker 
    e_OutlookPicker 
End Enum 
 
Public Enum e_TableLoopTypes 
    e_Uniques 
    e_UniquesWithFilter 
    e_Visible 
    e_Formulas 
    e_Blanks 
    e_NonBlanks '//change to Visible 
End Enum 
 
Public Enum e_DirectionCopy 
    e_InsideTable 
    e_OutsideTable 
End Enum 
  
Public Enum e_Action 
    e_Yes = -1 
    e_No = 0 
    e_Prompt = 1 
End Enum 
   
Public Enum e_QueryParameter 
    e_adBigInt = 20  'Indicates an eight-byte signed integer (DBTYPE_I8). 
    e_adBoolean = 11  'Indicates a Boolean value (DBTYPE_BOOL). 
    e_adChar = 129  'Indicates a string value (DBTYPE_STR). 
    e_adDate = 7  'Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day. 
    e_adDouble = 5  'Indicates a double-precision floating-point value (DBTYPE_R8). 
    e_adInteger = 3  'Indicates a four-byte signed integer (DBTYPE_I4). 
    'e_adVarChar = 200  'Indicates a string value. 
End Enum 
     
'Linked Applications 
Dim app_Word As Object 'Word.Application 
Dim app_Outlook As Object 
Private fso_Object As Object 
 
Dim bool_OutlookStarted As Boolean 
 
'DataCollection 
Private cc_WorkbookPaths As Collection 
Private cc_WorkbookIndexes As Collection 
 
'TableLoopCollection 
Private cc_TableLoopHeaders As Collection 
Private cc_TableLoopCursors As Collection 
Private cc_TableLoopLevel As Collection 
 
'OutlookLoopCollection 
Private cc_OutlookFolderItemIndexes As Collection 
Private cc_OutlookFolderCursors As Collection 
Private cc_OutlookFolders As Collection 
    
'Application states 
Private lng_ApplicationEvents As Long 
Private lng_ApplicationScreen As Long 
Private lng_ApplicationCalculations  As Long 
Private lng_ApplicationAlerts As Long 
 
   
Private Const adr_Calculations As String = "C1" 
Private Const adr_RunLevel As String = "C2" 
 
Private Const adr_LocalGitPath As String = "C5" 
Private Const adr_IssueEmail As String = "C6" 
 
Private lng_RunLevel As Long 
 
Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long 
  
Public Function zATK_Version(ByRef lng_Version As Long) As Long 
    lng_Version = 3 
End Function 
  
 Sub xCleanCustomStyles(Optional ByVal bool_Prompt As Boolean = False, Optional ByRef wbk_Target As Workbook = Nothing) 
    Dim styT As Style 
    Dim intRet As Integer 
 
    If wbk_Target Is Nothing Then Set wbk_Target = ThisWorkbook 
     
 
    For Each styT In ActiveWorkbook.Styles 
        If Not styT.BuiltIn Then 
             
            If bool_Prompt Then 
                intRet = MsgBox("Delete style '" & styT.Name & "'?", vbYesNo) 
            Else 
                intRet = vbYes 
            End If 
             
         
            If intRet = vbYes Then 
                Debug.Print styT.Name & " - deleted" 
                Call styT.Delete 
            End If 
             
        End If 
    Next styT 
End Sub 
 
Private Function z_mLinkWord() As Boolean 
    Application.Statusbar = "Connecting Word" 
 
    If app_Word Is Nothing Then 
        GoTo CREATE_APP 
    Else 
        On Error GoTo OBJECT_ZOMBIE 
        Debug.Print TypeName(app_Word.Name) 
OBJECT_ZOMBIE: 
        If Err.Number <> 0 Then 
            Resume CREATE_APP 
        End If 
    End If 
     
    Application.Statusbar = False 
Exit Function 
 
CREATE_APP: 
        Set app_Word = CreateObject("Word.Application") 
        app_Word.Application.ScreenUpdating = False 
        app_Word.Application.DisplayAlerts = False 
        app_Word.Visible = False 
        Application.Statusbar = False 
End Function 
  
Private Function z_mLinkOutlook() As Boolean 
 
    z_mLinkOutlook = False 
 
    Application.Statusbar = "Connecting Outlook" 
 
    If app_Outlook Is Nothing Then 
        GoTo CREATE_APP 
    Else 
        On Error GoTo OBJECT_ZOMBIE 
        Debug.Print TypeName(app_Outlook.Name) 
OBJECT_ZOMBIE: 
        If Err.Number <> 0 Then 
            Resume CREATE_APP 
        End If 
    End If 
     
    z_mLinkOutlook = True 
    Application.Statusbar = False 
 
Exit Function 
CREATE_APP: 
    If Err.Number <> 0 Then 
        Resume CREATE_APP 
    End If 
     
    On Error GoTo CREATE_APP2: 
    Set app_Outlook = GetObject(, "Outlook.Application") 
    bool_OutlookStarted = False 
     
    z_mLinkOutlook = True 
    Application.Statusbar = False 
 
Exit Function 
CREATE_APP2: 
    If Err.Number <> 0 Then 
        Resume CREATE_APP2 
    End If 
 
    On Error GoTo FAIL: 
    Set app_Outlook = CreateObject("Outlook.Application") 
    bool_OutlookStarted = True 
     
    z_mLinkOutlook = True 
    Application.Statusbar = False 
 
Exit Function 
FAIL: 
    z_mLinkOutlook = False 
    Application.Statusbar = False 
End Function 
  
Function z_mChDirNet( _ 
    ByVal str_FilePath As String _ 
    ) As Boolean 
         
    Dim lng_Result As Long 
    lng_Result = SetCurrentDirectoryA(str_FilePath) 
    z_mChDirNet = lng_Result <> 0 
 
End Function 
 
Function MacroStart( _ 
                    Optional bool_Calculations As Boolean = False, _ 
                    Optional bool_ScreenUpdating As Boolean = False, _ 
                    Optional bool_ApplicationEvents As Boolean = False, _ 
                    Optional bool_DisplayAlerts As Boolean = False _ 
                    ) 
                     
                     
'Try to load Variables to Memory 
 
    Debug.Print "ATk-S:" & vbTab & lng_RunLevel & vbTab & "A:"; Join(Array(CLng(Application.DisplayAlerts), CLng(Application.ScreenUpdating), Application.Calculation, CLng(Application.EnableEvents)), "|" & vbTab) & vbTab & Now 
     
    'check if succesfully finished 
    If Me.Range(adr_RunLevel) <> lng_RunLevel Then 
        Application.Calculation = Me.Range(adr_Calculations) 
        lng_RunLevel = 0 'reset calculation 
        Me.Range(adr_RunLevel) = lng_RunLevel 
    End If 
     
    'Icrease Run Counter 
    lng_RunLevel = lng_RunLevel + 1 
    Me.Range(adr_RunLevel) = lng_RunLevel 
 
    'if already started skip init procedure 
    If lng_RunLevel > 1 Then Exit Function 
       
    Call Application.Run("'" & ThisWorkbook.Name & "'!LinkWorkbookTables") 
 
    'RESET ErrOR 
    If Err.Number <> 0 Then Err.Number = 0 
 
    With Application 
     
    'Save Application States 
        Me.Range(adr_Calculations) = .Calculation 
        lng_ApplicationEvents = .EnableEvents 
        lng_ApplicationScreen = .ScreenUpdating 
        lng_ApplicationAlerts = .DisplayAlerts 
 
    'Setup Application for Run 
        .Calculation = bool_Calculations 
        .ScreenUpdating = bool_ScreenUpdating 
        .EnableEvents = bool_ApplicationEvents 
        .DisplayAlerts = bool_DisplayAlerts 
     
    End With 
     
'Clear Loop Containers 
    Set cc_TableLoopHeaders = Nothing 
    Set cc_TableLoopCursors = Nothing 
    Set cc_WorkbookPaths = Nothing 
    Set cc_WorkbookIndexes = Nothing 
 
    Set cc_OutlookFolderItemIndexes = Nothing 
    Set cc_OutlookFolderCursors = Nothing 
    Set cc_OutlookFolders = Nothing 
     
     
End Function 
 
Function PickerFile( _ 
    ByRef rng_FilePath As Range, _ 
    Optional ByVal str_DialogText As String = "" _ 
    ) As Boolean 
 
    On Error GoTo err_AtkError 
 
    With Application.FileDialog(msoFileDialogFilePicker) 
 
        .AllowMultiSelect = False 
        If str_DialogText <> "" Then .Title = str_DialogText 
        If .Show Then 
            If .SelectedItems(1) <> "" Then 
                Call rng_FilePath.Hyperlinks.Add(rng_FilePath, .SelectedItems(1), , , .SelectedItems(1)) 
            End If 
        End If 
 
    End With 
 
exit_ok: 
Exit Function 
err_AtkError: 
  
    If frm_GlErr.z_Show(Err) Then: Stop: Resume 
  
End Function 
 
  
  
Function PickerFolder( _ 
    ByRef rng_FilePath As Range, _ 
    Optional ByVal str_DialogText As String = "" _ 
    ) As Boolean 
 
    On Error GoTo err_AtkError 
 
    With Application.FileDialog(msoFileDialogFolderPicker) 
 
        .AllowMultiSelect = False 
        If str_DialogText <> "" Then .Title = str_DialogText 
        If .Show Then 
            If .SelectedItems(1) <> "" Then 
                Call rng_FilePath.Hyperlinks.Add(rng_FilePath, .SelectedItems(1), , , .SelectedItems(1)) 
            End If 
        End If 
    End With 
 
  
exit_ok: 
Exit Function 
err_AtkError: 
  
     If frm_GlErr.z_Show(Err) Then: Stop: Resume 
    
End Function 
 
  
  
Function PickerOutlook( _ 
    ByRef rng_FolderPath As Range, _ 
    Optional ByVal str_DialogText As String = "" _ 
    ) As Boolean 
 
    If Not z_mLinkOutlook Then Exit Function 
 
    If str_DialogText <> "" Then MsgBox str_DialogText 
     
    Dim obj_OutlookFolder As Object 
     
    Call z_mOutlookGetFolder(obj_OutlookFolder) 
     
    rng_FolderPath.Value = CStr(obj_OutlookFolder.FolderPath) 
 
End Function 
 
Private Function z_mOutlookGetFolder( _ 
    ByRef obj_OutlookFolder As Object, _ 
    Optional ByVal FolderPath As String _ 
    ) As Boolean 
     
    z_mOutlookGetFolder = False 
     
    If Not z_mLinkOutlook Then Exit Function 
     
    Dim SearchedFolder As Object 
    Dim FoldersArray As Variant 
    Dim i As Integer 
     
    If FolderPath = "" Then 
        Set obj_OutlookFolder = app_Outlook.GetNamespace("MAPI").PickFolder 
       
    Else 
        On Error GoTo GetFolder_Error 
        If Left(FolderPath, 2) = "\\" Then 
            FolderPath = Mid(FolderPath, 3) 
        End If 
        'Convert folderpath to array 
        FoldersArray = Split(FolderPath, "\") 
         
         
        Set SearchedFolder = app_Outlook.session.Folders.Item(FoldersArray(0)) 
         
        If Not SearchedFolder Is Nothing Then 
            For i = 1 To UBound(FoldersArray, 1) 
                Dim SubFolders As Object 
                Set SubFolders = SearchedFolder.Folders 
                Set SearchedFolder = SubFolders.Item(FoldersArray(i)) 
                 
                If SearchedFolder Is Nothing Then 
                    Set obj_OutlookFolder = Nothing 
                End If 
            Next 
        End If 
        'Return the TestFolder 
        Set obj_OutlookFolder = SearchedFolder 
        z_mOutlookGetFolder = True 
        Exit Function 
 
    End If 
GetFolder_Error: 
    'Set GetFolder = Nothing 
    Exit Function 
End Function 
 
 
Function EmailMoveToFolder( _ 
    ByRef obj_OlItem As Object, _ 
    ByVal var_FolderMoveTo As Variant _ 
    ) As Boolean 
 
    If Not z_mLinkOutlook Then Exit Function 
 
    Dim str_FolderPath As String 
     
 
   ' str_FolderPath = var_FolderMoveTo 
     
    Select Case TypeName(var_FolderMoveTo) 
        Case "String" 
            str_FolderPath = var_FolderMoveTo 
        Case "Range" 
            str_FolderPath = var_FolderMoveTo.Value 
        Case Else 
    End Select 
 
    Dim folder As Object 
 
    If z_mOutlookGetFolder(folder, str_FolderPath) = False Then EmailMoveToFolder = False: Exit Function 
 
    Call obj_OlItem.move(folder) 
 
 
End Function 
 
Function EmailReply( _ 
        ByRef OlReply As Object, _ 
        ByRef obj_OlItem As Object, _ 
        Optional ByVal var_Body, _ 
        Optional ByVal str_Subject As String = "Skip", _ 
        Optional var_To, _ 
        Optional var_Cc, _ 
        Optional var_Bcc _ 
    ) As Boolean 
 
    'CHECK OUTLOOK CONNNECTION 
    If Not z_mLinkOutlook Then Exit Function 
 
    'Dim OlReply As Object 
     
    Set OlReply = obj_OlItem.Reply 
 
 
    With OlReply 
     
        On Error Resume Next 
        If Not IsMissing(var_To) Then .To = Join(z_mCovertToSimpleArray(var_To), ";") 
        If Not IsMissing(var_Cc) Then .cc = Join(z_mCovertToSimpleArray(var_Cc), ";") 
        If Not IsMissing(var_Bcc) Then .Bcc = Join(z_mCovertToSimpleArray(var_Bcc), ";") 
        If Not str_Subject = "Skip" Then .Subject = str_Subject 
        On Error GoTo 0 
         
    'HTMLBODY PART 
        Dim str_HtmlBody As String 
         
        .display 
        str_HtmlBody = .htmlBody 
         
        'Split mailBody Mail Body on part before empty line and part after empty line 
        Dim str_BodyStart As String 
        Dim str_BodyEnd As String 
        Dim str_Body As String 
               
        Dim lng_BodyStart As Long 
        Dim lng_BodyEnd As Long 
        Dim lng_BodyText As Long 
           
        Const str_EmptyLineText$ = "<p class=MsoNormal><o:p>&nbsp;</o:p></p>" 'empty line of the body 
        Const str_BodyStartTag$ = "<body " 
        Const str_BodyEndEnd$ = "</body>" 
  
        'DEFINE BODY START 
        lng_BodyStart = InStr(1, str_HtmlBody, str_BodyStartTag) ', lng_BodyStart + Len(lng_BodyStart)) 
        lng_BodyStart = InStr(lng_BodyStart, str_HtmlBody, ">") + 1 
         
        'DEFINE BODY END 
        lng_BodyEnd = InStr(1, str_HtmlBody, str_BodyEndEnd) 
         
        'GET BODY STRING 
        str_Body = Mid(str_HtmlBody, lng_BodyStart, lng_BodyEnd - lng_BodyStart) 
         
        'GET BODY START AND END STRINGS 
        str_BodyStart = Mid(str_HtmlBody, 1, lng_BodyStart - 1) 
        str_BodyEnd = Mid(str_HtmlBody, lng_BodyStart, Len(str_HtmlBody)) 
 
        'PAST CODE 
        If Not IsMissing(var_Body) Then 
            .htmlBody = str_BodyStart & "<p class=MsoNormal>" & Join(z_mCovertToSimpleArray(var_Body), "<br>") & "<br>" & "</o:p></p>" & str_BodyEnd 
        End If 
         
        Call obj_OlItem.Recipients.ResolveAll 
         
        .Save 
         
    End With 
     
    Set obj_OlItem = OlReply 
 
    EmailReply = True 
 
End Function 
 
Function EmailFolderLoop(ByRef obj_OlItem As Object, ByVal var_FolderInput As Variant, Optional ByVal var_Criteria As Variant, Optional ByVal var_Attachments As Variant) As Boolean 
 
    Dim lng_OlItemType As Long 
    Dim folder As Object 
    Dim folderItems As Object 
    Dim lng_ItemCursor As Long 
    Dim str_FolderPath As String 
     
     
    str_FolderPath = var_FolderInput 
    lng_OlItemType = e_Email 
     
    Select Case TypeName(var_FolderInput) 
        Case "String" 
            str_FolderPath = var_FolderInput 
        Case "Range" 
            str_FolderPath = var_FolderInput.Value 
        Case Else 
    End Select 
     
    Dim arr_Criteria 
    Dim arr_Attachements 
     
    If Not IsMissing(var_Criteria) Then 
        arr_Criteria = z_mCovertToSimpleArray(var_Criteria) 
    End If 
     
    If Not IsMissing(var_Attachments) Then 
        arr_Attachements = z_mCovertToSimpleArray(var_Attachments) 
    End If 
     
    Dim cc_OutlookFolderItemIndexesTemp As Collection 
     
    If z_mKeyExists(cc_OutlookFolders, str_FolderPath) Then 
 
        Set folderItems = cc_OutlookFolders.Item(str_FolderPath) 
         
        'With folder 
     
            lng_ItemCursor = cc_OutlookFolderCursors.Item(str_FolderPath) 
             
            Set cc_OutlookFolderItemIndexesTemp = cc_OutlookFolderItemIndexes.Item(str_FolderPath) 
 
            If lng_ItemCursor > cc_OutlookFolderItemIndexesTemp.Count Then 
                Set obj_OlItem = Nothing 
                 
                Call cc_OutlookFolderCursors.Remove(str_FolderPath) 'reset 
                Call cc_OutlookFolderCursors.Add(1, str_FolderPath) 
                EmailFolderLoop = False 
                Exit Function 
            Else 
             
                Set obj_OlItem = folderItems(cc_OutlookFolderItemIndexesTemp.Item(lng_ItemCursor)) 
             
                Call cc_OutlookFolderCursors.Remove(str_FolderPath) 
                Call cc_OutlookFolderCursors.Add(lng_ItemCursor + 1, str_FolderPath) 
                EmailFolderLoop = True 
            End If 
             
        'End With 
     
    Else 
     
        If z_mOutlookGetFolder(folder, str_FolderPath) = False Then EmailFolderLoop = False: Exit Function 
 
        
 
        Set folderItems = folder.items 
 
 
    'FOLDER FILTERING 
        'Attachment Restrict 
        If Not IsMissing(var_Attachments) Then 
            Set folderItems = folderItems.Restrict("[Attachment] > 0") 
        End If 
 
        'Criteria Restrict 
        If Not IsMissing(var_Criteria) Then 
            Dim lng_CriteriaLoop As Long 
         
            'Criteria restrict 
            For lng_CriteriaLoop = LBound(arr_Criteria) To UBound(arr_Criteria) 
                Set folderItems = folderItems.Restrict(arr_Criteria(lng_CriteriaLoop)) 
            Next 
        End If 
  
    'POPULATING COLLECTION 
        'With folderItems 
         
            Dim lng_ItemsCount As Long 
            Dim bool_FirstPicked As Boolean 
     
            Set cc_OutlookFolderItemIndexesTemp = New Collection 
        
            lng_ItemsCount = folderItems.Count 
            bool_FirstPicked = False 
     
            If lng_ItemsCount > 0 Then 'on empty check 
     
                'Attachment Check for performace completly separated from the loop 
                If IsMissing(var_Attachments) Then 
                     
                    'without attachment 
                    For lng_ItemCursor = lng_ItemsCount To 1 Step -1 
                         
                        If folderItems(lng_ItemCursor).Class = lng_OlItemType Then 
                            If bool_FirstPicked = False Then 
                                Set obj_OlItem = folderItems(lng_ItemCursor) 
                                bool_FirstPicked = True 
                            Else 
                                Call cc_OutlookFolderItemIndexesTemp.Add(lng_ItemCursor) 
                            End If 
                        End If 
                 
                    Next 
                     
                Else 
                    'with attachmetns 
                    For lng_ItemCursor = lng_ItemsCount To 1 Step -1 
                         
                        If folderItems(lng_ItemCursor).Class = lng_OlItemType Then 
                             
                            Dim bool_FilenameMatch As Boolean 
                            Dim obj_Attachment As Object 
                            Dim lng_AttachmentCursor As Long 
                             
                            bool_FilenameMatch = False 
                             
                            'attachment lookup 
                            For Each obj_Attachment In folderItems(lng_ItemCursor).Attachments 
                                 
                                'check if criteria contain criteria attachments 
                                For lng_AttachmentCursor = LBound(arr_Attachements) To UBound(arr_Attachements) 
                                    If obj_Attachment.DisplayName Like arr_Attachements(lng_AttachmentCursor) Then 
                                        bool_FilenameMatch = True 
                                        Exit For 
                                    End If 
                                Next 
 
                            Next obj_Attachment 
 
                            If bool_FilenameMatch Then 
                                If bool_FirstPicked = False Then 
                                    Set obj_OlItem = folderItems(lng_ItemCursor) 
                                    bool_FirstPicked = True 
                                Else 
                                    Call cc_OutlookFolderItemIndexesTemp.Add(lng_ItemCursor) 
                                End If 
                            End If 
                         
                        End If 
                 
                    Next 
                 
                End If 
                 
                'Items 
                If cc_OutlookFolders Is Nothing Then Set cc_OutlookFolders = New Collection 
                If cc_OutlookFolderItemIndexes Is Nothing Then Set cc_OutlookFolderItemIndexes = New Collection 
                If cc_OutlookFolderCursors Is Nothing Then Set cc_OutlookFolderCursors = New Collection 
                 
                Call cc_OutlookFolders.Add(folderItems, str_FolderPath) 
                 
                'Only one item dummy collection 
                If cc_OutlookFolderItemIndexesTemp.Count = 0 And bool_FirstPicked = True Then 
                    Call cc_OutlookFolderItemIndexesTemp.Add(1) 'enter dummy collection 
                    Call cc_OutlookFolderItemIndexes.Add(cc_OutlookFolderItemIndexesTemp, str_FolderPath) 
                    Call cc_OutlookFolderCursors.Add(2, str_FolderPath) 
                 
                'more then on item found 
                ElseIf cc_OutlookFolderItemIndexesTemp.Count > 0 Then 
                    Call cc_OutlookFolderItemIndexes.Add(cc_OutlookFolderItemIndexesTemp, str_FolderPath) 
                    Call cc_OutlookFolderCursors.Add(1, str_FolderPath) 
                 
                End If 
                 
                EmailFolderLoop = bool_FirstPicked 'return true only if at least one corresponding match 
            End If 
 
            'End With 
 
    End If 
 
End Function 
 
Public Sub DocumentOpen(ByRef obj_WordDoc As Object, ByVal str_FileName As String, ByVal str_DirPath As String) 
 
    Call z_mLinkWord 
 
    Application.Statusbar = "Openning Document ... " & str_DirPath & "\" & str_FileName 
 
    Set obj_WordDoc = app_Word.Documents.Open(str_DirPath & "\" & str_FileName) 
 
    Application.Statusbar = False 
 
End Sub 
 
Public Sub DocumentToPDF(ByVal obj_WordDoc As Object, ByVal str_FileName As String, ByVal str_DirPath As String) 
 
    Call z_mLinkWord 
 
    'save as PDF' 
    obj_WordDoc.ExportAsFixedFormat OutputFileName:= _ 
        str_DirPath & str_FileName, _ 
        ExportFormat:=17, OpenAfterExport:=False, OptimizeFor:= _ 
        0, Range:=0, From:=1, To:=1, _ 
        Item:=7, IncludeDocProps:=False, KeepIRM:=True, _ 
        CreateBookmarks:=0, DocStructureTags:=True, _ 
        BitmapMissingFonts:=True, UseISO19005_1:=False 
 
End Sub 
 
Public Sub DocumentSaveAs(ByVal obj_WordDoc As Object, ByVal str_FileName As String, ByVal str_DirPath As String) 
 
    Call z_mLinkWord 
 
    Application.Statusbar = "Openning Saving Document as ... " & str_DirPath & "\" & str_FileName 
 
    obj_WordDoc.SaveAs str_DirPath & "\" & str_FileName 
 
    Application.Statusbar = False 
 
End Sub 
 
Public Sub DocumentClose(ByVal obj_WordDoc As Object, ByVal bool_SaveDocument As Boolean) 
 
    Call z_mLinkWord 
         
    Application.Statusbar = "Closing Document ... " & obj_WordDoc.Name & IIf(bool_SaveDocument, "", " Without Saving") 
     
    Call obj_WordDoc.Close(bool_SaveDocument) 
 
    Application.Statusbar = False 
 
End Sub 
 
Public Function DocumentReplaceText( _ 
        ByRef obj_WordDoc As Object, _ 
        Optional ByRef t_SingleItems As ListObject, _ 
        Optional ByRef t_Table0 As ListObject, _ 
        Optional ByRef t_Table1 As ListObject, _ 
        Optional ByRef t_Table2 As ListObject, _ 
        Optional ByRef t_Table3 As ListObject, _ 
        Optional ByRef t_Table4 As ListObject, _ 
        Optional ByRef t_Table5 As ListObject, _ 
        Optional ByRef t_Table6 As ListObject, _ 
        Optional ByRef t_Table7 As ListObject, _ 
        Optional ByRef t_Table8 As ListObject, _ 
        Optional ByRef t_Table9 As ListObject _ 
    ) 
         
 
    t_SingleItems.Range.Calculate 
    
    'FILL BODY ITEMS 
     
    Dim rng_Cursor As Range 
     
    With obj_WordDoc.Range.Find 
    'Document single items 
        If Not t_SingleItems Is Nothing Then 
            For Each rng_Cursor In t_SingleItems.HeaderRowRange 
                .Execute rng_Cursor.Value, True, True, False, , , , , , rng_Cursor.Offset(1).Value, 2 
            Next 
        End If 
    End With 
 
 
'IF LIST EXIST ADD LIST ITEMS 
    Dim obj_Table As Object 
    Dim lng_TableCorsor As Long 
 
    For lng_TableCorsor = 1 To obj_WordDoc.tables.Count 
 
        Set obj_Table = obj_WordDoc.tables(lng_TableCorsor) 
 
        Select Case obj_Table.Title 
            Case 0: If Not t_Table0 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table0, 0) 
            Case 1: If Not t_Table1 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table1, 1) 
            Case 2: If Not t_Table2 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table2, 2) 
            Case 3: If Not t_Table3 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table3, 3) 
            Case 4: If Not t_Table4 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table4, 4) 
            Case 5: If Not t_Table5 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table5, 5) 
            Case 6: If Not t_Table6 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table6, 6) 
            Case 7: If Not t_Table7 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table7, 7) 
            Case 8: If Not t_Table8 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table8, 8) 
            Case 9: If Not t_Table9 Is Nothing Then Call DocumentTableLoad(obj_Table, t_Table9, 9) 
        End Select 
     
    Next 
 
 
End Function 
 
Public Function DocumentTableLoad(ByVal obj_Table As Object, ByVal t_Table As ListObject, ByVal lng_TableIndex As Long) 
     
    Dim lng_RowsCount As Long 
    Dim lng_RowsCounter As Long 
          
    t_Table.Range.Calculate 
          
    t_Table.MoveLast 
    lng_RowsCount = t_Table.DataBodyRange.Rows.Count 
    t_Table.MoveFirst 
     
    With obj_Table 
 
            .Rows(1).Select 
            app_Word.Selection.InsertRowsBelow (lng_RowsCount - 1) 
            .Rows(1).Select 
             app_Word.Selection.Copy 
             app_Word.Selection.MoveDown Unit:=5, Count:=1 '5=wdLine 
             app_Word.Selection.MoveDown Unit:=5, Count:=(lng_RowsCount - 2), Extend:=1  '5=wdLine, 1= wdExtend 
             app_Word.Selection.Paste 
             
            lng_RowsCounter = 1 
            Dim rng_Cursor As Range 
            Do Until lng_RowsCounter <> lng_RowsCount 
                 
                With .Rows(lng_RowsCounter).Range.Find 
                    For Each rng_Cursor In t_Table.HeaderRowRange 
                        .Execute "[" & rng_Cursor.Value & "]" & lng_TableIndex, True, True, False, , , , , , rng_Cursor.Offset(lng_RowsCounter - 1).Value, 2 
                    Next 
                End With 
 
                t_Table.MoveNext 
                lng_RowsCounter = lng_RowsCounter + 1 
     
            Loop 
 
    End With 
     
End Function 
 
Function MacroFinish(Optional bool_ForceFinish As Boolean = False) 
 
    DoEvents 
  
    Debug.Print "ATk-E:" & vbTab & lng_RunLevel & vbTab & "A:"; Join(Array(lng_ApplicationAlerts, lng_ApplicationScreen, lng_ApplicationCalculations, lng_ApplicationEvents), "|") & vbTab & Now 
  
    If lng_RunLevel <> Me.Range(adr_RunLevel) Or bool_ForceFinish Then 'not same then error reset 
        'Error Recover 
        lng_RunLevel = 1 
        Me.Range(adr_RunLevel) = lng_RunLevel 
    End If 
 
    lng_RunLevel = lng_RunLevel - 1 
    Me.Range(adr_RunLevel) = lng_RunLevel 
 
    If lng_RunLevel > 0 Then Exit Function 
 
'Check if in words is no anything open if yes show word 
    If Not app_Word Is Nothing Then 
        If app_Word.Documents.Count = 0 Then 
            app_Word.Application.ScreenUpdating = True 
            app_Word.Application.Quit 0 'quit without saving 
            Set app_Word = Nothing 
        Else 
            app_Word.Visible = False 
        End If 
    End If 
 
'Check if outlook was started then close 
    If Not app_Outlook Is Nothing And bool_OutlookStarted Then 
        app_Outlook.Quit 
    End If 
 
'Check if FilesSystem is not running 
    If Not fso_Object Is Nothing Then 
        Set fso_Object = Nothing 
    End If 
  
'Loading original data to Excel 
    With Application 
        .Calculation = Me.Range(adr_Calculations) 
        .ScreenUpdating = lng_ApplicationScreen 
        .EnableEvents = lng_ApplicationEvents 
        .DisplayAlerts = lng_ApplicationAlerts 
        .Statusbar = False 
    End With 
     
'Clear Loop Containers to freeup memory 
    Set cc_TableLoopHeaders = Nothing 
    Set cc_TableLoopCursors = Nothing 
    Set cc_WorkbookPaths = Nothing 
    Set cc_WorkbookIndexes = Nothing 
 
    Set cc_OutlookFolderItemIndexes = Nothing 
    Set cc_OutlookFolderCursors = Nothing 
    Set cc_OutlookFolders = Nothing 
 
End Function 
 
Function TableClear(ByVal tbl_Object As ListObject) 
 
    On Error GoTo err_AtkError 
    Application.Statusbar = "Clearing Table " & tbl_Object.Name 
 
    'EXECUTE 
 
    Dim rng_ClearRange As Range 
     
'check if there is rows to remove 
    If tbl_Object.Range.Rows.Count > 2 Then 
         
        Set rng_ClearRange = tbl_Object.Range.Worksheet.Range(tbl_Object.Range.Offset(2).Resize(tbl_Object.Range.Rows.Count - 2).Address) 
     
    'resize table 
        Call tbl_Object.Resize(tbl_Object.HeaderRowRange.Resize(2)) 
         
        'clear the rest 
        rng_ClearRange.Clear 'clear with formulas 
         
    End If 
 
'reset autofilter 
    tbl_Object.Range.AutoFilter 
    tbl_Object.ShowAutoFilter = True 
 
'clear first input row table data 
    If tbl_Object.Range.Columns.Count = 1 Then 'avoid whole table removal 
         tbl_Object.HeaderRowRange.Offset(1).ClearContents 'clear just data 
    Else 
        If z_ContainsCellType(tbl_Object.HeaderRowRange.Offset(1), xlCellTypeConstants) Then 
         tbl_Object.HeaderRowRange.Offset(1).SpecialCells(xlCellTypeConstants).ClearContents 'clear just data 
            tbl_Object.Range.Calculate 
    End If 
    End If 
     
    'EXIT 
     
exit_ok: 
Exit Function 
err_AtkError: 
 
    Debug.Print Err.Number & vbTab & Err.Description 
 
    If Err.Number = 1004 Then Resume exit_ok 'skip if this error nothing to delete 
    If frm_GlErr.z_Show(Err) Then Resume 'general err_handling 
    
End Function 
 
Function TableClearAll(ByVal tbl_Object As ListObject) 
 
    Dim rng_TempTable As Range 
     
 
    On Error GoTo Err 
    Set rng_TempTable = tbl_Object.DataBodyRange 
 
    'reset autofilter 
    Call tbl_Object.Range.AutoFilter 
    tbl_Object.ShowAutoFilter = True 
 
    Call tbl_Object.Resize(tbl_Object.HeaderRowRange.Resize(2)) 
 
    rng_TempTable.ClearContents 
         
Err: 
End Function 
 
Function TableClearVisible(ByVal tbl_Object As ListObject) 
 
    Dim rng_BodyToDelete As Range 
    'deleting 
    On Error Resume Next 
    Set rng_BodyToDelete = tbl_Object.DataBodyRange.SpecialCells(xlCellTypeVisible) 
    On Error GoTo 0 
     
    If Not rng_BodyToDelete Is Nothing Then 
      If rng_BodyToDelete.Address = tbl_Object.DataBodyRange.Address Then 
        Call TableClear(tbl_Object) 
      Else 
         
        rng_BodyToDelete.Delete 
      End If 
    End If 
     
    'reset autofilter 
    tbl_Object.Range.AutoFilter 
    tbl_Object.ShowAutoFilter = True 
 
End Function 
 
Private Function z_mCovertToSimpleArray( _ 
                                            ByVal arr_Criteria As Variant, _ 
                                            Optional ByVal bool_ReturnDateVersion As Boolean = False) As Variant 
 
    Dim arr_TempCriteria 
 
    'If Entered table then change to Range 
    If TypeName(arr_Criteria) = "ListObject" Then 
        Dim rng_Temp As Range 
         
        On Error Resume Next 
        Set rng_Temp = arr_Criteria.DataBodyRange 
        On Error GoTo 0 
         
        If Not rng_Temp Is Nothing Then Set arr_Criteria = rng_Temp 
         
    End If 
     
    'New Range Item 
    Select Case TypeName(arr_Criteria) 
        Case "Range" 
         
            'merge area check 
            If arr_Criteria.MergeCells Then 
               ' ReDim arr_TempCriteria(0) 
                 
                If arr_Criteria.Worksheet.Range(Split(arr_Criteria.Address, ":")(0)).Value = Empty Then 
                    arr_TempCriteria = CVar(Array(vbNullString)) 
                Else 
                    arr_TempCriteria = Array(arr_Criteria.Worksheet.Range(Split(arr_Criteria.Address, ":")(0)).Value) 
                End If 
                 
                GoTo Exit_Function: 
            Else 
             
                If arr_Criteria.Count > 1 Then 
                    arr_TempCriteria = arr_Criteria.Value 
                    arr_TempCriteria = z_mCovertToSimpleArray(arr_TempCriteria) 
                Else 
                    arr_Criteria.Calculate 
                    arr_TempCriteria = Array(arr_Criteria.Value) 
                End If 
     
            End If 
     
        Case "Variant()" 
            If z_mIs2dArray(arr_Criteria) Then 
                Dim r_Cursor As Long 
                Dim c_Cursor As Long 
                Dim lng_ItemCounter As Long 
                 
                lng_ItemCounter = 0 
                'Dim arr_TempCriteria1D() 
                ReDim arr_TempCriteria(0) 
                ReDim arr_TempCriteria((UBound(arr_Criteria) * UBound(arr_Criteria, 2)) - 1) 
                 
                For r_Cursor = LBound(arr_Criteria) To UBound(arr_Criteria) 
                    For c_Cursor = LBound(arr_Criteria, 2) To UBound(arr_Criteria, 2) 
     
                        arr_TempCriteria(lng_ItemCounter) = arr_Criteria(r_Cursor, c_Cursor) 
                        lng_ItemCounter = lng_ItemCounter + 1 
                     
                    Next 
                Next 
                 
            Else 
                z_mCovertToSimpleArray = arr_Criteria 
                Exit Function 
            End If 
     
        Case "String", "Long", "Integer" 
            arr_TempCriteria = Array(arr_Criteria) 
        Case "Date" 
            arr_TempCriteria = Array(arr_Criteria) 
        Case "Boolean" 
            arr_TempCriteria = Array(arr_Criteria) 
        Case Else 
             
    End Select 
 
    If bool_ReturnDateVersion Then 
     
    Dim lng_Date As Long 
    'Criteria Range To Date 
        For r_Cursor = LBound(arr_TempCriteria) To UBound(arr_TempCriteria) 
            lng_Date = CLng(Fix(CCur(arr_TempCriteria(r_Cursor)))) 
            arr_TempCriteria(r_Cursor) = CStr(Month(lng_Date)) & "/" & CStr(Day(lng_Date)) & "/" & CStr(Year(lng_Date)) 
        Next 
 
    End If 
 
    z_mCovertToSimpleArray = arr_TempCriteria 
 
Exit_Function: 
End Function 
' 
'Private Function z_mCovertToSimpleArray(ByRef arr_Criteria As Variant, Optional ByVal bool_ReturnDateVersion As Boolean = False) 
' 
'Dim arr_TempCriteria() 
' 
'Select Case TypeName(arr_Criteria) 
'    Case "Range" 
'        arr_TempCriteria = arr_Criteria.Value 
' 
'        If arr_Criteria.Count > 1 Then 
'            Call z_mCovertToSimpleArray(arr_TempCriteria) 
'        Else 
'            arr_TempCriteria = Array(arr_TempCriteria) 
'        End If 
' 
'    Case "Variant()" 
'        If z_mIs2dArray(arr_Criteria) Then 
'            Dim r_Cursor As Long 
'            Dim c_Cursor As Long 
'            Dim lng_ItemCounter As Long 
' 
'            lng_ItemCounter = 0 
'            ReDim arr_TempCriteria(0) 
'            ReDim arr_TempCriteria((UBound(arr_Criteria) * UBound(arr_Criteria, 2)) - 1) 
' 
'            For r_Cursor = LBound(arr_Criteria) To UBound(arr_Criteria) 
'                For c_Cursor = LBound(arr_Criteria, 2) To UBound(arr_Criteria, 2) 
' 
'                    arr_TempCriteria(lng_ItemCounter) = arr_Criteria(r_Cursor, c_Cursor) 
'                    lng_ItemCounter = lng_ItemCounter + 1 
' 
'                Next 
'            Next 
'        Else 
'            Exit Function 
'        End If 
' 
'    Case "String", "Long", "Integer" 
'        arr_TempCriteria = Array(arr_Criteria) 
'    Case "Date" 
'        arr_TempCriteria = Array(arr_Criteria) 
'    Case "Boolean" 
'    Case Else 
' 
'End Select 
' 
'arr_Criteria = arr_TempCriteria 
' 
'End Function 
 
Private Sub z_mVariantValue(ByRef ChangedVariant As Variant, ByVal ChangingVariant) 
    ChangedVariant = ChangingVariant 
End Sub 
 
Private Function z_mIs2dArray(ByVal var_Array) As Boolean 
     
    z_mIs2dArray = False 
     
    On Error GoTo Err: 
 
    Dim lng_2d As Long 
     
    lng_2d = UBound(var_Array, 2) 
 
    z_mIs2dArray = True 
 
    Exit Function 
 
Err: 
 
End Function 
 
Function TableAutofilterDate( _ 
                        ByRef tbl_Object As ListObject, _ 
                        Optional ByRef str_ColumnName As String = "", _ 
                        Optional ByRef arr_CriteriaInput As Variant, _ 
    Optional ByRef InvertedAutoFilter As e_FilterType = e_ShowOnlyCriteriaItems, _ 
    Optional ByRef RemoveAutoFilter As e_ResetFilter = e_LeaveCurrentFilters _ 
                        ) 
     
    On Error GoTo Err: 
     
'PREPARE TABLE 
    Dim str_TableName$ 
    Dim lng_ColumnIndex& 
    Dim i_Values& 
    Dim lng_Date& 
    Dim arr_Criteria() 
     
    str_TableName = tbl_Object.Name 
     
'reset autofilter 
    If str_ColumnName = "" Then 
        tbl_Object.Range.AutoFilter 
        tbl_Object.ShowAutoFilter = True 
        Exit Function 
    End If 
     
'find column index of the name of column 
    lng_ColumnIndex = tbl_Object.Range.Rows(1).Find(str_ColumnName).Column - (tbl_Object.Range.Column - 1) 
     
'autofilter before filtering 
    If RemoveAutoFilter Then tbl_Object.Range.AutoFilter
     
'check if autofilter is not turned of 
    If Not tbl_Object.ShowAutoFilter Then tbl_Object.ShowAutoFilter = True 
     
'PREPARE CRITERIA 
'check if filter list is not array 
     
    'If TypeName(arr_Criteria) = "Range" Then arr_Criteria = RangeToArray(arr_Criteria) 
     
    arr_Criteria = z_mCovertToSimpleArray(arr_CriteriaInput, True) 
     
'    For i_Values = 0 To UBound(arr_Criteria) 
'        lng_Date = CLng(Fix(CCur(arr_Criteria(i_Values)))) 
'        arr_Criteria(i_Values) = CStr(Month(lng_Date) & "/" & Day(lng_Date) & "/" & Year(lng_Date)) 
'    Next 
     
'check if there is not required negative autofilter 
    If InvertedAutoFilter Then 
        Dim arr_ColumnData 
        arr_ColumnData = WorksheetFunction.Transpose(tbl_Object.Range.Worksheet.Range(str_TableName & "[" & str_ColumnName & "]")) 
         
        For i_Values = 1 To UBound(arr_ColumnData) 
            lng_Date = CLng(Fix(CCur(arr_ColumnData(i_Values)))) 
            arr_ColumnData(i_Values) = CStr(Month(lng_Date) & "/" & Day(lng_Date) & "/" & Year(lng_Date)) 
        Next 
         
        Call z_mArrayFilterUnique(arr_ColumnData, arr_Criteria) 
    End If 
     
'DateAutofilter Criteria Array 
    Dim arr_DateCritieria 
    Dim i& 
     
'Date Criteria 
    ReDim arr_DateCritieria((UBound(arr_Criteria) * 2) + 1) 
    For i = LBound(arr_Criteria) To UBound(arr_Criteria) 
        'lng_Date = Fix(CLng(arr_Criteria(i))) 
        arr_DateCritieria(i * 2) = 2 'criteria scope taken 0-Year,1-Month,2-Day,3-Hour,4-Minute 
        arr_DateCritieria(i * 2 + 1) = arr_Criteria(i) 
    Next 
     
'apply autofilter 
    tbl_Object.Range.AutoFilter _ 
        Field:=lng_ColumnIndex, _ 
        Criteria2:=arr_DateCritieria, _ 
        Operator:=xlFilterValues 
    Exit Function 
     
Err: 
    Debug.Print Err.Number & " " & Err.Description 
    Debug.Assert False 
    Resume 
 
End Function 
 
Function TableAutofilter( _ 
                    ByVal tbl_Object As ListObject, _ 
                    Optional ByVal str_ColumnName As String = "", _ 
                    Optional ByVal arr_CriteriaInput As Variant, _ 
                    Optional ByVal InvertedAutoFilter As e_FilterType = e_ShowOnlyCriteriaItems, _ 
                    Optional ByVal RemoveAutoFilter As e_ResetFilter = e_LeaveCurrentFilters _ 
                    ) 
                     
    Application.Statusbar = "Filtering table: " & tbl_Object.Name & " in Column: " & str_ColumnName 
                     
     
    Dim str_TableName$ 
    Dim lng_ColumnIndex& 
 
    str_TableName = tbl_Object.Name 
     
'check do - reset autofilter 
    If str_ColumnName = "" Then 
        tbl_Object.Range.AutoFilter 
        tbl_Object.ShowAutoFilter = True 
        Exit Function 
    End If 
 
     
'find column index of the name of column 
    lng_ColumnIndex = tbl_Object.HeaderRowRange.Find(str_ColumnName).Column - (tbl_Object.Range.Column - 1) 
     
'check do - reset filter in column 
    If str_ColumnName <> "" And IsMissing(arr_CriteriaInput) Then 
        tbl_Object.Range.AutoFilter lng_ColumnIndex 
        Exit Function 
    End If 
     
'autofilter before filtering 
    If RemoveAutoFilter Then tbl_Object.Range.AutoFilter 
     
'check if autofilter is not turned of 
    If tbl_Object.ShowAutoFilter Then tbl_Object.ShowAutoFilter = True 
     
'recalculate column data 
    Call z_mRangeRecalculate(tbl_Object.Range.Worksheet.Range(str_TableName & "[" & str_ColumnName & "]")) 
     
'check if filter list is not array 
    Dim arr_Criteria 
    arr_Criteria = z_mCovertToSimpleArray(arr_CriteriaInput) 
     
'check if there is not required negative autofilter 
    If InvertedAutoFilter Then 
        Dim arr_ColumnData 
        arr_ColumnData = WorksheetFunction.Transpose(tbl_Object.Range.Worksheet.Range(str_TableName & "[" & str_ColumnName & "]")) 
         
        arr_ColumnData = z_mCovertToSimpleArray(arr_ColumnData) 
         
        If Not z_mArrayFilterUnique(arr_ColumnData, arr_Criteria) Then 
            'Exit Function 
        End If 
    Else 
       'convert integer and long numbers to strings 
        Dim i& 
        For i = LBound(arr_Criteria) To UBound(arr_Criteria) 
            arr_Criteria(i) = CStr(arr_Criteria(i)) 
        Next 
    End If 
     
    'SPECIAL CHARACTER SETTINGS 
    If arr_Criteria(0) = "<>" And UBound(arr_Criteria) = 0 Then 
        tbl_Object.Range.AutoFilter _ 
            Field:=lng_ColumnIndex, _ 
            Criteria1:=arr_Criteria(0) 
             
        GoTo Exit_Function 
    End If 
     
    'apply autofilter 
    tbl_Object.Range.AutoFilter _ 
        Field:=lng_ColumnIndex, _ 
        Criteria1:=arr_Criteria, _ 
        Operator:=xlFilterValues 
 
     
    Application.Statusbar = False 
     
Exit_Function: 
 
End Function 
 
 
Private Function z_mRangeRecalculate(ByRef rng_Range As Range) As Boolean 
     
    z_mRangeRecalculate = False 
     
    Dim rng_Formulas As Range 
     
On Error Resume Next 
    Set rng_Formulas = rng_Range.SpecialCells(xlCellTypeFormulas) 
On Error GoTo 0 
 
    If Not rng_Formulas Is Nothing Then 
        rng_Formulas.Calculate 
        z_mRangeRecalculate = True 
    End If 
 
End Function 
 
Function RangeAutofilter( _ 
                ByVal rng_SourceRange As Range, _ 
                Optional ByVal lng_HeaderRow As Long = 1, _ 
                Optional ByVal str_ColumnName As String = "", _ 
                Optional ByVal arr_Criteria As Variant, _ 
                Optional ByVal InvertedAutoFilter As e_FilterType = e_ShowOnlyCriteriaItems, _ 
                Optional ByVal RemoveAutoFilter As e_ResetFilter = e_LeaveCurrentFilters _ 
                ) 
     
 '   Dim str_TableName$ 
    Dim lng_ColumnIndex& 
    Dim lng_DataRowsCount& 
     
    Set rng_SourceRange = rng_SourceRange.CurrentRegion 
    lng_DataRowsCount = rng_SourceRange.Rows.Count - lng_HeaderRow 'calculate number of rows below header row 
    lng_HeaderRow = lng_HeaderRow - 1 'remove one row for ofset 
 
'reset autofilter 
    If str_ColumnName = "" Then 
        rng_SourceRange.AutoFilter 
        rng_SourceRange.ShowAutoFilter = True 
        Exit Function 
    End If 
     
'find column index of the name of column 
    lng_ColumnIndex = rng_SourceRange.Resize(1).Offset(lng_HeaderRow).Find(str_ColumnName).Column - (rng_SourceRange.Resize(1, 1).Column - 1) 
     
'autofilter before filtering 
    If RemoveAutoFilter Then rng_SourceRange.AutoFilter 
 
'check if filter list is not array 
    If TypeName(arr_Criteria) = "Range" Then arr_Criteria = RangeToArray(arr_Criteria) 
     
'check if there is not required negative autofilter 
    If InvertedAutoFilter Then 
        Dim arr_ColumnData 
        arr_ColumnData = WorksheetFunction.Transpose(rng_SourceRange.Resize(lng_DataRowsCount, 1).Offset(lng_HeaderRow, lng_ColumnIndex)) 
         
        If Not z_mArrayFilterUnique(arr_ColumnData, arr_Criteria) Then 
            'Exit Function 
        End If 
    Else 
       'convert integer and long numbers to strings 
        Dim i& 
        For i = LBound(arr_Criteria) To UBound(arr_Criteria) 
            arr_Criteria(i) = CStr(arr_Criteria(i)) 
        Next 
    End If 
     
 
    If UBound(arr_Criteria) = 0 Then 
        rng_SourceRange.Resize(lng_DataRowsCount + 1).Offset(lng_HeaderRow).AutoFilter _ 
            Field:=lng_ColumnIndex, _ 
            Criteria1:=arr_Criteria(0) 
    Else 
        'apply autofilter 
        rng_SourceRange.Resize(lng_DataRowsCount + 1).Offset(lng_HeaderRow).AutoFilter _ 
            Field:=lng_ColumnIndex, _ 
            Criteria1:=arr_Criteria, _ 
            Operator:=xlFilterValues 
    End If 
 
End Function 
 
Function RangeAutofilterDate( _ 
                            ByVal rng_SourceRange As Range, _ 
                            Optional ByVal lng_HeaderRow& = 1, _ 
                            Optional ByVal str_ColumnName$ = "", _ 
                            Optional ByVal arr_Criteria, _ 
                            Optional ByVal InvertedAutoFilter As e_FilterType = e_ShowOnlyCriteriaItems, _ 
                            Optional ByVal RemoveAutoFilter As e_ResetFilter = e_LeaveCurrentFilters _ 
                            ) 
     
 '   Dim str_TableName$ 
    Dim lng_ColumnIndex& 
    Dim lng_DataRowsCount& 
     
    Set rng_SourceRange = rng_SourceRange.CurrentRegion 
    lng_DataRowsCount = rng_SourceRange.Rows.Count - lng_HeaderRow 'calculate number of rows below header row 
    lng_HeaderRow = lng_HeaderRow - 1 'remove one row for ofset 
 
'reset autofilter 
    If str_ColumnName = "" Then 
        rng_SourceRange.AutoFilter 
        rng_SourceRange.ShowAutoFilter = True 
        Exit Function 
    End If 
     
'find column index of the name of column 
    lng_ColumnIndex = rng_SourceRange.Resize(1).Offset(lng_HeaderRow).Find(str_ColumnName).Column - (rng_SourceRange.Resize(1, 1).Column - 1) 
     
'autofilter before filtering 
    If RemoveAutoFilter Then rng_SourceRange.AutoFilter 
 
'check if filter list is not array 
    If TypeName(arr_Criteria) = "Range" Then arr_Criteria = RangeToArray(arr_Criteria) 
     
'check if there is not required negative autofilter 
    If InvertedAutoFilter Then 
        Dim arr_ColumnData 
        arr_ColumnData = WorksheetFunction.Transpose(rng_SourceRange.Resize(lng_DataRowsCount, 1).Offset(lng_HeaderRow, lng_ColumnIndex)) 
        Call z_mArrayFilterUnique(arr_ColumnData, arr_Criteria) 'replace in arr_Criteria with all data which are in the column data and it's not in column criteria 
    End If 
     
'DateAutofilter Criteria Array 
    Dim arr_DateCritieria 
    Dim i& 
     
'Date Criteria 
    ReDim arr_DateCritieria((UBound(arr_Criteria) * 2) + 1) 
    For i = LBound(arr_Criteria) To UBound(arr_Criteria) 
        arr_DateCritieria(i * 2) = 2 'criteria scope taken 0-Year,1-Month,2-Day,3-Hour,4-Minute 
        arr_DateCritieria(i * 2 + 1) = CStr(Month(arr_Criteria(i)) & "/" & Day(arr_Criteria(i)) & "/" & Year(arr_Criteria(i))) 
    Next 
     
'apply autofilter 
   rng_SourceRange.Resize(lng_DataRowsCount + 1).Offset(lng_HeaderRow).AutoFilter _ 
        Field:=lng_ColumnIndex, _ 
        Criteria2:=arr_DateCritieria, _ 
        Operator:=xlFilterValues 
         
End Function 
 
Private Function z_mArrayFilterUnique( _ 
                                    ByRef arr_SourceData, _ 
                                    ByRef arr_ResultData _ 
                                                            ) As Boolean 
 
    z_mArrayFilterUnique = False 
 
    Dim i& 
    Dim arr_BlackListCriteria 
    Dim arr_CriteriaTemp 
 
    arr_CriteriaTemp = arr_ResultData 
    arr_BlackListCriteria = arr_ResultData 
 
    ReDim arr_ResultData(0) 
    For i = LBound(arr_SourceData) To UBound(arr_SourceData) 
 
        If Not z_mFoundInArray(arr_BlackListCriteria, arr_SourceData(i)) Then  'Not Found in Black List 
         
            If Not z_mFoundInArray(arr_ResultData, arr_SourceData(i)) Then    'Not Found in Result Array means new unique value 
                arr_ResultData(UBound(arr_ResultData)) = CStr(arr_SourceData(i)) 
                ReDim Preserve arr_ResultData(UBound(arr_ResultData) + 1) 
            End If 
     
        End If 
     
    Next 
 
    If IsEmpty(arr_ResultData(0)) Then 
        arr_ResultData = arr_CriteriaTemp 
    Else 
        ReDim Preserve arr_ResultData(UBound(arr_ResultData) - 1) 
        z_mArrayFilterUnique = True 
    End If 
     
End Function 
Function TableCopyVisible( _ 
                            ByVal tbl_Object As ListObject, _ 
                            Optional ByRef arr_ColumnNames _ 
                            ) 
    Dim i& 
    Dim str_TableName$ 
    Dim rng_TempRangeArea As Range 
 
    str_TableName = tbl_Object.Name 
 
    With tbl_Object.Range.Worksheet 
 
        'tbl_Object.Range.AutoFilter 
        tbl_Object.ShowAutoFilter = True 
 
 
        If Not IsMissing(arr_ColumnNames) Then 
             
            If TypeName(arr_ColumnNames) = "Range" Then arr_ColumnNames = Me.RangeToArray(arr_ColumnNames) 
             
            For i = LBound(arr_ColumnNames) To UBound(arr_ColumnNames) 
         
                'combine column to one copy clipboard 
                If i = LBound(arr_ColumnNames) Then 
                    Set rng_TempRangeArea = .Range(tbl_Object.Name & "[" & arr_ColumnNames(i) & "]").SpecialCells(xlCellTypeVisible) 
                Else 
                    Set rng_TempRangeArea = Union(rng_TempRangeArea, .Range(tbl_Object.Name & "[" & arr_ColumnNames(i) & "]").SpecialCells(xlCellTypeVisible)) 
                End If 
                 
            Next 
             
            rng_TempRangeArea.Copy 
         
        Else 
         
            tbl_Object.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy 
     
        End If 
         
    End With 
     
End Function 
 
'paste on the end of the select table data from the clipboard table is automatically exteded 
'if there are formulas in the table outside pasted data they will recalculated 
'default way of data paste is values 
 
'Parameters 
'mandatory  r   rng_Table       Range       : enter any range with is in the table (table name) 
'optional   r   lng_PasteType   Option      : let selected user in which way data should be pasted 
 
Function TablePasteAppend( _ 
                            ByVal tbl_Object As ListObject, _ 
                            Optional lng_PasteType As XlPasteType = xlPasteValues _ 
                            ) 
     
    On Error GoTo Err 
     
'VARIABLE DEFINE 
     
    Dim str_TableName$ 
    Dim lng_CopyRowsCount& 
 
 
 
    str_TableName = tbl_Object.Name 
     
'INPUT CHECK 
'PRINTING STATE 
    Application.Statusbar = "Pastining into table: " & str_TableName 
     
     
'FUNCTION ACTION 
    With tbl_Object.Range 
      
        If .Rows.Count = 2 Then 
             
            If z_ContainsCellType(tbl_Object.HeaderRowRange.Offset(1), xlCellTypeConstants) Then ' .SpecialCells(xlCellTypeConstants).Rows.Count = 0 Then 
                lng_CopyRowsCount = 1 
            Else 
                lng_CopyRowsCount = 0 
            End If 
         
        Else 
            lng_CopyRowsCount = .Rows.Count - 1 
        End If 
 
    End With 
 
    'Consolidating ERP 
    Dim clipboard As MSForms.DataObject 
 
 
    Set clipboard = New MSForms.DataObject 
    clipboard.GetFromClipboard 
 
 
    On Error Resume Next 
    'resize table on size of new range 
    Call tbl_Object.Resize(tbl_Object.Range.Resize(1 + lng_CopyRowsCount + UBound(Split(clipboard.GetText, vbCrLf)))) 
 
 
    Call tbl_Object.Range.Offset(1 + lng_CopyRowsCount).Cells(1, 1).PasteSpecial(lng_PasteType) 
 
    Application.Statusbar = False 
 
    tbl_Object.Range.Calculate 
 
    Exit Function 
 
Err: 
 
    Debug.Print Err.Number 
    Select Case Err.Number 
        Case 1004 
        Case 0 
    End Select 
 
End Function 
 
Function TableAppendData( _ 
                            ByVal tbl_PasteAppend As ListObject, _ 
    ByVal var_DataSource As Variant, _ 
                            Optional ByVal bool_Transpose As Boolean = False, _ 
                            Optional ByVal arr_HeaderMaskInput As Variant _ 
                            ) 
 
    Application.Statusbar = IIf(bool_Transpose, "Transposing and ", "") & "Appending Data to table to " & tbl_PasteAppend.Name 
     
'CALCULATE ROWS TO COPY 
    On Error GoTo Err 
 
'PREPARE COPY 
    Dim arr_CopyData 'data itself 
    Dim arr_CopyHeader 'data columen names 
    Dim arr_VisibleRows 'visible rows (used for filtered table ranges) 
     
    Call z_mSourceToArrays(var_DataSource, bool_Transpose, arr_CopyData, arr_CopyHeader, arr_VisibleRows) 
     
'EXIT IF EMPTY ROWS 
    If UBound(arr_VisibleRows) = 0 Then Exit Function ' nothing to append 
     
    ReDim Preserve arr_VisibleRows(UBound(arr_VisibleRows) - 1) 
 
'PREPARE PASTE 
    Dim arr_Paste As Variant 
    arr_Paste = tbl_PasteAppend.Range.Formula 
     
    Dim arr_HeaderRelation() 
    Dim arr_HeaderArrayFormula() 
     
    Dim c_AppendCursor As Long 
    Dim c_CopyCursor As Long 
     
    Dim lng_ColumnInCopy As Long 
     
    ReDim arr_HeaderRelation(1 To UBound(arr_Paste, 2)) 
    ReDim arr_HeaderArrayFormula(1 To UBound(arr_Paste, 2)) 
     
'IF THERE IS header names MASK CONVERT original TABLE HEADER TO IT 
    If Not IsMissing(arr_HeaderMaskInput) Then 
         
        Dim arr_HeaderMask() 
     
        arr_HeaderMask = z_mCovertToSimpleArray(arr_HeaderMaskInput) 
             
        For c_AppendCursor = LBound(arr_Paste, 2) To UBound(arr_Paste, 2) 
         
            If c_AppendCursor > UBound(arr_HeaderMask) + 1 Then 
                arr_Paste(1, c_AppendCursor) = "" 
            Else 
                arr_Paste(1, c_AppendCursor) = Replace(arr_HeaderMask(c_AppendCursor - 1), "#", "'#") 
            End If 
             
        Next 
 
    End If 
     
'Map founded Header column names indexes 
    For c_AppendCursor = LBound(arr_Paste, 2) To UBound(arr_Paste, 2) 
         
        lng_ColumnInCopy = -1 
         
    'SEARCHING LOOP 
        For c_CopyCursor = LBound(arr_CopyHeader, 2) To UBound(arr_CopyHeader, 2) 
            If arr_Paste(1, c_AppendCursor) = arr_CopyHeader(1, c_CopyCursor) Then 
                lng_ColumnInCopy = c_CopyCursor 
                Exit For 'Found 
            End If 
        Next 
     
        arr_HeaderArrayFormula(c_AppendCursor) = tbl_PasteAppend.HeaderRowRange.Resize(1, 1).Offset(1, c_AppendCursor - 1).HasArray 
        arr_HeaderRelation(c_AppendCursor) = lng_ColumnInCopy 
     
    Next 
     
    Dim lng_AppendRows As Long 
    lng_AppendRows = UBound(arr_Paste) 
     
'Merge original table data structure with new data 
    Dim arr_Append() 
    ReDim arr_Append(1 To UBound(arr_VisibleRows) + 1, 1 To UBound(arr_Paste, 2)) 
 
    For c_AppendCursor = LBound(arr_HeaderRelation) To UBound(arr_HeaderRelation) 
         
        If arr_HeaderRelation(c_AppendCursor) <> "-1" Then 
             
            For c_CopyCursor = LBound(arr_VisibleRows) To UBound(arr_VisibleRows) 
                arr_Append(c_CopyCursor + 1, c_AppendCursor) = arr_CopyData(arr_VisibleRows(c_CopyCursor), arr_HeaderRelation(c_AppendCursor)) 
            Next 
        Else 
             
            If arr_Paste(2, c_AppendCursor) Like "=*" Then 'extend formula 
             
                For c_CopyCursor = LBound(arr_VisibleRows) To UBound(arr_VisibleRows) 
                    arr_Append(c_CopyCursor + 1, c_AppendCursor) = arr_Paste(2, c_AppendCursor) 
                Next 
             
            End If 
 
        End If 
     
    Next 
 
'Empty Table first line check 
    Dim lng_PasteOffset As Long 
    Dim lng_PasteResizeCorrection As Long 
 
    If tbl_PasteAppend.Range.Rows.Count = 2 Then 
 
        If z_ContainsCellType(tbl_PasteAppend.HeaderRowRange.Offset(1), xlCellTypeConstants) Then 
            lng_PasteResizeCorrection = 0 
        Else 
            lng_PasteResizeCorrection = 1 
        End If 
    End If 
 
'Define append area rows count 
    Dim rng_AppendArea As Range 
     
    Set rng_AppendArea = tbl_PasteAppend.Range.Offset(tbl_PasteAppend.Range.Rows.Count - lng_PasteResizeCorrection).Resize(UBound(arr_Append)) 
     
    'RESIZE PASTE TABLE 
    Call tbl_PasteAppend.Resize(tbl_PasteAppend.Range.Resize((UBound(arr_Paste) - lng_PasteResizeCorrection) + (UBound(arr_VisibleRows) + 1))) 
     
    rng_AppendArea = arr_Append 
     
    'CHECK IF THERE WAS ANY ARRAY FORMULAS 
    For c_AppendCursor = LBound(arr_HeaderArrayFormula) To UBound(arr_HeaderArrayFormula) 
        If arr_HeaderArrayFormula(c_AppendCursor) And arr_HeaderRelation(c_AppendCursor) = -1 Then 
            rng_AppendArea.Resize(1, 1).Offset(, c_AppendCursor - 1).FormulaArray = arr_Paste(2, c_AppendCursor) 
            rng_AppendArea.Resize(, 1).Offset(, c_AppendCursor - 1).FillDown 
        End If 
    Next 
     
    rng_AppendArea.Calculate 
 
    DoEvents 
    DoEvents 
 
    Application.Statusbar = False 
 
    Exit Function 
Err: 
   
    Debug.Print Err.Number & Err.Description 
    Select Case Err.Number 
        Case 1004 
        Case 0 
    End Select 
  'Resume 
End Function 
 
Private Function z_mSourceToArrays( _ 
    ByVal obj_Source As Object, _ 
    ByVal bool_Transpose As Boolean, _ 
    ByRef var_Data As Variant, _ 
    ByRef var_Header As Variant, _ 
    ByRef var_VisibleRows As Variant _ 
        ) As Boolean 
 
    Dim str_TypeName As String 
     
    str_TypeName = TypeName(obj_Source) 
 
    Dim arr_CopyTemp As Variant 
    Dim arr_HeaderTemp As Variant 
    Dim arr_VisibleRowsTemp() 
     
    ReDim arr_VisibleRowsTemp(0) 
     
    Dim rng_Cursor As Range 
     
    Dim lng_RowCounter As Long 
    lng_RowCounter = 1 
     
    If Not bool_Transpose Then 
     
        Select Case str_TypeName 
 
            Case "ListObject" 
                 
                'DATA 
                'check if source contains data 
                If obj_Source.DataBodyRange Is Nothing Then Exit Function 
                 
                Call obj_Source.Range.Calculate 
                 
                arr_CopyTemp = obj_Source.DataBodyRange.Value2 
                 
                'HEADER 
                arr_HeaderTemp = obj_Source.HeaderRowRange.Value2 'header row as column 
             
                'VISIBLE ROWS 
                For Each rng_Cursor In obj_Source.DataBodyRange.Resize(, 1) 
                    If rng_Cursor.EntireRow.Hidden = False Then 
                        arr_VisibleRowsTemp(UBound(arr_VisibleRowsTemp)) = lng_RowCounter 
                        ReDim Preserve arr_VisibleRowsTemp(UBound(arr_VisibleRowsTemp) + 1) 
                    End If 
                       
                    lng_RowCounter = lng_RowCounter + 1 
                Next 
             
            Case "PivotTable" 
                 
                'DATA 
                obj_Source.PivotCache.Refresh 
                 
                Dim rng_DataRange As Range 
                  
                Set rng_DataRange = Intersect(obj_Source.TableRange1, obj_Source.DataBodyRange.EntireRow) 
                Set rng_DataRange = rng_DataRange.Resize(rng_DataRange.Rows.Count - Abs(CLng(obj_Source.RowGrand)), rng_DataRange.Columns.Count - Abs(CLng(obj_Source.ColumnGrand))) 
                 
                arr_CopyTemp = rng_DataRange.Value2 
                 
                'HEADER 
                ReDim arr_HeaderTemp(1 To 1, 1 To rng_DataRange.Columns.Count) 
                Dim lng_ColumnCounter As Long 
                   
                lng_ColumnCounter = 1 
                 
                For Each rng_Cursor In rng_DataRange.Resize(1).Offset(-1) 
                    arr_HeaderTemp(1, lng_ColumnCounter) = rng_Cursor.Text 
                   
                    lng_ColumnCounter = lng_ColumnCounter + 1 
                   
                Next 
                 
                'VISIBLE ROWS 
                ReDim arr_VisibleRowsTemp(0 To UBound(arr_CopyTemp)) 
                 
                For lng_RowCounter = 0 To UBound(arr_CopyTemp) - 1 
                    arr_VisibleRowsTemp(lng_RowCounter) = lng_RowCounter + 1 
                Next 
 
 
            Case "Recordset" 
             
                'DATA 
                arr_CopyTemp = obj_Source.GetRows 
                arr_CopyTemp = WorksheetFunction.Transpose(arr_CopyTemp) 
 
                'HEADER 
                Dim fdl_Cursor As Object 
                Dim lng_FieldCounter As Long 
                ReDim arr_HeaderTemp(1 To 1, 1 To obj_Source.Fields.Count) 
                 
                lng_FieldCounter = 1 
                 
                For Each fdl_Cursor In obj_Source.Fields 
 
                    arr_HeaderTemp(1, lng_FieldCounter) = fdl_Cursor.Name 
                    lng_FieldCounter = lng_FieldCounter + 1 
                Next 
 
                'VISIBLE ROWS 
                ReDim arr_VisibleRowsTemp(0 To UBound(arr_CopyTemp)) 
 
                For lng_RowCounter = 0 To UBound(arr_CopyTemp) - 1 
                     
                    arr_VisibleRowsTemp(lng_RowCounter) = lng_RowCounter + 1 
                     
                Next 
 
        End Select 
 
    Else 
     
        Select Case str_TypeName 
     
            Case "ListObject" 
                 
                'DATA 
                Call obj_Source.Range.Calculate 
                 
                arr_CopyTemp = obj_Source.Range.Offset(, 1).Resize(, obj_Source.HeaderRowRange.Count - 1).Value2 
                arr_CopyTemp = WorksheetFunction.Transpose(arr_CopyTemp) 
                 
                'HEADER 
                arr_HeaderTemp = WorksheetFunction.Transpose(obj_Source.Range.Resize(, 1).Value2) 'first column as header 
                 
                'VISIBLE ROWS 
                For Each rng_Cursor In obj_Source.HeaderRowRange.Offset(, 1).Resize(, obj_Source.HeaderRowRange.Count - 1) 
                    If rng_Cursor.EntireColumn.Hidden = False Then 
                        arr_VisibleRowsTemp(UBound(arr_VisibleRowsTemp)) = lng_RowCounter 
                        ReDim Preserve arr_VisibleRowsTemp(UBound(arr_VisibleRowsTemp) + 1) 
                    End If 
                       
                    lng_RowCounter = lng_RowCounter + 1 
                Next 
               
            Case "PivotTable" 
             
             
            Case "Recordset" 
     
                 
     
        End Select 
         
    End If 
     
    var_Data = arr_CopyTemp 
    var_Header = arr_HeaderTemp 
    var_VisibleRows = arr_VisibleRowsTemp 
     
'DELETE ALL DATA 
    Erase arr_CopyTemp 
    Erase arr_HeaderTemp 
    Erase arr_VisibleRowsTemp 
 
End Function 
 
Function TableAppendToTable( _ 
                            ByVal tbl_Copy As ListObject, _ 
                            ByVal tbl_PasteAppend As ListObject, _ 
                            Optional ByVal bool_Transpose As Boolean = False, _ 
                            Optional ByVal arr_HeaderMaskInput As Variant _ 
                            ) 
   
    Call TableAppendData(tbl_PasteAppend, tbl_Copy, bool_Transpose, arr_HeaderMaskInput) 
  
End Function 
 
Private Function z_ContainsCellType(ByVal rng_Target As Range, lng_CellType As XlCellType) As Boolean 
 
    z_ContainsCellType = False 
 
    On Error GoTo Err 
 
    z_ContainsCellType = rng_Target.SpecialCells(lng_CellType).Count > 0 
     
Exit Function 
 
Err: 
    Debug.Print "z_ContainsCellType: "; Err.Number & " " & Err.Description 
    Select Case Err.Number 
        Case 1004:  Resume Next 
        Case 0 
    End Select 
     
End Function 
 
Function TableSort( _ 
                    ByVal tbl_Object As ListObject, _ 
                    ByRef str_ColumnName$, _ 
                    ByRef Orientation As XlSortOrder _ 
                    ) 
 
    Application.Statusbar = "Sorting table " & tbl_Object.Name & " in Column " & str_ColumnName & IIf(Orientation = xlAscending, " ascending", " descending") 
 
    With tbl_Object.Sort 
     
        .SortFields.Clear 
        .SortFields.Add _ 
            Key:=tbl_Object.Range.Worksheet.Range(tbl_Object.Name & "[[#All],[" & Replace(str_ColumnName, "#", "'#") & "]]"), _ 
            SortOn:=xlSortOnValues, _ 
            Order:=Orientation, _ 
            DataOption:=xlSortNormal 
     
        .Header = xlYes 
        .MatchCase = False 
        .Orientation = xlTopToBottom 
        .SortMethod = xlPinYin 
        .Apply 
    End With 
 
    Application.Statusbar = False 
 
End Function 
 
Function TableRemoteRefresh( _ 
                            ByVal tbl_Object As ListObject _ 
                                                            ) As Boolean 
    On Error GoTo err_AtkError 
     
    Application.Statusbar = "Refreshing remote data in table " & tbl_Object.Name 
 
  
'Clear Clipboard 
    Application.CutCopyMode = False 
  
 
'EXECUTE 
    Call tbl_Object.QueryTable.Refresh(BackgroundQuery:=False) 
 
    DoEvents 
    DoEvents 
     
    Call tbl_Object.Range.Calculate 
 
    DoEvents 
    DoEvents 
     
exit_ok: 
    TableRemoteRefresh = True 
Exit Function 
err_AtkError: 
  
    If frm_GlErr.z_Show(Err) Then 
        Stop: Resume 
    Else 
        Resume err_AtkError 
    End If 
 
End Function 
 
Function TableToCSV( _ 
                        ByVal tbl_Object As ListObject, _ 
                        Optional ByRef str_FileName As String = "", _ 
                        Optional ByRef str_Path As String = "", _ 
                        Optional ByRef str_Separator As String = "|", _ 
                        Optional ByRef bool_AddSeparatorLine As Boolean = False, _ 
                        Optional ByRef bool_IncludeHeader As Boolean = True, _ 
                        Optional ByRef bool_ClearDummyHeaders As Boolean = False _ 
                                                                    ) As Boolean 
 
'CHECK DATA INPUT 
    If str_FileName = "" Then str_FileName = tbl_Object.Name 
    If str_Path = "" Then str_Path = tbl_Object.Range.Worksheet.Parent.Path 
 
    str_FileName = str_FileName & ".csv" 
 
'VARIABLE 
 
    Dim rng_Cursor As Range 
    Dim rng_Data As Range 
    Dim lng_ColumnsCount As Long 
    Dim arr_Source() 
     
    'check if there are any data in the range 
    On Error Resume Next 
        If bool_IncludeHeader Then 
            Set rng_Data = tbl_Object.Range 
        Else 
            Set rng_Data = tbl_Object.DataBodyRange 
        End If 
    On Error GoTo 0 
     
    If rng_Data Is Nothing Then 
        Exit Function 
    End If 
 
    arr_Source = rng_Data.Value 
     
    Dim r_Cursor As Long 
    Dim c_Cursor As Long 
     
    'Clearing Dummy headers 
    If bool_ClearDummyHeaders Then 
        For c_Cursor = 1 To UBound(arr_Source, 2) 
            If arr_Source(1, c_Cursor) Like "Dummy*" Then arr_Source(1, c_Cursor) = "" 
        Next 
    End If 
     
     
    Dim arr_RowTemp() 
    Dim arr_ItemsTemp() 
     
    ReDim arr_RowTemp(1 To UBound(arr_Source, 1)) 
    ReDim arr_ItemsTemp(1 To UBound(arr_Source, 2)) 
     
     
    For r_Cursor = 1 To UBound(arr_Source, 1) 
        For c_Cursor = 1 To UBound(arr_Source, 2) 
             arr_ItemsTemp(c_Cursor) = arr_Source(r_Cursor, c_Cursor) 
        Next 
         
        arr_RowTemp(r_Cursor) = Join(arr_ItemsTemp, str_Separator) 
    Next 
 
    Dim file_Output As Object 
 
    With CreateObject("Scripting.FileSystemObject") 
     
        'write file 
        Set file_Output = .CreateTextFile(str_Path & "\" & str_FileName, True) 
       
        Call file_Output.WriteLine(IIf(bool_AddSeparatorLine, "sep=" & str_Separator & vbCrLf, "") & Join(arr_RowTemp, vbCrLf)) 
        Call file_Output.Close 
 
    End With 
 
End Function 
 
Function TableSharePointLoad( _ 
    ByRef tbl_Object As ListObject, _ 
    ByVal str_Path As String, _ 
    ByVal str_ListId As String, _ 
    ByVal str_ViewId As String) 
 
    On Error GoTo Err: 
 
    Dim str_Name As String 
 
    Dim rng_Address As Range 
 
     
    str_Name = tbl_Object.Name 
    Set rng_Address = tbl_Object.Range.Resize(1, 1) 
     
    tbl_Object.Delete 
 
    Dim src(2) As Variant 
    src(0) = str_Path 
    src(1) = str_ListId 
    src(2) = str_ViewId 
 
 
    Set tbl_Object = rng_Address.Worksheet.ListObjects.Add(xlSrcExternal, src, True, xlYes, rng_Address) 
 
    GoTo RunExit 
     
Err: 
 
    Set tbl_Object = rng_Address.Worksheet.ListObjects.Add(, , , , rng_Address) 
 
RunExit: 
 
     tbl_Object.Name = str_Name 
End Function 
 
Function TableSharepointSave(ByRef tbl_Object As ListObject) 
 
    tbl_Object.UpdateChanges xlListConflictDialog ' 
     
End Function 
 
'Obsolete naming Left for Compatibilty 
Function TableColumnLoop( _ 
                            ByRef tbl_SourceTable As ListObject, _ 
                            Optional ByRef rng_Cursor As Range = Nothing, _ 
                            Optional ByRef str_ColumnName As String = "", _ 
                            Optional ByVal lng_SpecialCells As XlCellType = 0 _ 
                                                                                ) As Boolean 
 
   TableColumnLoop = TableLoop(tbl_SourceTable, rng_Cursor, str_ColumnName, lng_SpecialCells) 
 
End Function 
 
Function TableLoop( _ 
    ByRef tbl_SourceTable As ListObject, _ 
    Optional ByRef rng_Cursor, _ 
    Optional ByVal str_ColumnName As String = "", _ 
    Optional ByVal lng_SpecialCells As e_TableLoopTypes = 0, _ 
    Optional ByVal bool_RemoveFilters As e_ResetFilter = e_LeaveCurrentFilters _ 
        ) As Boolean 
 
On Error GoTo Err: 
 
'VARIABLES 
    TableLoop = False 
    Dim str_TableKey As String 
    Dim str_ColumnKey As String 
    Dim cc_Cursors As Collection 
       
'CREATE UNIQUE KEY TABLE SCOPE FILENAME + TABLENAME 
    str_TableKey = tbl_SourceTable.Range.Worksheet.Parent.FullName & tbl_SourceTable.Name 
     
'INIT  GLOBAL COLLECTION STORAGE 
    If cc_TableLoopHeaders Is Nothing Then 
        Set cc_TableLoopHeaders = New Collection 'storage of the table headers 
        Set cc_TableLoopCursors = New Collection 
        GoTo NEW_ITEM 
    End If 
 
'CHECK IF TABLE ALREADY INITIATED IN THE LOOP 
    If Not z_mKeyExists(cc_TableLoopHeaders, str_TableKey) Then GoTo NEW_ITEM 
     
'CREATE UNIQUE KEY COLUMN SCOPE FILENAME + TABLENAME + COLUMN 
    If str_ColumnName = "" Then str_ColumnName = tbl_SourceTable.HeaderRowRange.Resize(1, 1) 'this could be issue in multiple loop in one table casses 
     
    str_ColumnKey = str_TableKey + str_ColumnName 
     
'CHECK IF TABLE WITH COLUMN IS ALREADY INITIATED 
    If z_mKeyExists(cc_TableLoopCursors, str_ColumnKey) Then GoTo EXISTING_ITEM 
     
NEW_ITEM: 
'NEW ITEM => CREATE NEW COLLECTION AND START THE LOOP 
 
    Dim cc_ColumnHeadersRelative As Collection 
    Dim cc_ColumnHeadersAbsolute As Collection 
     
    Dim cc_RowNumbers As Collection 
 
    Dim lng_ColumnCounter As Long 
    Dim rng_HeaderCursor As Range 
 
    Set cc_ColumnHeadersRelative = New Collection 'table scope 
    Set cc_ColumnHeadersAbsolute = New Collection 'table scope 
     
    Set cc_RowNumbers = New Collection 'delete 
    Set cc_Cursors = New Collection 'table column scope 
 
    lng_ColumnCounter = 0 
     
    'Cache Headers 
    For Each rng_HeaderCursor In tbl_SourceTable.HeaderRowRange 
        lng_ColumnCounter = lng_ColumnCounter + 1 
        Call cc_ColumnHeadersRelative.Add(lng_ColumnCounter, rng_HeaderCursor.Value) 
        Call cc_ColumnHeadersAbsolute.Add(rng_HeaderCursor.Column, rng_HeaderCursor.Value) 
    Next 
     
    'COLUMN PICK add scope check 
    If str_ColumnName <> "" Then lng_ColumnCounter = cc_ColumnHeadersRelative.Item(str_ColumnName) Else lng_ColumnCounter = 1 
 
    Dim lng_SpecialCellsTemp As Long 
             
    Select Case lng_SpecialCells 
        Case e_UniquesWithFilter: lng_SpecialCellsTemp = xlCellTypeVisible 
        Case e_Uniques: lng_SpecialCellsTemp = xlCellTypeConstants 'never used as 0 type not using specialcells function 
        Case e_Visible: lng_SpecialCellsTemp = xlCellTypeVisible 
        Case e_Blanks: lng_SpecialCellsTemp = xlCellTypeBlanks 
        Case e_Formulas: lng_SpecialCellsTemp = xlCellTypeFormulas 
        Case e_NonBlanks: lng_SpecialCellsTemp = xlCellTypeConstants 
    End Select 
     
    'LOAD ROWS IN TABLE 
    Dim rng_Data As Range 
    On Error Resume Next 
        If lng_SpecialCells = 0 Then 
            Set rng_Data = tbl_SourceTable.DataBodyRange.Resize(, 1).Offset(, lng_ColumnCounter - 1) 'pick first column 
        Else 
 
            Set rng_Data = tbl_SourceTable.DataBodyRange.SpecialCells(lng_SpecialCellsTemp) 
 
            If Not rng_Data Is Nothing Then 
                Set rng_Data = tbl_SourceTable.DataBodyRange.Resize(, 1).Offset(, lng_ColumnCounter - 1).SpecialCells(lng_SpecialCellsTemp) 
            End If 
        End If 
    On Error GoTo Err 
     
    If rng_Data Is Nothing Then 
        Exit Function 
    End If 
 
    'INDEX ROW NUMBERS AND COLLECTIONS 
    Dim rng_RowCursor As Range 
 
    'CLEAR AUTOFILTER FIRST 
    If bool_RemoveFilters Then Call TableAutofilter(tbl_SourceTable) 
 
    'ADD ITEMS PICK-CHECK WITH UNIQUE 
    If lng_SpecialCells = e_UniquesWithFilter Or lng_SpecialCells = e_Uniques Then 
        Dim cc_Uniques As Collection 
        Set cc_Uniques = New Collection 
 
        For Each rng_RowCursor In rng_Data 
            If Not z_mKeyExists(cc_Uniques, CStr(rng_RowCursor.Value)) Then 
                Call cc_Uniques.Add(Null, CStr(rng_RowCursor.Value)) 
                Call cc_Cursors.Add(rng_RowCursor) 
            End If 
        Next 
    Else 
     
        For Each rng_RowCursor In rng_Data 
            Call cc_Cursors.Add(rng_RowCursor) 
        Next 
         
    End If 
         
    'if collection not initiated yet initiate them 
    'DEFINE COLUMN KEY 
    If str_ColumnName = "" Then str_ColumnName = tbl_SourceTable.HeaderRowRange.Resize(1, 1) 'this could be issue in multiple loop in one table casses 
    str_ColumnKey = str_TableKey + str_ColumnName 
 
    If z_mKeyExists(cc_TableLoopHeaders, str_TableKey) Then _ 
        Call cc_TableLoopHeaders.Remove(str_TableKey) 
 
    If z_mKeyExists(cc_TableLoopCursors, str_ColumnKey) Then _ 
        Call cc_TableLoopCursors.Remove(str_ColumnKey) 
 
    Call cc_TableLoopHeaders.Add(cc_ColumnHeadersAbsolute, str_TableKey) 
    Call cc_TableLoopCursors.Add(cc_Cursors, str_ColumnKey) 
     
EXISTING_ITEM: 
'EXISTING_ITEM => JUST GO TO ANOTHER IN LOOP 
 
    'DEFINING INNER CLASS 
    Set cc_Cursors = cc_TableLoopCursors.Item(str_ColumnKey) 
 
    'END OF LOOP CASE => RESET CHECK ACTION 
    If cc_Cursors.Count = 0 Then 'everything looped exit function 
        Call cc_TableLoopCursors.Remove(str_ColumnKey)  'reset 
        Call TableAutofilter(tbl_SourceTable, str_ColumnName) 
        Exit Function 
    End If 
     
    'SETTING NEW CURSOR 
    Set rng_Cursor = cc_Cursors.Item(1) 
    Call cc_Cursors.Remove(1) 'remove used cursor from collection 
     
    If lng_SpecialCells = e_UniquesWithFilter Then 
        Call TableAutofilter(tbl_SourceTable, str_ColumnName, rng_Cursor, , e_LeaveCurrentFilters) 
    End If 
     
    TableLoop = True 
 
    Exit Function 
Err: 
 
    Err.Raise Err.Number 
    Resume 
 
End Function 
 
Function TableLoopCursor( _ 
    ByRef rng_Cursor As Range, _ 
    ByRef str_ColumnName As String _ 
        ) As Range 
 
    Dim str_TableKey As String 
     
'CHECK DATA INPUT 
    str_TableKey = rng_Cursor.Worksheet.Parent.FullName & rng_Cursor.ListObject 
 
 
    'check if table exists 
    If Not z_mKeyExists(cc_TableLoopHeaders, str_TableKey) Then Exit Function 
     
    Set TableLoopCursor = rng_Cursor.Offset(0, cc_TableLoopHeaders.Item(str_TableKey).Item(str_ColumnName) - rng_Cursor.Column) 
 
End Function 
 
 
Function RangeToArray( _ 
    ByRef rng_Input As Variant _ 
        ) As Variant 
 
    If rng_Input.Rows.Count > 1 And rng_Input.Columns.Count = 1 Then 
        RangeToArray = WorksheetFunction.Transpose(rng_Input.Value2) 
    End If 
     
    If rng_Input.Columns.Count > 1 And rng_Input.Rows.Count = 1 Then 
        RangeToArray = WorksheetFunction.Index(rng_Input.Value2, 0) 
    End If 
   
    If rng_Input.Columns.Count = 1 And rng_Input.Rows.Count = 1 Then 
        RangeToArray = Array(rng_Input.Value2) 
    End If 
   
End Function 
 
Function WorkbookByCopySheets( _ 
    ByRef wbk_Result As Workbook, _ 
    ByRef arr_SheetNamesInput, _ 
    ByVal str_NewWorkbookName As String, _ 
    Optional ByRef str_SaveWorkBookToPath$, _ 
    Optional ByRef wbk_Source As Workbook, _ 
    Optional bool_BreakLinks As Boolean = True, _ 
    Optional bool_RemoveMacros As Boolean = True _ 
        ) As Boolean 
 
    WorkbookByCopySheets = False 
 
On Error GoTo Exit_Function 
 
    'Call z_mCovertToSimpleArray(arr_SheetNames) 
    Dim arr_SheetNames 
    arr_SheetNames = z_mCovertToSimpleArray(arr_SheetNamesInput) 
 
    If wbk_Source Is Nothing Then Set wbk_Source = ThisWorkbook 
    If str_SaveWorkBookToPath = "" Then str_SaveWorkBookToPath = wbk_Source.Path 
     
    'Application.ScreenUpdating = False 
     
    'save workbook 
    wbk_Source.Save 
    
    'enter suffix to the file 
    str_NewWorkbookName = str_NewWorkbookName & "." & Split(wbk_Source.Name, ".")(1) 
     
    'Check if workbook is not open 
    Dim i_WorkbooksCursor As Long 
     
    For i_WorkbooksCursor = 1 To Application.Workbooks.Count 
        If Application.Workbooks(i_WorkbooksCursor).Name = str_NewWorkbookName Then 
            Call MsgBox("Workbook with name " & str_NewWorkbookName & " is now open. Can't open another one", vbCritical, "Error") 
            Exit Function 
        End If 
    Next 
 
    'prepare workbook path 
    Dim str_NewWorkbookPathTemp$ 
    str_NewWorkbookPathTemp = str_SaveWorkBookToPath & "\" & str_NewWorkbookName 
     
    'copy current to new file 
    Dim fs As Object 
    Dim oldPath As String, newPath As String 
    oldPath = wbk_Source.FullName 'Folder file is located in 
    newPath = str_NewWorkbookPathTemp 'Folder to copy file to 
    Set fs = CreateObject("Scripting.FileSystemObject") 
    fs.copyfile oldPath, newPath   'This file was an .xls file 
    Set fs = Nothing 
 
 
    Dim bool_UpdateLinksTemp As Boolean 
 
    bool_UpdateLinksTemp = Application.AskToUpdateLinks 
    Application.AskToUpdateLinks = False 
    'Application.DisplayAlerts = False 
    'open the file 
    Set wbk_Result = Workbooks.Open(str_NewWorkbookPathTemp) 
 
    Application.AskToUpdateLinks = bool_UpdateLinksTemp 
    'Application.DisplayAlerts = True 
 
    Dim i& 
    Dim sht_Cursor As Worksheet 
    Dim bool_SheetFound As Boolean 
 
    'check what sheets should be deleted 
    For Each sht_Cursor In wbk_Result.Worksheets 
        bool_SheetFound = False 
        For i = LBound(arr_SheetNames) To UBound(arr_SheetNames) 
            If sht_Cursor.Name = arr_SheetNames(i) Then 
                bool_SheetFound = True 
            End If 
        Next 
        'Application.DisplayAlerts = False 
        If bool_SheetFound = False Then sht_Cursor.Delete 
        'Application.DisplayAlerts = True 
    Next 
 
    'remove all links which remains by default 
    If bool_BreakLinks Then 
        Dim arr_Links As Variant 
        Dim i_LinkCursor As Long 
        arr_Links = wbk_Result.LinkSources(Type:=xlLinkTypeExcelLinks) 
        If arr_Links <> Empty Then 
            For i_LinkCursor = 1 To UBound(arr_Links) 
                Call wbk_Result.BreakLink( _ 
                    Name:=arr_Links(i_LinkCursor), _ 
                    Type:=xlLinkTypeExcelLinks) 
            Next i_LinkCursor 
        End If 
    End If 
 
     
    If bool_RemoveMacros Then 
        Dim str_WorkbookTempPath As String 
        Call wbk_Result.SaveAs(, xlOpenXMLWorkbook) 
        str_WorkbookTempPath = wbk_Result.FullName 
        wbk_Result.Close 
        Set wbk_Result = Workbooks.Open(str_WorkbookTempPath) 
        Call wbk_Result.SaveAs(, xlExcel12) 
        Kill str_WorkbookTempPath 
    End If 
 
    WorkbookByCopySheets = True 
     
Exit_Function: 
 
    If Err.Number <> 0 Then 
 
        Err.Raise Err.Number, Err.Source, Err.Description 
    End If 
 
End Function 
 
Function TableRemoveDuplicates( _ 
    ByVal tbl_Object As ListObject, _ 
    Optional ByRef arr_ColumnsCriteriaInput _ 
        ) 
 
 
    Dim str_TableName$ 
 
 
    str_TableName = tbl_Object.Name 
 
    Dim i& 
    Dim arr_FunctionCriteria() 
    Dim lng_TopLeftColumn& 
    Dim arr_ColumnsCriteria() 
 
    If Not IsMissing(arr_ColumnsCriteriaInput) Then 
         
        'DEFIN TOP LEFT COLUMN 
        lng_TopLeftColumn = tbl_Object.Range.Resize(1, 1).Column 
         
        arr_ColumnsCriteria = z_mCovertToSimpleArray(arr_ColumnsCriteriaInput) 
         
        'If TypeName(arr_ColumnsCriteria) = "Range" Then arr_ColumnsCriteria = Me.RangeToArray(arr_ColumnsCriteria) 
         
        ReDim arr_FunctionCriteria(0) 
         
        For i = LBound(arr_ColumnsCriteria) To UBound(arr_ColumnsCriteria) 
 
            arr_FunctionCriteria(UBound(arr_FunctionCriteria)) = tbl_Object.HeaderRowRange.Find(CStr(arr_ColumnsCriteria(i))).Column - (tbl_Object.HeaderRowRange.Resize(, 1).Column - 1) 
             
            If i <> UBound(arr_ColumnsCriteria) Then 
                 ReDim Preserve arr_FunctionCriteria(UBound(arr_FunctionCriteria) + 1) 
            End If 
             
        Next 
 
        Call tbl_Object.Range.RemoveDuplicates(Columns:=CVar(arr_FunctionCriteria), Header:=xlYes) 
 
    Else 
     
        ReDim arr_FunctionCriteria(tbl_Object.HeaderRowRange.Count - 1) 
     
        For i = 0 To tbl_Object.HeaderRowRange.Count - 1 
            arr_FunctionCriteria(i) = i + 1 
        Next 
     
        Call tbl_Object.Range.RemoveDuplicates(Columns:=CVar(arr_FunctionCriteria), Header:=xlYes) 
 
    End If 
 
 
End Function 
 
Function RangeRemoveDuplicates( _ 
    ByRef rng_SourceRange As Range, _ 
    Optional lng_HeaderRows& = 1, _ 
    Optional ByRef arr_ColumnsCriteria _ 
    ) 
 
 
    Dim i& 
    Dim arr_FunctionCriteria() 
    Dim lng_TopLeftColumn& 
 
    Set rng_SourceRange = rng_SourceRange.CurrentRegion 
     
    Dim lng_CopyRowsCount& 
     
    lng_CopyRowsCount = rng_SourceRange.CurrentRegion.Rows.Count - lng_HeaderRows 
 
 
    If Not IsMissing(arr_ColumnsCriteria) Then 
         
        'DEFIN TOP LEFT COLUMN 
        lng_TopLeftColumn = rng_SourceRange.Resize(1, 1).Column + 2 
         
        If TypeName(arr_ColumnsCriteria) = "Range" Then arr_ColumnsCriteria = Me.RangeToArray(arr_ColumnsCriteria) 
         
        ReDim arr_FunctionCriteria(0) 
         
        For i = LBound(arr_ColumnsCriteria) To UBound(arr_ColumnsCriteria) 
 
           arr_FunctionCriteria(UBound(arr_FunctionCriteria)) = lng_TopLeftColumn - rng_SourceRange.Resize(1).Find(CStr(arr_ColumnsCriteria(i))).Column 
             
           If i <> UBound(arr_ColumnsCriteria) Then 
                ReDim Preserve arr_FunctionCriteria(UBound(arr_FunctionCriteria) + 1) 
           End If 
             
        Next 
     
        Call rng_SourceRange.RemoveDuplicates(Columns:=CVar(arr_FunctionCriteria), Header:=xlYes) 
    Else 
     
        Call rng_SourceRange.RemoveDuplicates(Header:=xlYes) 
    End If 
 
End Function 
 
 
Function RemoveColumnsByRow( _ 
    ByRef rngRow As Range, _ 
    ByVal lng_SpecialCell As XlCellType _ 
        ) 
 
 
    rngRow.SpecialCells(lng_SpecialCell).EntireColumn.Delete 
 
End Function 
 
Function RemoveRowsByColumn( _ 
    ByRef rngColumn As Range, _ 
    ByVal lng_SpecialCell As XlCellType _ 
        ) 
 
 
    rngColumn.SpecialCells(lng_SpecialCell).EntireRow.Delete 
 
End Function 
 
'put to copy mode range of all cells which are neighbors with each other and with the defined 
'cell in sourcerange parameter 
'its allways as square selection 
'there is possible to difine header row to identify in where table data starts 
'other option is enter array of header names to put to copy only selected columns if this is 
'avoid all data are copied 
 
'Parameters 
'mandatory  r   rng_SourceRange Range       : enter one of the cell of the source data 
'optional   r   lng_HeaderRows  Long        : enter define line in selecte data range where is header 
'optional   r   arr_ColumnNames Array       : array with names of headers to be copied 
 
Function RangeCopyWithoutHeader( _ 
    ByRef rng_SourceRange As Range, _ 
    Optional lng_HeaderRows& = 1, _ 
    Optional ByRef arr_ColumnNames _ 
        ) 
 
 
    Dim lng_CopyRowsCount& 
     
    lng_CopyRowsCount = rng_SourceRange.CurrentRegion.Rows.Count - lng_HeaderRows 
     
    'check columns to copy 
    Dim rng_TempRangeArea As Range 
    Dim rng_HeaderCursor As Range 
    Dim i& 'array index 
     
    If Not IsMissing(arr_ColumnNames) Then 
         
        If TypeName(arr_ColumnNames) = "Range" Then arr_ColumnNames = Me.RangeToArray(arr_ColumnNames) 
         
        For i = LBound(arr_ColumnNames) To UBound(arr_ColumnNames) 
     
            'combine column to one copy clipboard 
            If i = LBound(arr_ColumnNames) Then 
                Set rng_TempRangeArea = rng_SourceRange.CurrentRegion.Offset(lng_HeaderRows - 1).Resize(1).Find(arr_ColumnNames(i)).Offset(1).Resize(lng_CopyRowsCount) 
            Else 
                Set rng_TempRangeArea = Union(rng_TempRangeArea, rng_SourceRange.CurrentRegion.Offset(lng_HeaderRows - 1).Resize(1).Find(arr_ColumnNames(i)).Offset(1).Resize(lng_CopyRowsCount)) 
            End If 
             
        Next 
         
        rng_TempRangeArea.Copy 
     
    Else 
        rng_SourceRange.CurrentRegion.Offset(lng_HeaderRows).Resize(lng_CopyRowsCount).Copy 
     
    End If 
 
End Function 
 
Function TableCopy( _ 
    ByVal tbl_Object As ListObject, _ 
    Optional ByRef arr_ColumnNames _ 
        ) 
         
 
    Dim i& 
    Dim str_TableName$ 
    Dim rng_TempRangeArea As Range 
 
    str_TableName = tbl_Object.Name 
 
    With tbl_Object.Range.Worksheet 
 
        tbl_Object.Range.AutoFilter 
        tbl_Object.ShowAutoFilter = True 
 
 
        If Not IsMissing(arr_ColumnNames) Then 
             
            If TypeName(arr_ColumnNames) = "Range" Then arr_ColumnNames = Me.RangeToArray(arr_ColumnNames) 
             
            For i = LBound(arr_ColumnNames) To UBound(arr_ColumnNames) 
         
                'combine column to one copy clipboard 
                If i = LBound(arr_ColumnNames) Then 
                    Set rng_TempRangeArea = .Range(z_mTableColumn(tbl_Object, arr_ColumnNames(i))) 
                Else 
                    Set rng_TempRangeArea = Union(rng_TempRangeArea, .Range(z_mTableColumn(tbl_Object, arr_ColumnNames(i)))) 
                End If 
                 
            Next 
             
             
             
            rng_TempRangeArea.Copy 
         
        Else 
         
            tbl_Object.DataBodyRange.Copy 
     
        End If 
         
    End With 
     
End Function 
 
Private Function z_mTableColumn( _ 
    ByVal tbl_Object As ListObject, _ 
    ByVal var_ColumnName As Variant _ 
        ) As String 
 
    If IsNumeric(var_ColumnName) Then 
        var_ColumnName = tbl_Object.HeaderRowRange.Resize(, 1).Offset(, var_ColumnName).Value 
    End If 
 
    If InStr(1, var_ColumnName, "#") > 0 Then 
        z_mTableColumn = tbl_Object.Name & "[" & Replace(var_ColumnName, "#", "'#") & "]" 
    Else 
        z_mTableColumn = tbl_Object.Name & "[" & var_ColumnName & "]" 
    End If 
 
End Function 
 
Function RangeReOrderColumns( _ 
    ByRef rng_SourceRange As Range, _ 
    Optional lng_HeaderRows& = 1, _ 
    Optional ByRef arr_ColumnNames _ 
        ) 
 
 
    If IsMissing(arr_ColumnNames) Then Exit Function 
 
    Dim lng_CopyRowsCount& 
     
    lng_CopyRowsCount = rng_SourceRange.CurrentRegion.Rows.Count - lng_HeaderRows 
     
    'check columns to copy 
    Dim rng_TempRangeArea As Range 
    Dim rng_HeaderCursor As Range 
    Dim i& 'array index 
     
    Dim str_FirstColumnAddress$ 
     
    If TypeName(arr_ColumnNames) = "Range" Then arr_ColumnNames = Me.RangeToArray(arr_ColumnNames) 
     
    str_FirstColumnAddress = rng_SourceRange.Resize(, 1).EntireColumn.AddressLocal 
     
    For i = UBound(arr_ColumnNames) To LBound(arr_ColumnNames) Step -1 
 
        'cut and insert 
        rng_SourceRange.CurrentRegion.Offset(lng_HeaderRows - 1).Resize(1).Find(arr_ColumnNames(i)).Resize(lng_CopyRowsCount + lng_HeaderRows).Cut 
        rng_SourceRange.Worksheet.Columns(str_FirstColumnAddress).Insert Shift:=xlToRight 
        
    Next 
 
 
End Function 
 
 
'TO BE REMOVED 
'Function OpenWorkbook( _ 
'                        ByRef wbk_Source As Workbook, _ 
'                        Optional ByVal str_Path$ = "", _ 
'                        Optional ByVal str_DialogHeader _ 
'                                                        ) As Boolean 
' 
'    'for compatibility with old versions of ATk 
'    OpenWorkbook = WorkbookOpen(wbk_Source, , , str_DialogHeader) 
' 
'End Function 
 
 
'This function takes care about opening workbook there is several enchantements against original version 
'-if no path parameter is inserted File dialog is automatically open in workbook location 
'-if path parameter is added without file dialog file dialog is open in that location 
'-if no file selected in file dialog it ask for again selection 
'-text parameter for changing header of the window 
'-returns true/false by sitation if succesfully oppened workbook or not 
 
'Parameters 
'mandatory  rw  wbk_Source      workbook    : returns shortcut on the opened workbook 
'optional   r   str_FileName    string      : enter workbook name to search 
'optional   r   str_DirPath     string      : enter path where opened workbook should be stored 
'optional   r   str_WindowTitle string      : enter window title for possible filedialog 
 
Function WorkbookOpen( _ 
                        ByRef wbk_Source As Workbook, _ 
                        Optional ByVal str_FileName As String = "", _ 
                        Optional ByVal str_DirPath As String = "", _ 
                        Optional ByRef str_WindowTitle = "" _ 
                                                                ) As Boolean 
     
'WorkbookOpen: Open workbook
'wbk_Source: returns reference on the opened workbook
'str_Filename: text with file name > then path same as macro or full path of the file
     
    
    WorkbookOpen = False 
     
    Dim strTempDate As String 
     
    'CHECK PARAMETER IS OPEN EXIST AND IF LOCATION EXIST 
 
    If str_DirPath = "" Then 
     
        With CreateObject("Scripting.FileSystemObject") 
     
            If .FileExists(str_FileName) Then 
                 
                str_DirPath = .getabsolutepathname(str_FileName) 
                str_FileName = .GetFileName(str_FileName) 
                 
                str_DirPath = Mid(str_DirPath, 1, Len(str_DirPath) - Len(str_FileName)) 
            Else 
                str_DirPath = ThisWorkbook.Path & "\" 
            End If 
         
        End With 
         
    Else 
         
        'Dir path fix 
        If Not str_DirPath Like "*\" Then str_DirPath = str_DirPath & "\" 
        'Dir path check 
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(str_DirPath) Then str_DirPath = ThisWorkbook.Path & "\" 
    End If 
       
    Call z_mChDirNet(str_DirPath) 
       
    If str_WindowTitle = "" Then 
        str_WindowTitle = "Please select file " & str_FileName 
    End If 
       
       
'CHECK IF WORKBOOK ISN'T ALREADY OPEN 
    If z_mWorkbookAlreadyOpen(wbk_Source, str_FileName) Then GoTo FILE_OPENED: 
 
 
SELECT_FILE: 
       
    Const str_ExcelCompatibleFileFormats As String = "*.xls *.xlsx *.xlsb *.xlsm *.csv *.txt, *.*" 
       
    If str_FileName = "" Then 
        str_FileName = Application.GetOpenFilename(Title:=str_WindowTitle, FileFilter:=str_ExcelCompatibleFileFormats) 
         
        Dim bool_FileTypeAccepted As Boolean 
         
        bool_FileTypeAccepted = False 
        If LCase(str_FileName) Like "*.xls" Then bool_FileTypeAccepted = True 
        If LCase(str_FileName) Like "*.xlsx" Then bool_FileTypeAccepted = True 
        If LCase(str_FileName) Like "*.xlsb" Then bool_FileTypeAccepted = True 
        If LCase(str_FileName) Like "*.xlsm" Then bool_FileTypeAccepted = True 
        If LCase(str_FileName) Like "*.csv" Then bool_FileTypeAccepted = True 
        If LCase(str_FileName) Like "*.txt" Then bool_FileTypeAccepted = True 
         
         
        If bool_FileTypeAccepted = False Then 
             Select Case MsgBox("Wrong file type selected. Try again?", vbCritical + vbYesNo, "Wrong File type") 
                Case vbYes: str_FileName = "": GoTo SELECT_FILE 
                Case vbNo: Exit Function 
            End Select 
        End If 
         
        'exit routine when file is empty 
        If str_FileName = "False" Then 
            Select Case MsgBox("File was not selected try again?", vbCritical + vbYesNo, "File Selection problem") 
                Case vbYes: str_FileName = "": GoTo SELECT_FILE 
                Case vbNo: Exit Function 
            End Select 
        End If 
               
    Else 
         
        'CHECK IF FILENAME EXIST 
        If Not CreateObject("Scripting.FileSystemObject").FileExists(str_DirPath & str_FileName) Then 
            Select Case MsgBox("File " & str_FileName & " was not found try again?", vbCritical + vbYesNo, "File not found") 
                Case vbYes: str_FileName = "": GoTo SELECT_FILE 
                Case vbNo: Exit Function 
            End Select 
        End If 
     
        str_FileName = str_DirPath & str_FileName 
     
    End If 
 
 
 
 
 
    Set wbk_Source = Workbooks.Open(str_FileName) 
     
FILE_OPENED: 
     
    WorkbookOpen = True 
     
End Function 
 
Private Function z_mWorkbookAlreadyOpen( _ 
    ByRef wbk_Source As Workbook, _ 
    Optional ByRef str_FileName As String = "" _ 
        ) As Boolean 
 
    z_mWorkbookAlreadyOpen = False 
 
    If str_FileName = "" Then Exit Function 
 
    'CHECK IF WORKBOOK ISN'T ALREADY OPEN 
    On Error Resume Next 
    Set wbk_Source = Workbooks(str_FileName) 
    On Error GoTo 0 
       
    z_mWorkbookAlreadyOpen = Not wbk_Source Is Nothing 
 
End Function 
 
Private Function z_FillUnamedParameters() 
 
    Dim rng_Cursor As Range 
 
     
    For Each rng_Cursor In ThisWorkbook.Sheets("Main").ListObjects("t_Parameters").DataBodyRange.Resize(, 1).Offset(, 1) 
         
        Dim str_Name As String 
         
        str_Name = "" 
         
        'get name from the names library 
        On Error Resume Next 
           str_Name = ThisWorkbook.Names(, , CStr(rng_Cursor.Offset(, 1).Name)).Name 
        On Error GoTo 0 
         
        If str_Name = "" And rng_Cursor <> "" Then 
            rng_Cursor.Offset(, 1).Name = "rng_" & Replace(WorksheetFunction.Proper(rng_Cursor.Value), " ", "") 
        End If 
 
    Next 
 
End Function 
 
Private Function z_CleanRefErrNames() 
     
    Dim name_Cur As Name 
     
    For Each name_Cur In ThisWorkbook.Names 
     
        If Right(CStr(name_Cur.RefersTo), 5) = "#REF!" Then 
            name_Cur.Delete 
        End If 
     
    Next 
 
End Function 
 
Private Function z_mFilePathChecker( _ 
    ByRef str_FilePathOutput As String, _ 
    Optional ByRef str_FileName As String = "", _ 
    Optional ByVal str_DirPath As String = "", _ 
    Optional ByRef str_WindowTitle As String = "", _ 
    Optional ByVal str_ExcelCompatibleFileFormats As String = "*.*" _ 
        ) As Boolean 
     
    z_mFilePathChecker = False 
     
    Dim strTempDate As String 
     
    'CHECK PARAMETER IS OPEN EXIST AND IF LOCATION EXIST 
     
    If str_DirPath = "" Then 
     
         
        With CreateObject("Scripting.FileSystemObject") 
     
            If .FileExists(str_FileName) Then 
                 
                str_DirPath = .getabsolutepathname(str_FileName) 
                str_FileName = .GetFileName(str_FileName) 
                 
                str_DirPath = Mid(str_DirPath, 1, Len(str_DirPath) - Len(str_FileName)) 
            Else 
                str_DirPath = ThisWorkbook.Path & "\" 
            End If 
         
        End With 
         
    Else 
         
        'Dir path fix 
        If Not str_DirPath Like "*\" Then str_DirPath = str_DirPath & "\" 
        'Dir path check 
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(str_DirPath) Then str_DirPath = ThisWorkbook.Path & "\" 
    End If 
       
    Call z_mChDirNet(str_DirPath) 
       
    If str_WindowTitle = "" Then 
        str_WindowTitle = "Please select file " & str_FileName 
    End If 
  
SELECT_FILE: 
       
    'Const str_ExcelCompatibleFileFormats As String = "*.xls *.xlsx *.xlsb *.xlsm *.csv *.txt, *.*" 
       
    Dim arr_CompatibleFormats 
       
    arr_CompatibleFormats = Split(str_ExcelCompatibleFileFormats, " ") 
       
    If str_FileName = "" Then 
        str_FileName = Application.GetOpenFilename(Title:=str_WindowTitle, FileFilter:=str_ExcelCompatibleFileFormats & " ,*.*") 
         
        Dim bool_FileTypeAccepted As Boolean 
        Dim i_FormatCursor As Long 
         
        bool_FileTypeAccepted = False 
         
        'loop & check if file format fits to options 
        For i_FormatCursor = LBound(arr_CompatibleFormats) To UBound(arr_CompatibleFormats) 
            If LCase(str_FileName) Like LCase(arr_CompatibleFormats(i_FormatCursor)) Then bool_FileTypeAccepted = True 
        Next 
         
        If bool_FileTypeAccepted = False Then 
             Select Case MsgBox("Wrong file type selected, posible formats: " & str_ExcelCompatibleFileFormats & vbCrLf & "Try again?", vbCritical + vbYesNo, "Wrong File type") 
                Case vbYes: str_FileName = "": GoTo SELECT_FILE 
                Case vbNo: Exit Function 
            End Select 
        End If 
         
        'exit routine when file is empty 
        If str_FileName = "False" Then 
            Select Case MsgBox("File was not selected try again?", vbCritical + vbYesNo, "File Selection problem") 
                Case vbYes: str_FileName = "": GoTo SELECT_FILE 
                Case vbNo: Exit Function 
            End Select 
        End If 
               
        str_FilePathOutput = str_FileName 
               
    Else 
         
        'CHECK IF FILENAME EXIST 
        If Not CreateObject("Scripting.FileSystemObject").FileExists(str_DirPath & str_FileName) Then 
            Select Case MsgBox("File " & str_FileName & " was not found try again?", vbCritical + vbYesNo, "File not found") 
                Case vbYes: str_FileName = "": GoTo SELECT_FILE 
                Case vbNo: Exit Function 
            End Select 
        End If 
     
        str_FilePathOutput = str_DirPath & str_FileName 
     
    End If 
 
    z_mFilePathChecker = True 
 
End Function 
 
Function RangeColumnFilter( _ 
    rng_ToFilter As Range, _ 
    Optional ByVal arr_CriteriaInput As Variant, _ 
    Optional ByVal InvertedAutoFilter As e_FilterType = e_ShowOnlyCriteriaItems, _ 
    Optional ByVal RemoveAutoFilter As e_ResetFilter = e_LeaveCurrentFilters _ 
        ) 
 
'recalculate column data 
    Call z_mRangeRecalculate(rng_ToFilter) 
 
'check if filter list is not array 
    Dim arr_Criteria 
    arr_Criteria = z_mCovertToSimpleArray(arr_CriteriaInput) 
         
    Dim rng_Cursor As Range 
    Dim rng_ColumnsToHide As Range 
    Dim lng_CriteriaCursor As Long 
 
'Loop for filter 
    For Each rng_Cursor In rng_ToFilter 
         
        Dim bool_Found As Boolean 
         
        bool_Found = False 
         
        For lng_CriteriaCursor = LBound(arr_Criteria) To UBound(arr_Criteria) 
     
            If rng_Cursor Like arr_Criteria(lng_CriteriaCursor) Then 
                 bool_Found = True 
                 Exit For 
            End If 
 
        Next 
 
        If InvertedAutoFilter = e_ShowOnlyCriteriaItems Then 
            bool_Found = Not bool_Found 
        End If 
 
        If bool_Found Then 
         
            If rng_ColumnsToHide Is Nothing Then 
                Set rng_ColumnsToHide = rng_Cursor 
            Else 
                Set rng_ColumnsToHide = Union(rng_ColumnsToHide, rng_Cursor) 
            End If 
 
        End If 
     
    Next 
     
    If RemoveAutoFilter = e_ClearCurrentFilters Then 
        rng_ToFilter.EntireColumn.Hidden = False 
    End If 
 
    If Not rng_ColumnsToHide Is Nothing Then 
        rng_ColumnsToHide.EntireColumn.Hidden = True 
    End If 
 
End Function 
 
 
Function WorkbookOpenLoop( _ 
    ByRef wbk_Source As Workbook, _ 
    Optional ByVal str_DirPath As String = "", _ 
    Optional ByVal arr_CriteriaInput As Variant, _ 
    Optional ByVal InvertedAutoFilter As e_FilterType = e_ShowOnlyCriteriaItems _ 
        ) As Boolean 
     
    WorkbookOpenLoop = False 
     
    Dim arr_Crititeria() 
     
'CHECK DATA INPUT 
    If str_DirPath = "" Then str_DirPath = ThisWorkbook.Path & "\" 'no path defined 
    If Not str_DirPath Like "*\" Then str_DirPath = str_DirPath & "\" 'fix missing slash 
    If IsMissing(arr_CriteriaInput) Then 
        ReDim arr_Crititeria(0) 
        arr_Crititeria(0) = "*.xls*"    'set default mask 
    Else 
        arr_Crititeria = z_mCovertToSimpleArray(arr_CriteriaInput) 
    End If 
     
'VARIABLES 
 
    Dim cc_FilePaths As Collection 
    Dim lng_ItemCursor As Long 
    Dim str_FilePath As String 
     
    If z_mKeyExists(cc_WorkbookPaths, str_DirPath) Then 
     
        lng_ItemCursor = cc_WorkbookIndexes.Item(str_DirPath) 
        'str_FilePath = cc_WorkbookPaths.Item(str_DirPath & str_FileMask) 
 
        Set cc_FilePaths = cc_WorkbookPaths.Item(str_DirPath) 
         
        If lng_ItemCursor > cc_FilePaths.Count Then 'everything looped exit function 
            Set wbk_Source = Nothing 
            Call cc_WorkbookIndexes.Remove(str_DirPath)  'reset 
            Call cc_WorkbookIndexes.Add(1, str_DirPath) 
            Exit Function 
        End If 
         
        str_FilePath = cc_FilePaths.Item(lng_ItemCursor) 
     
        Call cc_WorkbookIndexes.Remove(str_DirPath) 
        Call cc_WorkbookIndexes.Add(lng_ItemCursor + 1, str_DirPath) 
     
    Else 
             
        Dim file As Object 
         
        Set cc_FilePaths = New Collection 
 
        'loop all files 
        With CreateObject("Scripting.FileSystemObject").GetFolder(str_DirPath) 
     
            'loop all files 
            For Each file In .Files 
             
                Dim str_FileName As String 
                 
                str_FileName = file.Name 
             
                Debug.Print str_FileName 
             
                Dim lng_Cursor As Long 
                Dim bool_Match As Boolean 
                 
                bool_Match = False 
                 
                For lng_Cursor = LBound(arr_Crititeria) To UBound(arr_Crititeria) 
                    If UCase(str_FileName) Like UCase(arr_Crititeria(lng_Cursor)) Then 
                        bool_Match = True 
                        Exit For 
                    End If 
                Next 
                                
                If InvertedAutoFilter = e_HideCriteriaItems Then 
                    bool_Match = Not bool_Match 
                End If 
                 
                If bool_Match And (str_FileName Like "*.xls*" Or str_FileName Like "*.csv") Then 
                    If str_FileName <> ThisWorkbook.Name Or Left(str_FileName, 2) <> "~$" Then 'protect open itslef 
                        Call cc_FilePaths.Add(str_FileName) 
                    End If 
                End If 
 
            Next 
     
        End With 
         
        
        If cc_FilePaths.Count = 0 Then 'no files with criteria were found 
            Set wbk_Source = Nothing 
            Exit Function 
        End If 
         
        str_FilePath = cc_FilePaths.Item(1) 'set first item to this 
         
        If cc_WorkbookPaths Is Nothing Then 
            Set cc_WorkbookPaths = New Collection 
            Set cc_WorkbookIndexes = New Collection 
        End If 
         
        Call cc_WorkbookPaths.Add(cc_FilePaths, str_DirPath) 
        Call cc_WorkbookIndexes.Add(2, str_DirPath) 
         
    End If 
     
 
   'Load Workbook 
    WorkbookOpenLoop = WorkbookOpen(wbk_Source, str_FilePath, str_DirPath) 
     
End Function 
 
 
Function WorkbookSave( _ 
    ByVal wbk_Source As Workbook, _ 
    Optional ByVal str_FileName As String = "", _ 
    Optional ByVal str_DirPath As String = "", _ 
    Optional dte_DateStampDay As Variant _ 
    ) 
     
    Dim wbk_Temp As Workbook 
    Dim strTempDate$ 
     
    Set wbk_Temp = Workbooks(wbk_Source.Name) 
 
    If Not IsMissing(dte_DateStampDay) And IsDate(dte_DateStampDay) Then 
        strTempDate = "_" & Format(CDate(dte_DateStampDay), "yyyymmdd") 
    Else 
        strTempDate = "" 
    End If 
     
    If str_DirPath = "" Then 
        str_DirPath = ThisWorkbook.Path & "\" 
    End If 
       
    ChDir str_DirPath 
       
    If str_FileName = "" Then 
       str_FileName = Application.GetSaveAsFilename(strTempDate & ".xlsm") 
    Else 
       str_FileName = str_DirPath & str_FileName & strTempDate & ".xlsm" 
    End If 
     
    Call wbk_Temp.SaveAs(str_FileName, 52) 
     
End Function 
 
'Workbook function to close workbook there is only one 
 
'Parameters 
'mandatory  r   wbk_Source              workbook    : close workbook in this parameter 
'optional   r   bool_SaveBeforeQuit     boolean     : predefined boolean to check if workbook should be automatically closed 
 
Function WorkbookClose( _ 
    Optional ByRef wbk_Source As Variant, _ 
    Optional BeforeClose As e_ClosingAction = e_ClosingAction.e_DontSave _ 
    ) 
     
    If IsMissing(wbk_Source) Then 
        Set wbk_Source = ThisWorkbook 
    End If 
     
    Select Case BeforeClose 
        Case e_ClosingAction.e_Save 
            Call wbk_Source.Close(True) 
        Case e_ClosingAction.e_DontSave 
            'Application.DisplayAlerts = False 
            Call wbk_Source.Close(False) 
            'Application.DisplayAlerts = True 
        Case e_ClosingAction.e_Delete 
             
            'Application.DisplayAlerts = False 
            Call wbk_Source.ChangeFileAccess(Mode:=xlReadOnly) 
            Call Kill(wbk_Source.FullName) 
            Call wbk_Source.Close(False) 
            'Application.DisplayAlerts = True 
             
    End Select 
     
    Set wbk_Source = Nothing 
     
End Function 
 
'creates link between worksheet variable and the worksheet in the defined workbook 
'TBR ?? 
'Parameters 
'mandatory  rw  sht_Link        worksheet   : returns shortcut on the opened worksheet 
'mandatory  r   wbk_Source      workbook    : enter shortcut to the opened workbook 
'optional   r   var_SheetId     variant     : enter sheet name or sheet order index if nothing is passed 
'                                              it automaticaly takes first sheet 
Function LinkSheetToShortcut( _ 
    ByRef sht_Link As Worksheet, _ 
    ByRef wbk_Source As Workbook, _ 
    Optional ByRef var_SheetId _ 
    ) As Boolean 
 
    LinkSheetToShortcut = False 
 
    On Error GoTo Err 
 
    If IsMissing(var_SheetId) Then 
        var_SheetId = 1 
    End If 
 
    Set sht_Link = wbk_Source.Sheets(var_SheetId) 
     
    LinkSheetToShortcut = True 
 
Err: 
 
    If Err.Number <> 0 Then 
        Call MsgBox("Sheet called " & var_SheetId & " not found in the workbook " & wbk_Source.Name, vbCritical, "Error") 
    End If 
     
End Function 
 
 
Function PivotRefresh( _ 
    ByRef pvt_Object As PivotTable _ 
        ) 
 
    pvt_Object.PivotCache.Refresh 
     
End Function 
 
Function PivotCleanCache( _ 
    ByRef pvt_Object As PivotTable _ 
        ) 
 
    Dim lng_MissingFlag As Long 
 
    lng_MissingFlag = pvt_Object.PivotCache.MissingItemsLimit 
 
    pvt_Object.PivotCache.MissingItemsLimit = xlMissingItemsNone 
    pvt_Object.PivotCache.Refresh 
    pvt_Object.PivotCache.MissingItemsLimit = lng_MissingFlag 
 
End Function 
 
Function PivotAutofilter( _ 
    ByRef pvt_Object As PivotTable, _ 
    Optional ByRef str_ColumnName As String = "", _ 
    Optional ByRef arr_CriteriaInput, _ 
    Optional ByVal InvertedAutoFilter As e_FilterType = e_ShowOnlyCriteriaItems, _ 
    Optional ByVal RemoveAutoFilter As e_ResetFilter = e_LeaveCurrentFilters _ 
    ) 
                             
 
    Dim arr_Criteria 
    arr_Criteria = z_mCovertToSimpleArray(arr_CriteriaInput) 
     
    Dim str_ValuePart As String 
     
    If Not IsEmpty(arr_Criteria) Then 
        str_ValuePart = " value" & IIf(UBound(arr_Criteria) > 0, "s: ", ": ") & arr_Criteria(0) & IIf(UBound(arr_Criteria) > 0, " ... ", "") 
    End If 
     
    Application.Statusbar = Mid("Filtering pivot table: " & pvt_Object.Name & " by field " & str_ColumnName & str_ValuePart, 1, 255) 
 
    'Clear filters in pivot table 
    If (str_ColumnName = "" Or RemoveAutoFilter = e_ClearCurrentFilters) And Not pvt_Object.PivotCache.OLAP Then 
     
        Dim pvf_Cursor As PivotField 
     
        For Each pvf_Cursor In pvt_Object.PivotFields 
            pvf_Cursor.EnableMultiplePageItems = True 
            pvf_Cursor.ClearAllFilters 
        Next 
     
        If RemoveAutoFilter <> e_ClearCurrentFilters Then Exit Function 
     
    End If 
 
    'Do pivot Filter 
    With pvt_Object.PivotFields(str_ColumnName) 
 
        If Not pvt_Object.PivotCache.OLAP Then 
            .EnableMultiplePageItems = True 
            .ClearAllFilters 
        End If 
 
        'just Clear pivot Filter 
        If IsMissing(arr_CriteriaInput) Then Exit Function 
 
        If InvertedAutoFilter Then 
 
            Dim i& 'array index 
             
            For i = LBound(arr_Criteria) To UBound(arr_Criteria) 
                On Error Resume Next 
                    .PivotItems(CStr(arr_Criteria(i))).Visible = False 
                On Error GoTo 0 
            Next 
             
        Else 
            If pvt_Object.PivotCache.OLAP Then 
                If .Orientation = xlPageField Then 
                    .CurrentPageName = arr_Criteria(0) 
                Else 
                    .VisibleItemsList = arr_Criteria 
                End If 
            Else 
             
                Dim pvi_Cursor As PivotItem 'pivot item 
                Dim cc_HiddenItems As Collection 
                Call z_mFillCollection(arr_Criteria, cc_HiddenItems) 
     
                On Error GoTo err1004: 
     
                For Each pvi_Cursor In .PivotItems 
                    pvi_Cursor.Visible = CLng(z_mKeyExists(cc_HiddenItems, pvi_Cursor.Name)) 
                Next 
     
            End If 
 
    Exit Function 
 
err1004: 
        If Err.Number = 1004 Then 'empty data error skip 
            On Error GoTo 0 
            Resume Next 
        End If 
         
        End If 
     
    End With 
     
End Function 
 
Function PivotAutofilterDate( _ 
    ByRef pvt_Object As PivotTable, _ 
    ByRef str_ColumnName As String, _ 
    Optional ByRef arr_CriteriaInput As Variant, _
    Optional ByVal InvertedAutoFilter As e_FilterType = e_ShowOnlyCriteriaItems, _ 
    Optional ByVal RemoveAutoFilter As e_ResetFilter = e_LeaveCurrentFilters _ 
    ) 
                                    
    'Application.ScreenUpdating = False 
                                 
    Dim i& 'array index 
     
    'check if filter list is not array 
    'If TypeName(arr_Criteria) = "Range" Then arr_Criteria = RangeToArray(arr_Criteria) 
    Dim arr_Criteria
    arr_Criteria = z_mCovertToSimpleArray(arr_CriteriaInput, True)
     
    'Criteria Range To Date 
'    For i = LBound(arr_Criteria) To UBound(arr_Criteria)
'       arr_Criteria(i) = CStr(Month(arr_Criteria(i)) & "/" & Day(arr_Criteria(i)) & "/" & Year(arr_Criteria(i)))
'    Next
 
    With pvt_Object.PivotFields(str_ColumnName) 
     
        .EnableMultiplePageItems = True 
        .ClearAllFilters 
 
        If InvertedAutoFilter Then 
                        
            For i = LBound(arr_Criteria) To UBound(arr_Criteria) 
                .PivotItems(arr_Criteria(i)).Visible = False 
            Next 
             
        Else 
            Dim pvi_Cursor As PivotItem 'pivot item 
         
            Dim cc_HiddenItems As Collection 
            Call z_mFillCollection(arr_Criteria, cc_HiddenItems) 
         
            For Each pvi_Cursor In .PivotItems 
                With pvi_Cursor 
                    .Visible = z_mKeyExists(cc_HiddenItems, .Name) 'z_mFoundInArray(arr_Criteria, pvi_Cursor.Name) 
                End With 
            Next 
 
        End If 
     
    End With 
     
    'Application.ScreenUpdating = True 
End Function 
 
Private Function z_mFoundInArray( _ 
    ByRef arr_Criteria, _ 
    ByRef var_Criteria _ 
    ) As Boolean 
 
    z_mFoundInArray = True 
 
On Error GoTo NOT_FOUND 
   Call WorksheetFunction.Match(var_Criteria, arr_Criteria, 0) 
    
NOT_FOUND: 
    If Err.Number = 1004 Then 
        z_mFoundInArray = False 
        Resume Next 
    End If 
 
End Function 
 
Function EmailCopyFromDraft( _ 
    ByRef obj_Email As Object, _ 
    ByRef str_DraftName As String, _ 
    Optional str_Subject As String, _ 
    Optional str_To As String, _ 
    Optional str_Cc As String, _ 
    Optional str_Bcc As String _ 
    ) As Boolean 
 
    Dim obj_OutlookFolder As Object 
    Dim obj_DraftEmail As Object 
     
    Call z_mOutlookGetDefaultFolder(obj_OutlookFolder, e_Drafts) 
    Call OutlookFindEmailInFolder(obj_DraftEmail, obj_OutlookFolder, str_DraftName) 
 
 
    Set obj_Email = obj_DraftEmail.Copy 
 
    With obj_Email 
        If str_Subject <> "" Then 
            .Subject = str_Subject 
        Else 
            .Subject = Mid(.Subject, 7, Len(.Subject)) 
        End If 
         
        If str_To <> "" Then .To = str_To 
        If str_Cc <> "" Then .cc = str_Cc 
        If str_Bcc <> "" Then .Bcc = str_Bcc 
     
        .Save 
        .display 
     
    End With 
 
End Function 
 
Function EmailRangePaste( _ 
    ByRef obj_Email As Object, _ 
    ByRef str_Identifier As String, _ 
    ByRef rng_ToCopy As Object, _ 
    Optional ByRef lng_PasteType As e_WordPasteAs = e_Original) 
  
On Error GoTo Err 
  
    Dim obj_Selection As Object 'Word Selection 
 
    Set obj_Selection = obj_Email.GetInspector.WordEditor.Windows(1).Selection 
     
    With obj_Selection.Find 
        .ClearFormatting 
        .MatchCase = True 
        .Execute FindText:=str_Identifier 
    End With 
     
    Select Case TypeName(rng_ToCopy) 
     
        Case "ListObject" 
            rng_ToCopy.Range.Copy 
        Case "Range" 
    rng_ToCopy.Copy 
             
    End Select 
     
     
    Select Case lng_PasteType 
        
        Case e_Original 
    obj_Selection.Paste 
   
        Case e_Picture 
            Call obj_Selection.Range.PasteAndFormat(lng_PasteType) 
             
    End Select 
     
    
    DoEvents 
    DoEvents 
         
    obj_Email.Save 
     
    Exit Function 
 
Err: 
    Debug.Assert False: Resume 
  
End Function 
  
  
Function EmailTableFill( _ 
    ByRef obj_Email As Object, _ 
    ByRef tbl_Source As ListObject, _ 
    ByRef lng_Identifier As Long _ 
    ) 
 
On Error GoTo Err 
 
    Dim arr_QueryHeader() 
    Dim arr_QueryResult() As String 
    Dim lng_RowCounter As Long 
    Dim lng_ColumnsCount As Long 
    Dim rng_Cursor As Range 
    Dim c_CopyCursor As Long 
        
    arr_QueryHeader = tbl_Source.HeaderRowRange.Value 
 
    lng_ColumnsCount = tbl_Source.HeaderRowRange.Count 
 
    ReDim arr_QueryResult(1 To lng_ColumnsCount, 1 To 1) 
 
    lng_RowCounter = 1 
     
'DATA ROW TO ARRAY 
    'On Error GoTo EMPTY_TABLE: 
 
    For Each rng_Cursor In tbl_Source.DataBodyRange.Resize(, 1) 
        
        If rng_Cursor.EntireRow.Hidden = False Then 
            For c_CopyCursor = 1 To lng_ColumnsCount 
                arr_QueryResult(c_CopyCursor, lng_RowCounter) = rng_Cursor.Offset(, c_CopyCursor - 1).Text 
            Next 
             
            ReDim Preserve arr_QueryResult(1 To lng_ColumnsCount, 1 To UBound(arr_QueryResult, 2) + 1) 
            lng_RowCounter = lng_RowCounter + 1 
        End If 
 
    Next 
     
    Dim bool_EmptySourceTable As Boolean 
     
    If UBound(arr_QueryResult, 2) = 1 Then 
        bool_EmptySourceTable = True 
    Else 
        ReDim Preserve arr_QueryResult(1 To lng_ColumnsCount, 1 To UBound(arr_QueryResult, 2) - 1) 
    End If 
     
     
'TOTAL ROW EXIST 
    If Not tbl_Source.TotalsRowRange Is Nothing Then 
        Dim arr_TotalRowData 
        ReDim arr_TotalRowData(1 To lng_ColumnsCount, 1 To 1) 
 
        c_CopyCursor = 1 
        For Each rng_Cursor In tbl_Source.TotalsRowRange 
            arr_TotalRowData(c_CopyCursor, 1) = rng_Cursor.Text 
            c_CopyCursor = c_CopyCursor + 1 
        Next 
    End If 
     
    
 
'OUTOLOOK PART 
    Dim obj_OutlookFolder As Object 
    Dim obj_NewEmail As Object 
    Dim str_HtmlBody$ 
 
    With obj_Email 
 
        str_HtmlBody = .htmlBody 
         
        Dim str_ListLine$ 
        Dim str_ListFirstText$ 
        Dim str_BodyBegining$ 
        Dim str_BodyEnd$ 
        Dim lng_ListBeginning& 
        Dim lng_ListEnd& 
        Dim str_FinalList$ 
        Dim str_TempLine$ 
     
        'str_ListFirstText = "[" & tbl_Source.HeaderRowRange.Resize(1, 1).Value & "]" & lng_Identifier 
     
        'find first item 
        Dim rng_HeaderCursor As Range 
         
        For Each rng_HeaderCursor In tbl_Source.HeaderRowRange 
            str_ListFirstText = "[" & rng_HeaderCursor.Text & "]" & lng_Identifier 
             
            If InStr(1, str_HtmlBody, str_ListFirstText) <> 0 Then 
                Exit For 
            End If 
     
        Next 
     
        If str_ListFirstText = "" Then Exit Function 
     
        lng_ListBeginning = InStrRev(str_HtmlBody, "<tr ", InStr(1, str_HtmlBody, str_ListFirstText))  'find end of paragraph text afte end first line tag 
         
         
        lng_ListEnd = InStr(InStr(1, str_HtmlBody, str_ListFirstText), str_HtmlBody, "</tr>") + 5 
 
        str_BodyBegining = Mid(str_HtmlBody, 1, lng_ListBeginning - 1) 
        str_BodyEnd = Mid(str_HtmlBody, lng_ListEnd) 
 
        str_ListLine = Mid(str_HtmlBody, lng_ListBeginning, lng_ListEnd - lng_ListBeginning) 
 
'If no data in table skip loop part this will remove whole row from table in email 
        If bool_EmptySourceTable Then GoTo ENTERING_BODY 
 
        Dim r_Cursor As Long 
        Dim c_Cursor As Long 
         
        For r_Cursor = 1 To UBound(arr_QueryResult, 2) 
 
            For c_Cursor = LBound(arr_QueryHeader, 2) To UBound(arr_QueryHeader, 2) 
 
                If IsNull(arr_QueryResult(c_Cursor, r_Cursor)) Then arr_QueryResult(c_Cursor, r_Cursor) = "" 
 
                If c_Cursor = 1 Then 
                    str_TempLine = Replace(str_ListLine, "[" & arr_QueryHeader(1, c_Cursor) & "]" & lng_Identifier, arr_QueryResult(c_Cursor, r_Cursor)) 
                Else 
                    str_TempLine = Replace(str_TempLine, "[" & arr_QueryHeader(1, c_Cursor) & "]" & lng_Identifier, arr_QueryResult(c_Cursor, r_Cursor)) 
                End If 
            Next 
            str_FinalList = str_FinalList & str_TempLine 
        Next 
 
        If Not tbl_Source.TotalsRowRange Is Nothing Then 
            For c_Cursor = LBound(arr_QueryHeader, 2) To UBound(arr_QueryHeader, 2) 
                str_BodyEnd = Replace(str_BodyEnd, "[" & arr_QueryHeader(1, c_Cursor) & "]Total" & lng_Identifier, arr_TotalRowData(c_Cursor, 1)) 
            Next 
        End If 
         
ENTERING_BODY: 
         
        str_HtmlBody = str_BodyBegining & str_FinalList & str_BodyEnd 
 
        .htmlBody = str_HtmlBody 
         
        .Save 
    End With 
 
Exit Function 
Err: 
Debug.Assert False: Resume 
 
End Function 
 
Function EmailCreateNew( _ 
    ByRef obj_Email As Object, _ 
    Optional ByVal var_Body, _ 
    Optional ByVal str_Subject, _ 
    Optional ByVal var_To, _ 
    Optional ByVal var_Cc, _ 
    Optional ByVal var_Bcc _ 
    ) As Boolean 
 
    , 
    'Dim app_Outlook As Object 'Outlook.Application 
 
     
    'Setup Outlook 
    'Set app_Outlook = GetObject(, "Outlook.Application") 'Set app_Outlook = Outlook.Application 
    Call z_mLinkOutlook 
     
    Set obj_Email = app_Outlook.CreateItem(0) 
     
    Dim str_HtmlBody$ 
     
    With obj_Email 
 
        On Error Resume Next 
        If Not IsMissing(var_To) Then .To = Join(z_mCovertToSimpleArray(var_To), ";") 
        If Not IsMissing(var_Cc) Then .cc = Join(z_mCovertToSimpleArray(var_Cc), ";") 
        If Not IsMissing(var_Bcc) Then .Bcc = Join(z_mCovertToSimpleArray(var_Bcc), ";") 
        On Error GoTo 0 
         
        If Not IsMissing(str_Subject) Then .Subject = Join(z_mCovertToSimpleArray(str_Subject), "<br>") 
 
        .display 
        str_HtmlBody = .htmlBody 
         
        'Split mailBody Mail Body on part before empty line and part after empty line 
        Dim str_BodyStart$ 
        Dim str_BodyEnd$ 
        Dim lng_BodyText& 
         
        Const str_EmptyLineText$ = "<p class=MsoNormal><o:p>&nbsp;</o:p></p>" 'empty line of the body 
         
        lng_BodyText = InStr(1, str_HtmlBody, str_EmptyLineText) 
     
        str_BodyStart = Mid(str_HtmlBody, 1, lng_BodyText - 1) 
        str_BodyEnd = Mid(str_HtmlBody, lng_BodyText + Len(str_EmptyLineText), Len(str_HtmlBody)) 
 
        If Not IsMissing(var_Body) Then 
            .htmlBody = str_BodyStart & "<p class=MsoNormal>" & Join(z_mCovertToSimpleArray(var_Body), "<br>") & "<br>" & "</o:p></p>" & str_BodyEnd 
        End If 
         
        .Save 
         
    End With 
     
    Call obj_Email.Recipients.ResolveAll 
     
End Function 
 
 
Function xDebugPrintTableLinker( _ 
                                Optional wbk_Source As Workbook = Nothing _ 
                                ) 
 
    If wbk_Source Is Nothing Then Set wbk_Source = ThisWorkbook 
     
    Dim sht_Cursor As Worksheet 
    Dim tbl_Cursor As ListObject 
    Dim pvt_Cursor As PivotTable 
    Dim name_Cursor As Name 
     
    Dim arr_Dimensions_Tables() 
    Dim arr_Initialization_Tables() 
     
    Dim arr_Dimensions_Pivots() 
    Dim arr_Initialization_Pivots() 
     
    Dim arr_Dimensions_Ranges() 
    Dim arr_Initialization_Ranges() 
     
    ReDim arr_Dimensions_Tables(0) 
    ReDim arr_Initialization_Tables(0) 
     
    ReDim arr_Dimensions_Pivots(0) 
    ReDim arr_Initialization_Pivots(0) 
     
    ReDim arr_Dimensions_Ranges(0) 
    ReDim arr_Initialization_Ranges(0) 
     
     
    For Each name_Cursor In wbk_Source.Names 
        If name_Cursor.Name Like "rng_*" Then 
             
            arr_Dimensions_Ranges(UBound(arr_Dimensions_Ranges)) = "Dim " & name_Cursor.Name & " As Range" 
            arr_Initialization_Ranges(UBound(arr_Initialization_Ranges)) = vbTab & "Call shtATk.RangeCreateLink(" & name_Cursor.Name & ",""" & name_Cursor.Name & """)" 
             
     
            ReDim Preserve arr_Dimensions_Ranges(UBound(arr_Dimensions_Ranges) + 1) 
            ReDim Preserve arr_Initialization_Ranges(UBound(arr_Initialization_Ranges) + 1) 
     
        End If 
    Next 
     
    For Each sht_Cursor In wbk_Source.Worksheets 
        For Each tbl_Cursor In sht_Cursor.ListObjects 
         
            If tbl_Cursor.Name Like "t_*" Then 
                arr_Dimensions_Tables(UBound(arr_Dimensions_Tables)) = "Dim " & tbl_Cursor.Name & " As ListObject" 
                arr_Initialization_Tables(UBound(arr_Initialization_Tables)) = vbTab & "Call shtATk.TableCreateLink(" & tbl_Cursor.Name & ",""" & tbl_Cursor.Name & """)" 
                 
                ReDim Preserve arr_Dimensions_Tables(UBound(arr_Dimensions_Tables) + 1) 
                ReDim Preserve arr_Initialization_Tables(UBound(arr_Initialization_Tables) + 1) 
            End If 
        Next 
         
        For Each pvt_Cursor In sht_Cursor.PivotTables 
             
            If pvt_Cursor.Name Like "pvt_*" Then 
                arr_Dimensions_Pivots(UBound(arr_Dimensions_Pivots)) = "Dim " & pvt_Cursor.Name & " As PivotTable" 
                arr_Initialization_Pivots(UBound(arr_Initialization_Pivots)) = vbTab & "Call shtATk.PivotCreateLink(" & pvt_Cursor.Name & ",""" & pvt_Cursor.Name & """)" 
                 
                ReDim Preserve arr_Dimensions_Pivots(UBound(arr_Dimensions_Pivots) + 1) 
                ReDim Preserve arr_Initialization_Pivots(UBound(arr_Initialization_Pivots) + 1) 
            End If 
        Next 
 
    Next 
 
'PRINTING RESULTS 
 
    Debug.Print "'---- LINKER PART START --- Enter On the top of the routine module ----" & vbCrLf 
    Debug.Print Join(arr_Dimensions_Ranges, vbCrLf) 
    Debug.Print Join(arr_Dimensions_Tables, vbCrLf) 
    Debug.Print Join(arr_Dimensions_Pivots, vbCrLf) 
  
    Debug.Print "Function LinkWorkbookTables()" & vbCrLf 
 
    If UBound(arr_Dimensions_Ranges) > 0 Then 
        Debug.Print "'Global Ranges Part" & vbCrLf 
        Debug.Print Join(arr_Initialization_Ranges, vbCrLf) 
    End If 
 
    If UBound(arr_Dimensions_Tables) > 0 Then 
        Debug.Print "'Tables Part" & vbCrLf 
        Debug.Print Join(arr_Initialization_Tables, vbCrLf) 
    End If 
     
    If UBound(arr_Dimensions_Pivots) > 0 Then 
        Debug.Print "'Pivots Part" & vbCrLf 
        Debug.Print Join(arr_Initialization_Pivots, vbCrLf) 
    End If 
     
    Debug.Print "End Function" & vbCrLf 
    Debug.Print "'---- LINKER PART END ----" 
End Function 
 
Function TableCreateLink( _ 
    ByRef obj_Table As ListObject, _ 
    ByVal str_TableName As String, _ 
    Optional ByVal wbk_Source As Workbook _ 
    ) As Boolean 
         
    If wbk_Source Is Nothing Then Set wbk_Source = ThisWorkbook 
         
    Dim sht_Cursor As Worksheet 
         
    On Error Resume Next 
         
    For Each sht_Cursor In wbk_Source.Worksheets 
     
       Set obj_Table = sht_Cursor.ListObjects(str_TableName) 
  
    Next 
     
    If obj_Table Is Nothing Then 
        TableCreateLink = False 
    Else 
        TableCreateLink = True 
    End If 
 
End Function 
 
Function PivotCreateLink( _ 
    ByRef obj_PivotTable As PivotTable, _ 
    ByVal str_TableName As String, _ 
    Optional ByVal wbk_Source As Workbook _ 
    ) As Boolean 
         
    If wbk_Source Is Nothing Then Set wbk_Source = ThisWorkbook 
         
    Dim sht_Cursor As Worksheet 
         
    On Error Resume Next 
         
    For Each sht_Cursor In wbk_Source.Worksheets 
     
       Set obj_PivotTable = sht_Cursor.PivotTables(str_TableName) 
  
    Next 
     
    If obj_PivotTable Is Nothing Then 
        PivotCreateLink = False 
    Else 
        PivotCreateLink = True 
    End If 
 
End Function 
 
 
  
Function RangeCreateLink( _ 
    ByRef obj_Range As Range, _ 
    ByVal str_RangeName As String, _ 
    Optional ByVal wbk_Source As Workbook _ 
    ) As Boolean 
         
    If wbk_Source Is Nothing Then Set wbk_Source = ThisWorkbook 
         
    Dim sht_Cursor As Worksheet 
         
    On Error Resume Next 
 
    Dim arr_Address 
 
    If CStr(wbk_Source.Names(str_RangeName).RefersTo) Like "*!*" Then 
        arr_Address = Split(CStr(wbk_Source.Names(str_RangeName).RefersTo), "!") 
     
        If arr_Address(0) Like "='* *'" Then 
            arr_Address(0) = Mid(Left(arr_Address(0), Len(arr_Address(0)) - 1), 3) 
        Else 
            arr_Address(0) = Mid(arr_Address(0), 2) 
        End If 
         
        Set obj_Range = wbk_Source.Worksheets(arr_Address(0)).Range(arr_Address(1)) 
 
    Else 
        'TDB Linked on TABLE Or formula 
        If CStr(wbk_Source.Names(str_RangeName).RefersTo) Like "=*[[]*[]]" Then 
         
            Dim str_TableName As String 
     
            str_TableName = Mid(Split(CStr(wbk_Source.Names(str_RangeName).RefersTo), "[")(0), 2) 
 
            'Debug.Print wbk_Source.Names(str_RangeName).RefersTo 
      
            Dim ws As Worksheet 
            Dim LO As ListObject 
            Dim shName As String 
            Dim FindTableSheet As String 
         
            For Each ws In ThisWorkbook.Sheets 
                On Error Resume Next 
                Set LO = ws.ListObjects(str_TableName) 
                If Err.Number = 0 Then 
                    FindTableSheet = ws.Name 
                    Exit For 
                Else 
                    Err.Clear 
                    FindTableSheet = "Not Found" 
    End If 
                On Error GoTo 0 
            Next ws 
     
            If FindTableSheet <> "Not Found" Then 
                Set obj_Range = wbk_Source.Worksheets(FindTableSheet).Range(CStr(wbk_Source.Names(str_RangeName).RefersTo)) 
            End If 
     
        End If 
      
    End If 
      
      
    RangeCreateLink = Not (obj_Range Is Nothing) 
 
End Function 
 
 
Function EmailCopyFromFile( _ 
    ByRef obj_Email As Object, _ 
    Optional ByVal str_FileName As String = "", _ 
    Optional ByVal str_DirPath As String = "", _ 
    Optional ByVal str_Subject, _ 
    Optional var_To, _ 
    Optional var_Cc, _ 
    Optional var_Bcc _ 
    ) As Boolean 
  
 
    On Error GoTo err_AtkError 
   
    EmailCopyFromFile = False 
     
    Dim strTempDate$ 
    Dim str_FinalPath As String 
 
    If z_mFilePathChecker(str_FinalPath, str_FileName, str_DirPath, "Please select Outlook email file", "*.msg") = False Then Exit Function 
     
    'Setup Outlook 
    Call z_mLinkOutlook 
     
    Dim obj_DefaultFolder As Object 
    Call z_mOutlookGetDefaultFolder(obj_DefaultFolder, e_Drafts) 
     
    Dim obj_EmailTemplate As Object 
    Set obj_EmailTemplate = app_Outlook.CreateItemFromTemplate(str_FinalPath, obj_DefaultFolder) 
     
    EmailCopyFromFile = True 
     
    obj_EmailTemplate.display 
    DoEvents 
     
    'To correct language encoding is neccerary create new copy of email 
    Dim obj_EmailTemp As Object 
    Set obj_EmailTemp = obj_EmailTemplate.Copy 
     
    With obj_EmailTemp 
     
        On Error Resume Next 
            If Not IsMissing(var_To) Then .To = Join(z_mCovertToSimpleArray(var_To), ";") 
            If Not IsMissing(var_Cc) Then .cc = Join(z_mCovertToSimpleArray(var_Cc), ";") 
            If Not IsMissing(var_Bcc) Then .Bcc = Join(z_mCovertToSimpleArray(var_Bcc), ";") 
            If Not IsMissing(str_Subject) Then .Subject = Join(z_mCovertToSimpleArray(str_Subject), ";") 
        On Error GoTo 0 
          
         
        .Save 
        .display 
         
    End With 
     
    obj_EmailTemplate.Delete 
    DoEvents 
      
    'Resolve recipients 
    Call obj_EmailTemp.Recipients.ResolveAll 
      
    Set obj_Email = obj_EmailTemp 
     
    EmailCopyFromFile = True 
      
    Set obj_EmailTemp = Nothing 
     
     
exit_ok: 
    EmailCopyFromFile = True 
 
Exit Function 
err_AtkError: 
     
    If frm_GlErr.z_Show(Err) Then 
        Stop: Resume 
    Else 
        Resume err_AtkError 
    End If 
 
      
End Function 
 
Function EmailSend( _ 
        ByRef obj_Email As Object _ 
    ) 
 
    'Get Focus for Email Window 
    obj_Email.display 
     
    'Wait 1 Second 
    Application.Wait (Now + TimeValue("0:00:01")) 
     
    'Press CTRL+Enter to sent shortcut 
    SendKeys "^{ENTER}" 
 
End Function 
         
 
 
 
Function EmailReplaceText( _ 
    ByRef obj_Email As Object, _ 
    ByVal var_TextToReplace As Variant, _ 
    ByVal var_ReplaceWith As Variant _ 
    ) 
     
    Dim str_HtmlBody$ 
    Dim i_TextCursor As Long 
 
    var_TextToReplace = z_mCovertToSimpleArray(var_TextToReplace) 
    var_ReplaceWith = z_mCovertToSimpleArray(var_ReplaceWith) 
 
    If UBound(var_ReplaceWith) <> UBound(var_TextToReplace) Then Exit Function 
 
    With obj_Email 
     
        For i_TextCursor = LBound(var_ReplaceWith) To UBound(var_ReplaceWith) 
            .htmlBody = Replace(.htmlBody, CStr(var_TextToReplace(i_TextCursor)), CStr(var_ReplaceWith(i_TextCursor))) 
            .Subject = Replace(.Subject, CStr(var_TextToReplace(i_TextCursor)), CStr(var_ReplaceWith(i_TextCursor))) 
        Next 
         
        .Save 
        .Display 
    End With 
 
End Function 
 
Function EmailAttachWorkbook( _ 
    ByRef obj_Email As Object, _ 
    Optional ByVal wbk_Source As Workbook, _ 
    Optional ByVal str_WorkbookName As String _ 
    ) As Boolean 
    
    'Create Temporary file neccecary if different file name TDB 
    If wbk_Source Is Nothing Then 
        Set wbk_Source = ThisWorkbook 
    End If 
 
    wbk_Source.Save 
     
    Dim obj_Attachment As Object 
 
    With obj_Email 
        Set obj_Attachment = .Attachments.Add(wbk_Source.FullName) 
         
        If str_WorkbookName <> "" Then 
     
            Dim str_CurrentName As String 
            Dim str_NewName As String 
             
            'Set outmail.Attachments(1) = atmt 
            str_CurrentName = obj_Attachment.Filename                                 ' Get Name Of File Attached 
   
            str_NewName = str_WorkbookName & Mid(str_CurrentName, InStrRev(str_CurrentName, ".")) 
 
            Call obj_Attachment.SaveAsFile(Environ("Temp") & "\" & str_NewName)                           ' Save as NewFileName 
            Call .Attachments.Add(Environ("Temp") & "\" & str_NewName)              ' Attach the file with a new name 
 
            On Error Resume Next 
                Kill Environ("Temp") & "\" & str_NewName 
            On Error GoTo 0 
 
            DoEvents 
            obj_Attachment.Delete                   ' Delete the file with the old name 
            DoEvents 
     
        End If 
 
        .Save 
     
    End With 
      
End Function 
 
Function EmailAttachFile( _ 
    ByRef obj_Email As Object, _ 
    Optional ByRef str_DirPath As String, _ 
    Optional ByRef str_FileName As String, _ 
    Optional str_WindowTitle As String _ 
    ) As Boolean 
   
    EmailAttachFile = False 
     
    Dim strTempDate$ 
     
    'CHECK FILEPATH EXIST 
     
    If str_DirPath = "" Then 
        str_DirPath = ThisWorkbook.Path & "\" 
    End If 
       
    ChDir str_DirPath 
       
    If str_WindowTitle = "" Then 
        str_WindowTitle = "Please select file " & str_FileName 
    End If 
       
SELECT_FILE: 
       
    If str_FileName = "" Then 
       str_FileName = Application.GetOpenFilename(Title:=str_WindowTitle) 
    Else 
        If Not str_DirPath Like "*\" Then str_DirPath = str_DirPath & "\" 
         
        'CHECK IF FILENAME EXIST 
         
        str_FileName = str_DirPath & str_FileName 
    End If 
     
    'exit routine when file is empty 
    If str_FileName = "False" Then 
        Select Case MsgBox("File was not selected try again?", vbCritical + vbYesNo, "File Selection problem") 
            Case vbYes: str_FileName = "": GoTo SELECT_FILE 
            Case vbNo: Exit Function 
        End Select 
    End If 
     
     
    With obj_Email 
        .Attachments.Add str_FileName 
        .Save 
    End With 
      
    EmailAttachFile = True 
      
End Function 
 
 
Function EmailSaveAttachment( _ 
    ByRef obj_Email As Object, _ 
    Optional ByVal str_SavePath As Variant, _ 
    Optional ByVal str_AttachmentFileMask As Variant, _ 
    Optional ByVal str_FileName As Variant, _ 
    Optional ByVal lng_NameAction As e_FileNameAction = e_Prefix _ 
    ) 
 
 
    If obj_Email.Attachments.Count = 0 Then Exit Function 
 
    If IsMissing(str_SavePath) Then str_SavePath = ThisWorkbook.Path 
 
    Dim objAtt As Object 
    Dim str_OriginalName As String 
    
     
    For Each objAtt In obj_Email.Attachments 
     
        Dim str_AttFilename As String 
     
        str_AttFilename = objAtt.Filename 
        str_OriginalName = objAtt.DisplayName 
     
        If Not IsMissing(str_AttachmentFileMask) Then 
            If Not str_AttFilename Like str_AttachmentFileMask Then 
                GoTo NEXT_ATTACHEMENT 
            End If 
        End If 
     
     
        If Not IsMissing(str_FileName) Then 
            Select Case lng_NameAction 
                Case e_Prefix: str_OriginalName = str_FileName + str_AttFilename 
                Case e_Suffix:  str_OriginalName = str_OriginalName + str_FileName + Split(str_AttFilename, ".")(1) 
                Case e_Replace: str_OriginalName = Replace(str_AttFilename, str_OriginalName, str_FileName) 
            End Select 
        End If 
     
        objAtt.SaveAsFile str_SavePath & "\" & str_OriginalName 
         
NEXT_ATTACHEMENT: 
        Set objAtt = Nothing 
    Next 
 
End Function 
 
 
Private Function z_mOutlookGetDefaultFolder( _ 
    ByRef obj_OutlookFolder As Object, _ 
    ByVal DefaultFolderType As OutlookFolders _ 
    ) As Boolean 
     
     
    Dim lng_TryCounter As Long 
    lng_TryCounter = 0 
     
    Do Until lng_TryCounter = 6 
         
        DoEvents 
    Set obj_OutlookFolder = app_Outlook.GetNamespace("MAPI").GetDefaultFolder(16) '16 = olFolderDrafts 
        DoEvents 
          
        If obj_OutlookFolder Is Nothing Then 
            lng_TryCounter = lng_TryCounter + 1 
        Else 
            lng_TryCounter = 6 
        End If 
         
    Loop 
     
End Function 
 
Public Function EmailFlag(ByRef obj_Email As Object, Optional ByVal lng_Flag As e_OutlookFlag = e_Complete) 
 
    obj_Email.flagStatus = lng_Flag 
    obj_Email.Save 
 
End Function 
 
 
Private Function z_mEmailFindInOutlookFolder( _ 
    ByRef obj_Email As Object, _ 
    ByRef obj_Folder As Object, _ 
    ByRef str_Subject$ _ 
    ) As Boolean 
 
    Dim lng_MailsCount& 
    Dim lng_MailCursor& 
     
    Dim bool_MailFound As Boolean 
    'Find Draft Email 
     
    OutlookFindEmailInFolder = False 
     
    lng_MailsCount = obj_Folder.items.Count 
     
    lng_MailCursor = 1 
    Do Until lng_MailCursor > lng_MailsCount 
       
        If obj_Folder.items.Item(lng_MailCursor).Subject Like str_Subject Then 
            bool_MailFound = True 
            Exit Do 
        End If 
         
        lng_MailCursor = lng_MailCursor + 1 
    Loop 
     
    If bool_MailFound Then 
        OutlookFindEmailInFolder = True 
        Set obj_Email = obj_Folder.items.Item(lng_MailCursor) 
    End If 
 
End Function 
 
Private Function z_mFillCollection( _ 
    ByRef arr_List, _ 
    ByRef cc_Container As Collection _ 
    ) 
 
    Set cc_Container = New Collection 
     
    Dim i& 
     
    For i = LBound(arr_List) To UBound(arr_List) 
       Call cc_Container.Add(arr_List(i), CStr(arr_List(i))) 
    Next 
     
End Function 
 
Private Function z_mKeyExists( _ 
    ByRef cc_Container As Collection, _ 
    ByVal strKey As String _ 
    ) As Boolean 
    On Error GoTo Exists_Err 
    'Dim strKey$: strKey = CStr(varKey) 
    Dim lngType&: lngType = VarType(cc_Container.Item(strKey)) 
    z_mKeyExists = True 
Exit_Function: 
    lngType = 0 
    Exit Function 
Exists_Err: 
    If Err.Number = 9 Or Err.Number = 5 Then z_mKeyExists = False 
End Function 
 
Function TextToColumnsDelimeted( _ 
    ByVal rngColumns As Range, _ 
    Optional ByVal arr_FieldType As Variant, _ 
    Optional ByRef str_Delimeter$ = "", _ 
    Optional ByRef TextQualifier As XlTextQualifier, _ 
    Optional str_DecimalSeparator$ = "", _ 
    Optional str_ThousandSeparator$ = "" _ 
    ) 
    Dim arr_FieldTypesResult 
 
    Dim bool_Semicolon As Boolean 
    Dim bool_Comma As Boolean 
    Dim bool_Space As Boolean 
    Dim bool_Tab As Boolean 
    Dim bool_Other As Boolean 
     
    Dim str_OtherChar As String 
 
    If IsMissing(arr_FieldType) Then 
        arr_FieldTypesResult = Array(1, xlGeneralFormat) 
    Else 
        'prepare column data type 
        If UBound(arr_FieldType) = 0 Then 
            arr_FieldTypesResult = Array(1, arr_FieldType(0)) 
        Else 
            ReDim arr_FieldTypesResult(1 To UBound(arr_FieldType) + 1, 9) 
            Dim i_FieldCursor& 
             
            For i_FieldCursor = 0 To UBound(arr_FieldType) 
                arr_FieldTypesResult(i_FieldCursor) = Array(i_FieldCursor + 1, arr_FieldType(i_FieldCursor)) 
            Next 
             
        End If 
    End If 
     
 
     Select Case str_Delimeter 
         Case "" 
         Case "Semicolon": bool_Semicolon = True 
         Case "Comma": bool_Comma = True 
         Case "Space": bool_Space = True 
         Case "Tab": bool_Tab = True 
         Case Else 
             bool_Other = True 
             str_OtherChar = str_Delimeter 
     End Select 
 
    Dim rng_Cursor As Range 
    
    For Each rng_Cursor In rngColumns.Resize(1) 
 
        rng_Cursor.EntireColumn.TextToColumns _ 
            Destination:=rng_Cursor.EntireColumn, _ 
            DataType:=xlDelimited, _ 
            TextQualifier:=xlDoubleQuote, _ 
            ConsecutiveDelimiter:=False, _ 
            Tab:=bool_Tab, _ 
            Semicolon:=bool_Semicolon, _ 
            Comma:=bool_Comma, _ 
            Space:=bool_Space, _ 
            Other:=bool_Other, _ 
            OtherChar:=str_OtherChar, _ 
            FieldInfo:=arr_FieldTypesResult, _ 
            TrailingMinusNumbers:=True 
  
    Next 
 
End Function 
 
Function TableCreateFromRange( _ 
    ByRef tbl_NewTable As ListObject, _ 
    ByVal rng_SourceRange As Range, _ 
    Optional ByVal lng_HeaderRow As Long = 1, _ 
    Optional ByVal str_TableName As String _ 
    ) 
 
    If lng_HeaderRow < 1 Then Exit Function 
     
    If lng_HeaderRow > 1 Then 
        Set rng_SourceRange = rng_SourceRange.CurrentRegion.Offset(lng_HeaderRow - 1).Resize(rng_SourceRange.Rows.Count - lng_HeaderRow) 
    Else 
        Set rng_SourceRange = rng_SourceRange.CurrentRegion 
    End If 
 
    'Application.ScreenUpdating = False 
     
       
    'Remove Autofilter 
    On Error Resume Next 
    rng_SourceRange.Worksheet.AutoFilterMode = False 
    On Error GoTo 0 
     
    'Recalculate before add 
    Call z_mRangeRecalculate(rng_SourceRange) 
     
    'Create New Table 
    Set tbl_NewTable = rng_SourceRange.Worksheet.ListObjects.Add(xlSrcRange, rng_SourceRange, , xlYes) 
     
    'Clear TableStyle 
    tbl_NewTable.TableStyle = "" 
     
    'Name Table if there is name Set 
    If str_TableName <> "" Then tbl_NewTable.Name = str_TableName 
       
    'Application.ScreenUpdating = True 
       
End Function 
 
Function PivotDataAppendToTable( _ 
                                ByVal pvt_Copy As PivotTable, _ 
                                ByVal tbl_PasteAppend As ListObject, _ 
                                Optional ByVal arr_HeaderMask As Variant _ 
                                ) 
 
    Call TableAppendData(tbl_PasteAppend, tbl_Copy, bool_Transpose, arr_HeaderMaskInput) 
 
End Function 
 
Public Function xLinkExcelObjects() 
 
    On Error GoTo err_AtkError 
  
'Parameters Cleanup 
    Call z_CleanRefErrNames 
    Call z_FillUnamedParameters 
 
    Dim CodePan As Object 
     
    Call zSelectCodePan(CodePan, "mod_AtkBindings") 
 
    Dim FindWhat As String 
    Dim SL As Long ' start line 
    Dim EL As Long ' end line 
 
    With CodePan 
        SL = 1 
        EL = .CountOfLines 
    End With 
         
'LOAD ALL VARIABLES IN WORKBOOK 
 
    Dim cc_RangeDeclarations As Collection 
    Dim cc_TableDeclarations As Collection 
    Dim cc_PivotDeclarations As Collection 
     
    Dim cc_RangeInitiation As Collection 
    Dim cc_TableInitiation As Collection 
    Dim cc_PivotInitiation As Collection 
 
    Set cc_RangeDeclarations = New Collection 
    Set cc_TableDeclarations = New Collection 
    Set cc_PivotDeclarations = New Collection 
     
    Set cc_RangeInitiation = New Collection 
    Set cc_TableInitiation = New Collection 
    Set cc_PivotInitiation = New Collection 
     
    Dim str_CodeLine As String 
    Dim name_Cursor As Name 
    Dim sht_Cursor As Worksheet 
    Dim tbl_Cursor As ListObject 
    Dim pvt_Cursor As PivotTable 
      
    For Each name_Cursor In ThisWorkbook.Names 
        If name_Cursor.Name Like "rng_*" Then 
             
            str_CodeLine = Replace("Public XNAME As Range", "XNAME", name_Cursor.Name) 
            Call cc_RangeDeclarations.Add(str_CodeLine, str_CodeLine) 
             
            str_CodeLine = Replace("Call shtATk.RangeCreateLink(XNAME, ""XNAME"")", "XNAME", name_Cursor.Name) 
 
            Call cc_RangeInitiation.Add(str_CodeLine, str_CodeLine) 
           
        End If 
    Next 
     
    For Each sht_Cursor In ThisWorkbook.Worksheets 
        For Each tbl_Cursor In sht_Cursor.ListObjects 
         
            If tbl_Cursor.Name Like "t_*" Then 
             
                str_CodeLine = Replace("Public XNAME As ListObject", "XNAME", tbl_Cursor.Name) 
                Call cc_TableDeclarations.Add(str_CodeLine, str_CodeLine) 
                 
                str_CodeLine = Replace("Call shtATk.TableCreateLink(XNAME, ""XNAME"")", "XNAME", tbl_Cursor.Name) 
                Call cc_TableInitiation.Add(str_CodeLine, str_CodeLine) 
             
            End If 
        Next 
         
        For Each pvt_Cursor In sht_Cursor.PivotTables 
             
            If pvt_Cursor.Name Like "pvt_*" Then 
             
                str_CodeLine = Replace("Public XNAME As PivotTable", "XNAME", pvt_Cursor.Name) 
                Call cc_PivotDeclarations.Add(str_CodeLine, str_CodeLine) 
                 
                str_CodeLine = Replace("Call shtATk.PivotCreateLink(XNAME, ""XNAME"")", "XNAME", pvt_Cursor.Name) 
                Call cc_PivotInitiation.Add(str_CodeLine, str_CodeLine) 
             
            End If 
        Next 
    Next 
 
'CHECK IF EXIST LINKER FUNCTION IF NOT CREATE IT 
    Dim lng_LinkerFunctionRow As Long 
         
    With CodePan 
         
        SL = 1 '.CountOfDeclarationLines ' 1 find first 
        Do Until .lines(SL, 1) Like "*Function *(*)" Or .lines(SL, 1) Like "*Sub *(*)" Or SL > .CountOfLines 
            Call z_mPP(SL) 
        Loop 
             
        If SL = 1 Then SL = 2 
             
        If Not z_mLinkExcelObjects_Codefind(CodePan, "Function LinkWorkbookTables()", SL) Then 
            Call .InsertLines(SL - 1, Join(Array("", "Function LinkWorkbookTables()", "", "End Function", ""), vbCr)) 
        End If 
  
    End With 
  
  
'DEFINE LINKER ROW 
    Call z_mLinkExcelObjects_Codefind(CodePan, "Function LinkWorkbookTables()", lng_LinkerFunctionRow) 
 
'CHECK AND REMOVE NOT WORKING LINKER LINES 
    Dim obj_Dummy As Object 
 
    With CodePan 
        SL = 1 'find first 
        Do Until .lines(SL, 1) Like "*End Function*" 
            If .lines(SL, 1) Like "*Call shtATk.RangeCreateLink*" Or .lines(SL, 1) Like "*Call shtATk.TableCreateLink*" Or .lines(SL, 1) Like "*Call shtATk.PivotCreateLink*" Then 
                Call CallByName(shtATk, Split(Split(.lines(SL, 1), ".")(1), "(")(0), VbMethod, obj_Dummy, CStr(Split(.lines(SL, 1), """")(1))) 
                  
                If obj_Dummy Is Nothing Then 
                    Dim str_ItemName As String 
             
                    str_ItemName = CStr(Split(.lines(SL, 1), """")(1)) 
                     
                    Debug.Print str_ItemName & " not found will be removed" 
                     
                'REMOVE INITIATION LINE 
                    Call .DeleteLines(SL) 
                    SL = SL - 1 'decrease counter to avoid skipping in loop 
                     
                'REMOVE DELCLARATION LINE IF EXIST 
                    SL = 1 
                    If z_mLinkExcelObjects_Codefind(CodePan, Replace("Public XNAME As", "XNAME", str_ItemName), SL) Then 
                        Call .DeleteLines(SL) 
                        SL = SL - 1 'decrease counter to avoid skipping in loop 
                    End If 
                    
                Else 
                    Set obj_Dummy = Nothing 
                End If 
            End If 
            Call z_mPP(SL) 
        Loop 
    End With 
 
 
    Dim str_CodePart As String 
    Dim lng_Row As Long 
         
'RANGES 
    For lng_Row = 1 To cc_RangeDeclarations.Count 
        If Not z_mLinkExcelObjects_Codefind(CodePan, cc_RangeDeclarations.Item(lng_Row)) Then 
            Debug.Print "Added Link for " & Split(cc_RangeDeclarations.Item(lng_Row), " ")(1) 
        End If 
     
        str_CodePart = str_CodePart & cc_RangeDeclarations.Item(lng_Row) & vbCr 
    Next 
     
    If cc_RangeDeclarations.Count > 0 Then str_CodePart = str_CodePart & vbCr 
             
'TABLES 
    For lng_Row = 1 To cc_TableDeclarations.Count 
        If Not z_mLinkExcelObjects_Codefind(CodePan, cc_TableDeclarations.Item(lng_Row)) Then 
            Debug.Print "Added Link for " & Split(cc_TableDeclarations.Item(lng_Row), " ")(1) 
        End If 
        str_CodePart = str_CodePart & cc_TableDeclarations.Item(lng_Row) & vbCr 
    Next 
     
    If cc_TableDeclarations.Count > 0 Then str_CodePart = str_CodePart & vbCr 
     
'PIVOTS 
    For lng_Row = 1 To cc_PivotDeclarations.Count 
        If Not z_mLinkExcelObjects_Codefind(CodePan, cc_PivotDeclarations.Item(lng_Row)) Then 
            Debug.Print "Added Link for " & Split(cc_PivotDeclarations.Item(lng_Row), " ")(1) 
        End If 
        str_CodePart = str_CodePart & cc_PivotDeclarations.Item(lng_Row) & vbCr 
    Next 
         
    If cc_PivotDeclarations.Count > 0 Then str_CodePart = str_CodePart & vbCr 
       
         
'FUNCTION 
         
    str_CodePart = str_CodePart & vbCr & "Function LinkWorkbookTables()" & vbCr & vbCr 
         
 
    For lng_Row = 1 To cc_RangeInitiation.Count 
        str_CodePart = str_CodePart & vbTab & cc_RangeInitiation.Item(lng_Row) & vbCr 
    Next 
     
    If cc_RangeInitiation.Count > 0 Then str_CodePart = str_CodePart & vbCr 
             
    For lng_Row = 1 To cc_TableInitiation.Count 
        str_CodePart = str_CodePart & vbTab & cc_TableInitiation.Item(lng_Row) & vbCr 
    Next 
     
    If cc_TableInitiation.Count > 0 Then str_CodePart = str_CodePart & vbCr 
     
    For lng_Row = 1 To cc_PivotInitiation.Count 
        str_CodePart = str_CodePart & vbTab & cc_PivotInitiation.Item(lng_Row) & vbCr 
    Next 
     
    If cc_PivotInitiation.Count > 0 Then str_CodePart = str_CodePart & vbCr 
     
    str_CodePart = str_CodePart & "End Function" & vbCr & vbCr 
' 
'CLEAR & PRINT NEW 
    Set CodePan = Nothing 
 
    Call zCodeRemove("mod_AtkBindings") 
    Call zCodeAppend(str_CodePart, "mod_ATkBindings") 
     
      
exit_ok: 
    xLinkExcelObjects = True 
Exit Function 
err_AtkError: 
  
    If frm_GlErr.z_Show(Err) Then 
        Stop: Resume 
    Else 
        Resume err_AtkError 
    End If 
     
End Function 
 
Private Function z_mLinkExcelObjects_Codefind( _ 
    ByRef obj_Module As Object, _ 
    ByVal str_FindCase As String, _ 
    Optional ByRef StartLine As Long = 1, _ 
    Optional ByRef EndLine As Long = -1, _ 
    Optional ByRef StartColumn As Long = 1, _ 
    Optional ByRef EndColumn As Long = 255, _ 
    Optional ByVal Wholeword As Boolean = False, _ 
    Optional ByVal MatchCase As Boolean = True, _ 
    Optional ByVal Patternsearch As Boolean = False) As Boolean 
 
    With obj_Module 
 
        If EndLine = -1 Then EndLine = .CountOfLines 
 
        z_mLinkExcelObjects_Codefind = .Find( _ 
                                        Target:=str_FindCase, _ 
                                        StartLine:=StartLine, _ 
                                        StartColumn:=StartColumn, _ 
                                        EndLine:=EndLine, _ 
                                        EndColumn:=EndColumn, _ 
                                        Wholeword:=Wholeword, _ 
                                        MatchCase:=MatchCase, _ 
                                        Patternsearch:=Patternsearch) 
    End With 
 
End Function 
 
 
Private Function z_mLinkExcelObjects_AddCode( _ 
    ByRef CodePan As Object, _ 
    ByRef cc_Declarations As Collection, _ 
    ByRef cc_Initiations As Collection, _ 
    ByRef lng_GlobalLastRowInput As Long, _ 
    ByRef lng_InitiateLastRowInput As Long) 
 
 
    Dim lng_GlobalLastRow As Long 
    Dim lng_InitiateLastRow As Long 
    Dim i_StatementCursor As Long 
    Dim SL As Long 
      
    'IDENTIFY EXISTING GLOBAL VARIABLES ROW 
    For i_StatementCursor = cc_Declarations.Count To 1 Step -1 
        If z_mLinkExcelObjects_Codefind(CodePan, cc_Declarations.Item(i_StatementCursor), SL) Then 
         
            Call cc_Declarations.Remove(i_StatementCursor) 
            If SL > lng_GlobalLastRow Then lng_GlobalLastRow = SL 
            SL = 1 
        End If 
    Next 
 
         
    'ADD NOT YET ADDED VARIABLES 
    If cc_Declarations.Count > 0 Then 
     
        If lng_GlobalLastRow = 0 Then 
            Call z_mPP(lng_GlobalLastRowInput) 'lng_GlobalLastRowInput + 1 
            Call z_mPP(lng_InitiateLastRowInput) 
             
            Call z_mLinkExcelObjects_WriteLine(CodePan, "", lng_GlobalLastRowInput) 
 
        Else 
            lng_GlobalLastRowInput = lng_GlobalLastRow 
        End If 
     
        For i_StatementCursor = cc_Declarations.Count To 1 Step -1 
            Call z_mPP(lng_GlobalLastRowInput) 'lng_GlobalLastRowInput + 1 
            Call z_mPP(lng_InitiateLastRowInput) 'lng_InitiateLastRow + 1 
             
            Call z_mLinkExcelObjects_WriteLine(CodePan, cc_Declarations.Item(i_StatementCursor), lng_GlobalLastRowInput) 
             
            Debug.Print "Added Link for " & Split(cc_Declarations.Item(i_StatementCursor), " ")(1) 
        Next 
     
    Else 
        lng_GlobalLastRowInput = lng_GlobalLastRow 
    End If 
 
           
    'FIND PROCEDURE OR ADD PROCEDURE IF NOT EXIST 
    For i_StatementCursor = cc_Initiations.Count To 1 Step -1 
        If z_mLinkExcelObjects_Codefind(CodePan, cc_Initiations.Item(i_StatementCursor), SL) Then 
         
            Call cc_Initiations.Remove(i_StatementCursor) 
            If SL > lng_InitiateLastRow Then lng_InitiateLastRow = SL 
            SL = lng_InitiateLastRowInput 
         
        End If 
    Next 
 
    'ADD NOT YET ADDED INITIATITIONS 
    If cc_Initiations.Count > 0 Then 
     
        If lng_InitiateLastRow = 0 Then 
            Call z_mPP(lng_InitiateLastRowInput) 
            Call z_mLinkExcelObjects_WriteLine(CodePan, "", lng_InitiateLastRowInput) 
        Else 
            lng_InitiateLastRowInput = lng_InitiateLastRow 
        End If 
     
        For i_StatementCursor = cc_Initiations.Count To 1 Step -1 
            Call z_mPP(lng_InitiateLastRowInput) 
            Call z_mLinkExcelObjects_WriteLine(CodePan, vbTab & cc_Initiations.Item(i_StatementCursor), lng_InitiateLastRowInput) 
        Next 
     
    Else 
        lng_InitiateLastRowInput = lng_InitiateLastRow 
    End If 
 
End Function 
 
Private Function z_mLinkExcelObjects_WriteLine(ByRef obj_Codepane As Object, ByVal str_Text As String, ByVal lng_Line As Long) 
    Call obj_Codepane.InsertLines(lng_Line, str_Text) 
     
    'Debug.Print lng_Line & vbTab & str_Text 
End Function 
 
Private Function z_mPP(ByRef lng_Counter As Long) 
    lng_Counter = lng_Counter + 1 
End Function 
 
Function xAddButton(ByVal str_ButtonCaption As String, Optional ByVal bool_ChewieButton = False) 
     
    Dim shape_Cursor As Shape 
 
    'chewie initial blocks 
    If bool_ChewieButton Then 
        If Not zModuleExists("clsSap") And Not zModuleExists("clsCollection") Then 
            MsgBox "Can't add chewie missing SAP class modules ... Exiting" 
            Exit Function 
        End If 
    End If 
 
    Dim str_ButtonName As String 
    Dim str_ActionName As String 
     
    str_ActionName = Replace(WorksheetFunction.Proper(Trim(str_ButtonCaption)), " ", "") 
    str_ButtonName = "btn_" & str_ActionName 
 
    If zProdedureExists(str_ButtonName, "mod_AtkMain") Then 
        MsgBox "There is already button with same name ... Exiting" 
        Exit Function 
    End If 
 
    Dim CodePan As Object 'VBIDE.CodeModule 
 
    Dim lng_ControlsTop As Long 
    Dim lng_Top As Long 
    Dim lng_Height As Long 
    Dim lng_Left As Long 
     
    lng_Top = ThisWorkbook.Sheets("Main").Range("B11").Top 
    lng_Height = ThisWorkbook.Sheets("Main").Range("B1:B2").Height 
    lng_Left = ThisWorkbook.Sheets("Main").Range("B1").Left 
 
    lng_ControlsTop = lng_Top 
 
    'find Lowest Button on main sheet 
    For Each shape_Cursor In sht_Main.Shapes 
     
        If shape_Cursor.Name Like "btn_*" Then 
     
            If shape_Cursor.OnAction <> "" Then 
             
                If shape_Cursor.Top + shape_Cursor.Height > lng_ControlsTop Then lng_ControlsTop = shape_Cursor.Top + shape_Cursor.Height 
                 
            End If 
         
        End If 
     
    Next 
     
    'check size add one more button size 
    If lng_ControlsTop = lng_Top + lng_Height Then lng_ControlsTop = lng_ControlsTop + lng_Height 
 
    'Clean up 
    On Error Resume Next 
        ThisWorkbook.Sheets("Main").Shapes("btn_ButtonTemplate").Delete 
    On Error GoTo 0 
 
    'DUPLICATE BUTTON 
    ThisWorkbook.Sheets("ATk").Shapes("btn_ButtonTemplate").Copy 
    ThisWorkbook.Sheets("Main").Paste 
     
     
    Dim new_Button As Shape 
    Set new_Button = ThisWorkbook.Sheets("Main").Shapes("btn_ButtonTemplate") 
     
    new_Button.Top = lng_ControlsTop 
    new_Button.Left = lng_Left 
    new_Button.Name = str_ButtonName 
    new_Button.TextFrame2.TextRange.Characters.Text = str_ButtonCaption 
    new_Button.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1 
    new_Button.OnAction = str_ButtonName 
     
    Dim str_NewProcedure As String 
     
    Dim str_Path As String 
    Dim strcode As String 
     
    If bool_ChewieButton Then 
          
        'Pick VBS Script 
        Select Case MsgBox("Select any VBS file ?", vbYesNoCancel, "Import SAPGUI Check") 
             
            Case vbYes 
                str_Path = Application.GetOpenFilename 
         
                Do While str_Path = "" 
                    str_Path = Application.GetOpenFilename 
                Loop 
                 
                strcode = z_ChewCode(str_Path) 
                 
            Case vbNo 
             
                strcode = z_ChewCode("skip") 
                 
            Case vbCancel 
                Exit Function 
                 
        End Select 
         
        'Get Code Module 
        Call shtATk.zCodeAppend(Replace(strcode, "XSUBROUTINENAME", str_ActionName), "mod_Chewie") 
       
 
 
        str_NewProcedure = Join(Array( _ 
            "Public Sub " & str_ButtonName & "()", _ 
            "", _ 
            vbTab & "With shtATk", _ 
            vbTab & vbTab & ".MacroStart", _ 
             "", _ 
            vbTab & vbTab & "Call mod_Chewie." & str_ActionName, _ 
            vbTab & vbTab & "' your script goes here", _ 
             "", _ 
            vbTab & vbTab & ".MacroFinish", _ 
            vbTab & "End With", _ 
            "", _ 
            "End Sub"), vbCr) 
    Else 
     
        str_NewProcedure = Join(Array( _ 
            "Public Sub " & str_ButtonName & "()", _ 
            "", _ 
            vbTab & "With shtATk", _ 
            vbTab & vbTab & ".MacroStart", _ 
             "", _ 
            vbTab & vbTab & "' your script goes here", _ 
             "", _ 
            vbTab & vbTab & ".MacroFinish", _ 
            vbTab & "End With", _ 
            "", _ 
            "End Sub"), vbCr) 
    End If 
         
    Call zCodeAppend(str_NewProcedure, "mod_AtkMain") 
 
End Function 
 
Private Function z_ChewCode(ByVal str_Path) As String 
     
    'moved to private function to prevent compile error if clsSap class module is missing 
     
    Dim clsSap As Object 
    Set clsSap = New clsSap 
 
    z_ChewCode = clsSap.SapChewVBS(str_Path) 
      
    Set clsSap = Nothing 
     
End Function 
 
Function zSelectCodePan(ByRef CodePan As Variant, ByVal str_CodeModuleName As String, Optional ByVal wbk_Target As Workbook = Nothing) As Boolean 
 
    Dim vbProj As Object 
    Dim VBComp As Object 
 
 
    If wbk_Target Is Nothing Then Set wbk_Target = ThisWorkbook 
 
    Set vbProj = wbk_Target.VBProject 
     
    On Error Resume Next 
     
    Set VBComp = vbProj.VBComponents(str_CodeModuleName) 
     
    On Error GoTo 0 
     
    If VBComp Is Nothing Then 
     
        Set VBComp = vbProj.VBComponents.Add(1) '1 = vbext_ct_StdModule 
        VBComp.Name = str_CodeModuleName 
     
    End If 
     
    Set CodePan = VBComp.CodeModule 
     
End Function 
 
Function zProdedureExists(ByVal str_ProcName As String, ByVal str_CodeModuleName As String, Optional ByVal wbk_Target As Workbook = Nothing) As Boolean 
 
    Dim VBComp As Object 
 
 
    If wbk_Target Is Nothing Then Set wbk_Target = ThisWorkbook 
 
    If Not zModuleExists(str_CodeModuleName, wbk_Target) Then 
        MsgBox "Module called " & str_CodeModuleName & " not found ... Exiting" 
        Exit Function 
    End If 
     
     'ENTER CODE 
    Call zSelectCodePan(VBComp, "mod_AtkMain", wbk_Target) 
 
    Dim StartLine As Long 
    Dim NumLines As Long 
    Dim ProcName As String 
 
    On Error GoTo Err 
    With VBComp 
        StartLine = .ProcStartLine(str_ProcName, 0) 
    End With 
 
Err: 
    If Err.Number = 0 Then 
        zProdedureExists = True 
    Else 
        zProdedureExists = False 
    End If 
     
    Err.Clear 
 
End Function 
 
Function xAddPicker(ByVal str_PickerName As String, Optional lng_PickerType As e_PickerType = e_FilePicker) 
 
On Error GoTo Err 
 
    Dim lng_Top As Long 
    Dim lng_Width As Long 
    Dim lng_Height As Long 
    Dim lng_Left As Long 
     
    Dim ole_Cursor As OLEObject 
     
    'sht main teplate button 
    lng_Width = ThisWorkbook.Sheets("ATk").Shapes("btn_PickerTemplate").Width 
    lng_Height = ThisWorkbook.Sheets("ATk").Shapes("btn_PickerTemplate").Height 
     
    Dim lng_ControlsTop As Long 
     
    Dim str_PickerCodeName As String 
    Dim str_ButtonName As String 
     
    str_PickerCodeName = Replace(WorksheetFunction.Proper(Trim(str_PickerName)), " ", "") 
    str_ButtonName = "btn_" & str_PickerCodeName 
     
    'Existing Procedure Check 
    If zProdedureExists(str_ButtonName, "mod_AtkMain") Then 
        MsgBox "There is already button with same name ... Exiting" 
        Exit Function 
    End If 
     
    'Old Range I11:I30 
    Dim rng_Link As Range 
     
    With ThisWorkbook.Sheets("Main").ListObjects("t_Parameters") 
        Set rng_Link = .HeaderRowRange.Offset(.Range.Rows.Count, 1).Resize(, 1) 
    End With 
     
    rng_Link.Offset(, 1).Name = "rng_" & str_PickerCodeName 
    rng_Link.Value = str_PickerName 
     
    'Clean up Existing Dummy 
    On Error Resume Next 
        ThisWorkbook.Sheets("Main").Shapes("btn_PickerTemplate").Delete 
    On Error GoTo Err 
 
    'DUPLICATE BUTTON 
    ThisWorkbook.Sheets("ATk").Shapes("btn_PickerTemplate").Copy 
    ThisWorkbook.Sheets("Main").Paste 
     
    Dim new_Button As Shape 
    Set new_Button = ThisWorkbook.Sheets("Main").Shapes("btn_PickerTemplate") 
 
    'size, location of button 
    new_Button.Width = rng_Link.Offset(, -1).Width 
    new_Button.Height = rng_Link.Offset(, -1).Height 
    new_Button.Top = rng_Link.Top 
    new_Button.Left = rng_Link.Offset(, -1).Left 
    
     
    new_Button.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1 
    new_Button.Name = str_ButtonName 
    new_Button.OnAction = str_ButtonName 
 
     'ENTER CODE 
    'Call zSelectCodePan(CodePan, "mod_AtkMain") 
 
    Dim str_ActionName As String 
 
    Select Case lng_PickerType 
        Case e_FilePicker: str_ActionName = "PickerFile" 
        Case e_FolderPicker: str_ActionName = "PickerFolder" 
        Case e_OutlookPicker: str_ActionName = "PickerOutlook" 
    End Select 
 
    Dim str_NewProcedure As String 
     
    str_NewProcedure = Join(Array( _ 
        "Public Sub " & str_ButtonName & "()", _ 
        "", _ 
        vbTab & "With shtATk", _ 
        vbTab & vbTab & ".MacroStart", _ 
         "", _ 
        vbTab & vbTab & "Call ." & str_ActionName & "(rng_" & str_PickerCodeName & ", ""Please select " & str_PickerName & """)", _ 
         "", _ 
        vbTab & vbTab & ".MacroFinish", _ 
        vbTab & "End With", _ 
        "", _ 
        "End Sub"), vbCr) 
 
    Call zCodeAppend(str_NewProcedure, "mod_AtkMain") 
 
    DoEvents 
    Call xLinkExcelObjects 
 
Exit Function 
Err: 
 
Err.Raise Err.Number 
Resume 
 
End Function 
 
 Function zRemoveProcedure(ByVal str_ProcName As String, ByVal str_ModuleName As String, Optional ByRef wbk_Target As Workbook = Nothing) As Boolean 
 
    If wbk_Target Is Nothing Then Set wbk_Target = ThisWorkbook 
     
    Dim CodePan As Object 
     
    'ENTER CODE 
    Call zSelectCodePan(CodePan, "mod_AtkMain", wbk_Target) 
 
    Dim StartLine As Long 
    Dim NumLines As Long 
    Dim ProcName As String 
 
    With CodePan 
        StartLine = .ProcStartLine(str_ProcName, vbext_pk_Proc) 
        NumLines = .ProcCountLines(str_ProcName, vbext_pk_Proc) 
        .DeleteLines StartLine:=StartLine, Count:=NumLines 
    End With 
 
End Function 
 
Function zModuleExists( _ 
    ByVal str_ModuleName As String, _ 
    Optional ByRef wbk_Target As Workbook = Nothing _ 
        ) As Boolean 
 
    If wbk_Target Is Nothing Then Set wbk_Target = ThisWorkbook 
 
    Dim vbProj As Object 
    Dim VBComp As Object 
    
    Set vbProj = wbk_Target.VBProject 
     
    On Error Resume Next 
     
    Set VBComp = vbProj.VBComponents(str_ModuleName) 
     
    On Error GoTo 0 
     
    zModuleExists = Not VBComp Is Nothing 
 
End Function 
 
Function zCodeGet( _ 
    ByRef arr_Code As Variant, _ 
    ByVal str_ModuleName As String, _ 
    Optional ByRef wbk_Target As Workbook = Nothing _ 
        ) As Boolean 
     
    Dim CodePan As Object 
 
     'ENTER CODE 
    Call zSelectCodePan(CodePan, str_ModuleName, wbk_Target) 
 
    'Dim arr_Code() 
 
    With CodePan 
        arr_Code = Split(.lines(1, .CountOfLines), vbCr) 
    End With 
 
    zCodeGet = True 
     
End Function 
 
Function zCodeAppend( _ 
    ByVal str_Code As String, _ 
    ByVal str_ModuleName As String, _ 
    Optional ByRef wbk_Target As Workbook = Nothing _ 
        ) As Boolean 
     
    Dim CodePan As Object 
     
    'ENTER CODE 
    Call zSelectCodePan(CodePan, str_ModuleName, wbk_Target) 
 
    With CodePan 
        Call .InsertLines(IIf(.CountOfLines = 0, 1, .CountOfLines + 1), str_Code) 
    End With 
 
 
End Function 
 
Function xAddParameter( _ 
    ByVal str_ParameterName As String _ 
        ) As Boolean 
 
On Error GoTo Err 
 
    DoEvents 
    Call xLinkExcelObjects 
 
    Dim ole_Cursor As OLEObject 
     
    'sht main teplate butto 
    Dim lng_ControlsTop As Long 
     
    'find Lowest Button on main sheet 
    Dim rng_Link As Range 
     
    With ThisWorkbook.Sheets("Main").ListObjects("t_Parameters") 
        Set rng_Link = .HeaderRowRange.Offset(.Range.Rows.Count, 1).Resize(, 1) 
    End With 
     
    'add named range 
    rng_Link.Offset(, -1).Name = "rng_" & Replace(WorksheetFunction.Proper(Trim(str_ParameterName)), " ", "") 
    rng_Link.Value = str_ParameterName 
 
    DoEvents 
    Call xLinkExcelObjects 
 
Exit Function 
Err: 
 
Err.Raise Err.Number 
Resume 
 
End Function 
 
Function FileExist( _ 
    ByVal var_FilePath As Variant _ 
        ) As Boolean 
     
    Dim str_Path As String 
 
    str_Path = CStr(z_mCovertToSimpleArray(var_FilePath)(0)) 
 
    Call z_mFsoOpen 
     
    FileExist = fso_Object.FileExists(str_Path) 
 
End Function 
  
Function FileCopy( _ 
    ByVal var_FilePath As Variant, _ 
    Optional ByVal var_FileName As Variant, _ 
    Optional ByVal var_SourceFile As Variant, _ 
    Optional ByVal lng_Overwrite As e_Action = e_Prompt _ 
        ) As Boolean 
      
    Dim str_DestFilePath As String 
    Dim str_DestFileName As String 
    Dim str_SourceFullPath As String 
    Dim str_SourceFileName As String 
     
    'Source File 
    If IsMissing(var_SourceFile) Then 
        str_SourceFullPath = ThisWorkbook.FullName 
    Else 
        str_SourceFullPath = CStr(z_mCovertToSimpleArray(var_SourceFile)(0)) 
    End If 
     
    'TBD Check if Source File Exists 
    str_SourceFileName = Mid(str_SourceFullPath, InStrRev(str_SourceFullPath, "\") + 1) 
     
    'Destination File 
    str_DestFilePath = CStr(z_mCovertToSimpleArray(var_FilePath)(0)) 
  
    'open file 
    Call z_mFsoOpen 
         
    'rename file if needed (not missing) 
    If Not IsMissing(var_FileName) Then 
    str_DestFileName = CStr(z_mCovertToSimpleArray(var_FileName)(0)) 
    End If 
         
    'Copy File 
    With fso_Object 
     
        Dim bool_OverWrite As Boolean 
     
        If .FileExists(str_DestFilePath & "\" & str_DestFileName) Then 
         
            If lng_Overwrite = e_Prompt Then 
                bool_OverWrite = MsgBox("File Exists overwrite ?", vbQuestion Or vbYesNo) = vbYes 
            Else 
                bool_OverWrite = lng_Overwrite 
            End If 
         
        Else 
            bool_OverWrite = True 
        End If 
         
        Call .copyfile(str_SourceFullPath, str_DestFilePath & "\", bool_OverWrite) 
         
    End With 
 
    'Rename 
    If Not IsMissing(var_FileName) Then 
         
        If bool_OverWrite Then 
             
            On Error Resume Next 
            Kill str_DestFilePath & "\" & str_DestFileName 
            On Error GoTo 0 
             
        End If 
         
        Name str_DestFilePath & "\" & str_SourceFileName As str_DestFilePath & "\" & str_DestFileName 
     
    End If 
 
End Function 
   
Function z_mFsoOpen( _ 
        ) As Boolean 
 
    If fso_Object Is Nothing Then 
       Set fso_Object = CreateObject("Scripting.FileSystemObject") 
    End If 
 
End Function 
  
  
'Function TableRowAddByMap( _ 
'    ByRef tbl_SourceTable As ListObject, _ 
'    ByRef tbl_MapTable As ListObject, _ 
'    Optional ByRef wbk_Source As Workbook) As Boolean 
' 
'    Dim arr_HeaderColumns 
'    Dim arr_Map() 
' 
'    Dim wbk_Temp As Workbook 
' 
'    If wbk_Source Is Nothing Then 
'        Set wbk_Temp = ThisWorkbook 
'    Else 
'        Set wbk_Temp = wbk_Source 
'    End If 
' 
'    arr_HeaderColumns = WorksheetFunction.Index(tbl_SourceTable.HeaderRowRange, 1, 0) 
' 
'    Dim lng_Cursor As Long 
' 
'    arr_Map = tbl_MapTable.Range 
' 
'    Dim lng_TableNewLine As Long 
' 
'    Dim lng_SheetIndex As Long 
'    Dim lng_RangeRef As Long 
'    Dim lng_TargetColumn As Long 
' 
'    Dim lng_TableFirstColumn As Long 
' 
'    lng_TableFirstColumn = tbl_MapTable.HeaderRowRange.Resize(1, 1).Column - 1 
' 
'    With tbl_MapTable.HeaderRowRange 
' 
'        lng_SheetIndex = .Find("Sheet").Column - lng_TableFirstColumn 
'        lng_RangeRef = .Find("Range").Column - lng_TableFirstColumn 
'        lng_TargetColumn = .Find("ColumnName").Column - lng_TableFirstColumn 
' 
'    End With 
' 
' 
'    Dim lng_OffsetRow As Long 
' 
' 
'    lng_OffsetRow = tbl_SourceTable.Range.Rows.Count 
' 
'    If lng_OffsetRow = 2 Then 
' 
'        Dim rng_FirstRow As Range 
'        On Error Resume Next 
'        Set rng_FirstRow = tbl_SourceTable.HeaderRowRange.Offset(1).SpecialCells(xlCellTypeConstants) 
'        On Error GoTo 0 
' 
'        If rng_FirstRow Is Nothing Then 
'            lng_OffsetRow = 1 
'        End If 
' 
'    End If 
' 
' 
'    With tbl_SourceTable.HeaderRowRange.Offset(lng_OffsetRow) 
' 
' 
'        For lng_Cursor = LBound(arr_Map) + 1 To UBound(arr_Map) 
' 
'            On Error Resume Next 
'             .Cells(1, Application.Match(arr_Map(lng_Cursor, lng_TargetColumn), arr_HeaderColumns, 0)) = wbk_Temp.Sheets(arr_Map(lng_Cursor, lng_SheetIndex)).Range(arr_Map(lng_Cursor, lng_RangeRef)).Value 
' 
'                Debug.Print Err.Number 
'            On Error GoTo 0 
' 
'        Next 
' 
'    .Calculate 
' 
'    End With 
' 
' 
'End Function 
  
Function TableRecordByMap( _ 
    ByRef var_SourceInput As Variant, _ 
    ByRef tbl_MapTable As ListObject, _ 
    Optional ByRef wbk_Source As Variant, _ 
    Optional ByRef lng_Direction As e_DirectionCopy = e_InsideTable _ 
        ) As Boolean 
     
    Dim arr_HeaderColumns 
    Dim arr_Map() 
     
    'Define workbook 
    Dim wbk_Temp As Workbook 
     
    If IsMissing(wbk_Source) Then 
        Set wbk_Temp = ThisWorkbook 
    Else 
        Set wbk_Temp = wbk_Source 
    End If 
      
    'Define Table 
    Dim tbl_SourceTable As ListObject 
    Dim bool_UpdateData As Boolean 
      
    If TypeName(var_SourceInput) = "Range" Then 
        Set tbl_SourceTable = var_SourceInput.ListObject 
        bool_UpdateData = True 
    End If 
     
    If TypeName(var_SourceInput) = "ListObject" Then 
        Set tbl_SourceTable = var_SourceInput 
        bool_UpdateData = False 
    End If 
     
    arr_HeaderColumns = WorksheetFunction.Index(tbl_SourceTable.HeaderRowRange, 1, 0) 
     
    Dim lng_Cursor As Long 
     
    arr_Map = tbl_MapTable.Range 
     
    Dim lng_TableNewLine As Long 
     
    Dim lng_SheetIndex As Long 
    Dim lng_RangeRef As Long 
    Dim lng_TargetColumn As Long 
     
    Dim lng_TableFirstColumn As Long 
     
    lng_TableFirstColumn = tbl_MapTable.HeaderRowRange.Resize(1, 1).Column - 1 
     
    With tbl_MapTable.HeaderRowRange 
         
        lng_SheetIndex = .Find("Sheet").Column - lng_TableFirstColumn 
        lng_RangeRef = .Find("Range").Column - lng_TableFirstColumn 
        lng_TargetColumn = .Find("ColumnName").Column - lng_TableFirstColumn 
          
    End With 
     
     
    Dim lng_OffsetRow As Long 
      
      
    If bool_UpdateData Then 
     
        lng_OffsetRow = var_SourceInput.Row - tbl_SourceTable.HeaderRowRange.Row 
     
    Else 
     
    lng_OffsetRow = tbl_SourceTable.Range.Rows.Count 
     
    If lng_OffsetRow = 2 Then 
         
        Dim rng_FirstRow As Range 
        On Error Resume Next 
        Set rng_FirstRow = tbl_SourceTable.HeaderRowRange.Offset(1).SpecialCells(xlCellTypeConstants) 
        On Error GoTo 0 
         
        If rng_FirstRow Is Nothing Then 
            lng_OffsetRow = 1 
            End If 
          
        End If 
     
    End If 
     
    'recalculate table 
    tbl_MapTable.Range.Calculate 
     
    'Data Execution 
    With tbl_SourceTable.HeaderRowRange.Offset(lng_OffsetRow) 
 
 
        For lng_Cursor = LBound(arr_Map) + 1 To UBound(arr_Map) 
             
            On Error Resume Next 
             
            If lng_Direction = e_InsideTable Then 
             .Cells(1, Application.Match(arr_Map(lng_Cursor, lng_TargetColumn), arr_HeaderColumns, 0)) = wbk_Temp.Sheets(arr_Map(lng_Cursor, lng_SheetIndex)).Range(arr_Map(lng_Cursor, lng_RangeRef)).Value 
            Else 
                wbk_Temp.Sheets(arr_Map(lng_Cursor, lng_SheetIndex)).Range(arr_Map(lng_Cursor, lng_RangeRef)).Value = .Cells(1, Application.Match(arr_Map(lng_Cursor, lng_TargetColumn), arr_HeaderColumns, 0)) 
            End If 
 
            On Error GoTo 0 
         
        Next 
         
    .Calculate 
     
    End With 
 
 
End Function 
 
Public Sub zATK_Issue() 
 
    MacroStart 
         
    Dim obj_Email As Object 
     
    Call EmailCreateNew(obj_Email, Array( _ 
                        "<a href=""" & ThisWorkbook.Path & """>BAU</a>", _ 
                        "<a href=""" & rng_ManualLink.Text & """>MAN</a>"), _ 
                        ThisWorkbook.Name, _ 
                        Me.Range(adr_IssueEmail).Value2) 
     
     
    Call EmailAttachWorkbook(obj_Email, ThisWorkbook) 
     
    MacroFinish 
 
End Sub 
 
 
Public Sub zATK_PickManual() 
 
    MacroStart 
         
    Call shtATk.FilePicker(rng_ManualLink, "Please select Manual File") 
         
    MacroFinish 
     
End Sub 
 
 
Sub zCodeRemove(ByVal str_ModuleName As String, Optional ByRef wbk_Target As Workbook = Nothing) 
   
    Dim codemod As Object 
   
    Call shtATk.zSelectCodePan(codemod, str_ModuleName, wbk_Target) 
     
    With codemod 
        Call .DeleteLines(1, .CountOfLines) 
    End With 
 
End Sub 
 
Function zLocalGitCodeExport(ByVal str_ModuleName, Optional ByRef wbk_Data As Workbook) As Boolean 
     
    Dim arr_Code 
     
    Call zCodeGet(arr_Code, str_ModuleName, wbk_Data) 
 
    Dim file_Output As Object 
 
    With CreateObject("Scripting.FileSystemObject") 
     
        'write file 
         
        On Error Resume Next 
        MkDir shtATk.Range(adr_LocalGitPath) & "\Codes\" 
        On Error GoTo 0 
         
        With .CreateTextFile(shtATk.Range(adr_LocalGitPath) & "\Codes\" & str_ModuleName & ".vb", True) 
       
            Call .WriteLine(Join(arr_Code)) 
            Call .Close 
 
        End With 
 
    End With 
 
    zLocalGitCodeExport = True 
 
End Function 
 
Function zLocalGitCodeImport(ByVal str_ModuleName, Optional ByRef wbk_Data As Workbook) As Boolean 
 
    On Error GoTo Err: 
  
    zLocalGitCodeImport = False 
 
    'check if not importing ATK 
     
    If str_ModuleName = "shtATk" And wbk_Data Is Nothing Then 
        MsgBox "shtATk connot be imported via this function." & vbCrLf & "Please use full update via ""zCodeLocalGitUpdate"" ", vbCritical, "Import Error" 
        Exit Function 
    End If 
  
    Dim str_Code As String 
 
    With CreateObject("Scripting.FileSystemObject") 
     
        If .FileExists(shtATk.Range(adr_LocalGitPath) & "\Codes\" & str_ModuleName & ".vb") = False Then 
            Call MsgBox("File with module not found exitting ...") 
            Exit Function 
        End If 
     
        With .OpenTextFile(shtATk.Range(adr_LocalGitPath) & "\Codes\" & str_ModuleName & ".vb", 1) 
 
            str_Code = .ReadAll 
            .Close 
 
        End With 
 
    End With 
     
    Call zCodeRemove(str_ModuleName, wbk_Data) 
    Call zCodeAppend(str_Code, str_ModuleName, wbk_Data) 
     
    zLocalGitCodeImport = True 
 
Err: 
 
    If Err.Number <> 0 Then 
     
        Debug.Print str_ModuleName & "_" & Err.Description 
        Resume Err 
     
    End If 
  
End Function 
 
Function zLocalGitCodeUpdate() As Boolean 
         
        Dim str_OriginalName As String 
        str_OriginalName = ThisWorkbook.Name 
     
        'save original for update 
        Application.DisplayAlerts = False 
        Call ThisWorkbook.Save 
         
        'save backup 
        Call ThisWorkbook.SaveAs(ThisWorkbook.Path & "\" & Replace(str_OriginalName, ".xls", "_bckp" & Format(Now, "ddmmyy-hhnnss") & ".xls")) 
        Application.DisplayAlerts = True 
         
        'Open New Version 
        Dim wbk_NewUpdated As Workbook 
        Set wbk_NewUpdated = Workbooks.Open(ThisWorkbook.Path & "\" & str_OriginalName) 
         
        'Update Codes 
        Call zLocalGitCodeImport("shtATk", wbk_NewUpdated) 
        Call zLocalGitCodeImport("clsSap", wbk_NewUpdated) 
        Call zLocalGitCodeImport("frm_GlErr", wbk_NewUpdated) 
         
        'Close backup and activate updated 
        wbk_NewUpdated.Activate 
        Application.DisplayAlerts = True 
  
        Call ThisWorkbook.Close(False) 
 
End Function 
 
 

Function ConnectionOpen(ByRef db As Object, ByVal dbPath As Variant, Optional ByVal bool_CopyNetworkDb As Boolean = True) As Boolean 
 
    Dim connDb As Object 
    Set connDb = CreateObject("ADODB.Connection") 
     
'    Dim connDb As Object 'ADODB.Connection 
'    Set connDb = New ADODB.Connection ' CreateObject("ADODB.Connection") 
     
    Dim strConnection As String 
     
    Dim str_dbPath As String 
     
    str_dbPath = z_mCovertToSimpleArray(dbPath)(0) 
     
    If bool_CopyNetworkDb Then 
        Call Me.FileCopy(Environ("Temp"), , str_dbPath, e_Yes) 
     
        str_dbPath = Environ("Temp") & "\" & Mid(str_dbPath, InStrRev(str_dbPath, "\") + 1) 
    End If 
 
    'connDb.Provider = "" 
    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" _ 
         & "Data Source = " & str_dbPath 
 
    Call connDb.Open(strConnection) 
     
    Set db = connDb 
 
End Function 
  
  
Function ConnectionClose(ByRef db As Object) As Boolean 
 
    db.Close 
    Set db = Nothing 
 
End Function 
  
  
Function QueryInit(ByRef rs As Object, ByVal rs_Name As Variant, ByRef Connection As Object) 
     
    Dim cmd As Object 'ADODB.Command 
    Dim str_RsCommand As String 
     
    Const adOpenStatic As Long = 3 
    Const adLockReadOnly As Long = 1 
 
    'Change Name 
    str_RsCommand = z_mCovertToSimpleArray(rs_Name)(0) 
     
    Set cmd = CreateObject("ADODB.Command") 
    'Set cmd = New ADODB.Command 
     
    cmd.ActiveConnection = Connection 
    cmd.CommandText = rs_Name 
    cmd.CommandType = 4 'adCmdStoredProc 
     
    'cmd.NamedParameters = True 
    'cmd.Parameters.Refresh 
 
    Set rs = cmd 
 
    Set cmd = Nothing 
 
 
End Function 
 
Function QuerySetParameter(ByRef rs As Object, ByVal par_Name As Variant, ByVal par_Value As Variant, ByVal par_ValueType As e_QueryParameter) 
    Dim obj_Par As Object 'ADODB.parameter 
 
    Set obj_Par = rs.CreateParameter(par_Name, par_ValueType, 1) 'adParamInput 
 
    obj_Par.Value = par_Value 
     
    Call rs.Parameters.Append(obj_Par) 
 
    Set obj_Par = Nothing 
 
End Function 
 
Function QueryExecute(ByRef Command As Object) 
 
    'Dim rs_temp As ADODB.Recordset 
    'Set rs_temp = New ADODB.Recordset 
 
    Dim rs_temp As Object 
    Set rs_temp = CreateObject("ADODB.RecordSet") 
 
    Call rs_temp.Open(Command, , 3, 1) 'adOpenStatic,adLockReadOnly 
 
    If Not rs_temp.BOF And Not rs_temp.EOF Then 
         
        rs_temp.MoveLast 
        rs_temp.MoveFirst 
  
    End If 
  
    Set Command = Nothing 
    Set Command = rs_temp 
 
    Set rs_temp = Nothing 
 
End Function 

