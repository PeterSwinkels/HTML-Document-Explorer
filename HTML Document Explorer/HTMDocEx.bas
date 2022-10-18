Attribute VB_Name = "HTMLDocumentExplorerModule"
'This procedure contains this program's core procedures.
Option Explicit

'The Microsoft API constants, functions, and structures used by this program:
Private Type REFIID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_PROC_NOT_FOUND As Long = 127
Private Const ERROR_SUCCESS As Long = 0
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const SMTO_ABORTIFHUNG As Long = &H2&
Private Const S_OK As Long = &H0&

Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function EnumChildWindows Lib "User32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumWindows Lib "User32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetProcessImageFileNameW Lib "Psapi.dll" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ObjectFromLresult Lib "Oleacc.dll" (ByVal LResult As Long, riid As REFIID, ByVal wParam As Long, ppvObject As Any) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function RegisterWindowMessageA Lib "User32.dll" (ByVal lpString As String) As Long
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Function SendMessageTimeoutA Lib "User32.dll" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

'The constants and structures used by this program:

'This structure defines the information for any active HTML documents found.
Public Type HTMLDocumentStr
   DocumentO As HTMLDocument   'Defines the reference to a HTML document.
   ErrorMessage As String      'Defines a message if an error has occurred.
   WindowH As Long             'Defines the handle of the window containing the document.
End Type

Private Const MAX_PATH As Long = 260       'Defines the maximum number of characters allowed for a file path.
Private Const MAX_STRING As Long = 65535   'Defines the maximum number of characters used for a string buffer.
Private Const NO_HANDLE As Long = 0        'Defines "no handle".
Private Const NO_MESSAGE As Long = 0       'Defines "no window message."
Private Const NONE As Long = -1            'Defines "no HTML document."

'This procedure checks the specified window for a HTML document and returns it if found.
Private Function CheckForDocument(WindowH As Long, WMHTMLGetObjectMessage As Long) As HTMLDocument
On Error GoTo ErrorTrap
Dim DocumentO As HTMLDocument
Dim DocumentREFIID As REFIID
Dim LResult As Long
   
   With DocumentREFIID
      .Data1 = &H626FC520
      .Data2 = &HA41E
      .Data3 = &H11CF
      .Data4(0) = &HA7
      .Data4(1) = &H31
      .Data4(2) = &H0
      .Data4(3) = &HA0
      .Data4(4) = &HC9
      .Data4(5) = &H8
      .Data4(6) = &H26
      .Data4(7) = &H37
   End With
   
   CheckForError SendMessageTimeoutA(WindowH, WMHTMLGetObjectMessage, CLng(0), CLng(0), SMTO_ABORTIFHUNG, CLng(1000), LResult), ERROR_ACCESS_DENIED
   Set DocumentO = Nothing
   If Not LResult = 0 Then CheckForError ObjectFromLresult(LResult, DocumentREFIID, CLng(0), DocumentO), ERROR_PROC_NOT_FOUND
   
EndProcedure:
   Set CheckForDocument = DocumentO
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Function

'This procedure checks any API errors that have occurred.
Private Function CheckForError(ReturnValue As Long, Optional Ignored As Long = ERROR_SUCCESS)
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   On Error GoTo ErrorTrap
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
    
      Message = "API error code: " & CStr(ErrorCode) & vbCr
      Message = Message & Description
      Message = Message & "Return value: " & CStr(ReturnValue) & vbCr
      MsgBox Message, vbExclamation
   End If
   
EndProcedure:
   CheckForError = ReturnValue
   Exit Function

ErrorTrap:
   HandleError
   Resume EndProcedure
End Function

'This procedure checks for frames in the specified HTML document.
Private Sub CheckForFrames(DocumentO As HTMLDocumentStr)
On Error GoTo ErrorTrap
Dim ErrorMessage As String
Dim Frame As HTMLDocument
Dim FrameIndex() As Long
Dim NextFrame As HTMLDocument
Dim Parents() As HTMLDocument
Dim Level As Long

   Level = 0
   ReDim FrameIndex(0 To Level) As Long
   ReDim Parents(0 To Level) As HTMLDocument
   Set Frame = DocumentO.DocumentO
   Do Until (Level = 0) And (FrameIndex(Level) >= Frame.frames.Length)
      Do While FrameIndex(Level) < Frame.frames.Length
         Set NextFrame = GetFrame(Frame, FrameIndex(Level), ErrorMessage)
         If NextFrame Is Nothing Then Exit Do
         Level = Level + 1
         ReDim Preserve FrameIndex(0 To Level) As Long
         ReDim Preserve Parents(0 To Level) As HTMLDocument
         Set Parents(Level) = Frame
         Set Frame = NextFrame
      Loop
      
      If NextFrame Is Nothing Then
         DocumentList AddWindowH:=DocumentO.WindowH, AddDocument:=Nothing, Refresh:=False, NewErrorMessage:=ErrorMessage
      Else
         DocumentList AddWindowH:=DocumentO.WindowH, AddDocument:=Frame
         Set Frame = Parents(Level)
         Level = Level - 1
         ReDim Preserve FrameIndex(0 To Level) As Long
         ReDim Preserve Parents(0 To Level) As HTMLDocument
      End If
      
      If FrameIndex(Level) < Frame.frames.Length Then FrameIndex(Level) = FrameIndex(Level) + 1
   Loop
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure maintains a list of all active HTML documents found.
Public Function DocumentList(Optional AddWindowH As Long = 0, Optional AddDocument As HTMLDocument = Nothing, Optional Index As Long = NONE, Optional Refresh As Boolean = False, Optional NewErrorMessage As String = Empty) As HTMLDocumentStr
On Error GoTo ErrorTrap
Dim DocumentO As HTMLDocumentStr
Static DocumentsO() As HTMLDocumentStr

   Set DocumentO.DocumentO = Nothing
   DocumentO.WindowH = 0
      
   If Not AddDocument Is Nothing Then
      If CheckForError(SafeArrayGetDim(DocumentsO())) = 0 Then
         ReDim DocumentsO(0 To 0) As HTMLDocumentStr
      Else
         ReDim Preserve DocumentsO(0 To UBound(DocumentsO()) + 1) As HTMLDocumentStr
      End If
      
      Set DocumentsO(UBound(DocumentsO())).DocumentO = AddDocument
      DocumentsO(UBound(DocumentsO())).WindowH = AddWindowH
   ElseIf Not NewErrorMessage = vbNullString Then
      If CheckForError(SafeArrayGetDim(DocumentsO())) = 0 Then
         ReDim DocumentsO(0 To 0) As HTMLDocumentStr
      Else
         ReDim Preserve DocumentsO(0 To UBound(DocumentsO()) + 1) As HTMLDocumentStr
      End If
   
      DocumentsO(UBound(DocumentsO())).ErrorMessage = NewErrorMessage
      DocumentsO(UBound(DocumentsO())).WindowH = AddWindowH
   ElseIf Not Index = NONE Then
      If Not CheckForError(SafeArrayGetDim(DocumentsO())) = 0 Then
         If Index >= LBound(DocumentsO()) And Index <= UBound(DocumentsO()) Then DocumentO = DocumentsO(Index)
      End If
   ElseIf Refresh Then
      Erase DocumentsO()
   End If
   
EndProcedure:
   DocumentList = DocumentO
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Function

'This procedure fills the specified table with a list of any HTML documents found.
Public Sub FillDocumentTable(Table As MSFlexGrid)
On Error GoTo ErrorTrap
Dim Document As HTMLDocumentStr
Dim Index As Long

   With Table
      .rows = 1
      .Row = 0
      
      .Col = 0
      .ColAlignment(0) = flexAlignLeftCenter
      .Text = "Process:"
      
      .Col = 1
      .ColAlignment(1) = flexAlignLeftCenter
      .Text = "Document:"
      
      Index = 0
      Do
         Document = DocumentList(, , Index:=Index)
         If Document.DocumentO Is Nothing Then
            If Document.ErrorMessage = vbNullString Then
               Exit Do
            Else
               .rows = .rows + 1
               .Row = .rows - 1
               
               .CellForeColor = vbRed
               .Col = 0
               .Text = GetWindowProcess(Document.WindowH)
               
               .CellForeColor = vbRed
               .Col = 1
               .Text = Document.ErrorMessage
            End If
         Else
            .rows = .rows + 1
            .Row = .rows - 1
            If Not HasParent(Document.DocumentO) Then
               .Col = 0
               .Text = GetWindowProcess(Document.WindowH)
            End If
            .Col = 1
            .Text = Document.DocumentO.location
         End If
   
         Index = Index + 1
      Loop
   End With
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure fills the specified table with a list of elements found in the specified HTML document.
Public Sub FillElementTable(Table As MSFlexGrid, DocumentO As HTMLDocument)
On Error GoTo ErrorTrap
Dim ItemIndex As Long
Dim NodeIndex As Long

   With Table
      .rows = 1
      .Row = 0
      .Col = 0
      .ColAlignment(0) = flexAlignLeftCenter
      .Text = "Element:"
      .Col = 1
      .ColAlignment(1) = flexAlignLeftCenter
      .Text = "Attributes:"
   End With
   
   With DocumentO.All
      For ItemIndex = 0 To .Length - 1
         With Table
            .rows = .rows + 1
            .Row = .rows - 1
            .Col = 0
         End With
         
         Table.Text = .Item(ItemIndex).tagName
         
         If Not IsNull(.Item(ItemIndex).Attributes) Then
            Table.Col = 1
            Table.Text = vbNullString
            If Not .Item(ItemIndex).Attributes Is Nothing Then
               For NodeIndex = 0 To .Item(ItemIndex).Attributes.Length - 1
                  If Not IsNull(.Item(ItemIndex).Attributes(NodeIndex).nodeValue) Then
                     If IsObject(.Item(ItemIndex).Attributes(NodeIndex).nodeValue) Then
                        If Not .Item(ItemIndex).Attributes(NodeIndex).nodeValue Is Nothing Then
                           Table.Text = Table.Text & .Item(ItemIndex).Attributes(NodeIndex).nodeName & " = " & TypeName(.Item(ItemIndex).Attributes(NodeIndex).nodeValue) & vbCrLf
                        End If
                     Else
                        If Not .Item(ItemIndex).Attributes(NodeIndex).nodeValue = vbNullString Then
                           Table.Text = Table.Text & .Item(ItemIndex).Attributes(NodeIndex).nodeName & " = " & .Item(ItemIndex).Attributes(NodeIndex).nodeValue & vbCrLf
                        End If
                     End If
                  End If
               Next NodeIndex
            End If
         End If
         Table.RowHeight(Table.Row) = Table.Parent.TextHeight(Table.TextMatrix(Table.Row, 1)) * 240
NextItem:
         If DoEvents() = 0 Then Exit For
      Next ItemIndex
   End With
   Exit Sub
   
ErrorTrap:
   If Table.Row > 0 Then
      Table.CellForeColor = vbRed
      Table.Text = "Error: " & CStr(Err.Number) & " - " & Err.Description
   End If
   Resume NextItem
End Sub

'This procedure attempts to retrieve and return the specified frame.
Private Function GetFrame(DocumentO As HTMLDocument, FrameIndex As Long, ErrorMessage As String) As HTMLDocument
On Error GoTo ErrorTrap
Dim Frame As HTMLDocument
   
   ErrorMessage = vbNullString
   Set Frame = Nothing
   Set Frame = DocumentO.frames(FrameIndex).Document
   Set GetFrame = Frame
   Exit Function
   
ErrorTrap:
   ErrorMessage = "Error: " & CStr(Err.Number) & " - " & Err.Description
   Resume Next
End Function

'This procedure returns the process image name for the specified window.
Private Function GetWindowProcess(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim ProcessH As Long
Dim ProcessId As Long
Dim ProcessName As String

   ProcessName = vbNullString
   CheckForError GetWindowThreadProcessId(WindowH, ProcessId)
   
   ProcessH = CheckForError(OpenProcess(PROCESS_QUERY_INFORMATION, CLng(False), ProcessId))
   If Not ProcessH = NO_HANDLE Then
      ProcessName = String$(MAX_PATH, vbNullChar)
      Length = CheckForError(GetProcessImageFileNameW(ProcessH, StrPtr(ProcessName), Len(ProcessName)))
      ProcessName = Left$(ProcessName, Length)
      CheckForError CloseHandle(ProcessH)
   End If
   
EndProcedure:
   GetWindowProcess = ProcessName
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Function


'This procedure handles any active child windows.
Private Function HandleChildWindows(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
   DocumentList AddWindowH:=hwnd, AddDocument:=CheckForDocument(hwnd, lParam)
   
EndProcedure:
   HandleChildWindows = CLng(True)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Function

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Choice As Long
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.HelpContext
   
   On Error GoTo ErrorTrap
   Choice = MsgBox(Description & vbCr & "Error code: " & ErrorCode, vbExclamation Or vbOKCancel Or vbDefaultButton1)
   If Choice = vbCancel Then End
   Exit Sub
   
EndProgram:
   End

ErrorTrap:
   Resume EndProgram
End Sub

'This procedure handles any active windows.
Private Function HandleWindows(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
   DocumentList AddWindowH:=hwnd, AddDocument:=CheckForDocument(hwnd, lParam)
   CheckForError EnumChildWindows(hwnd, AddressOf HandleChildWindows, lParam), ERROR_PROC_NOT_FOUND
   
EndProcedure:
   HandleWindows = CLng(True)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Function

'This procedure checks whether the specified HTML document has a parent.
Private Function HasParent(DocumentO As HTMLDocument)
On Error GoTo ErrorTrap
Dim Has As Boolean
   
   Has = Not (DocumentO.frames.Parent.Document Is DocumentO)

EndProcedure:
   HasParent = Has
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Function

'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   InterfaceWindow.Show
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure scans for HTMl documents in all active windows and any document frames.
Public Sub ScanForDocuments()
On Error GoTo ErrorTrap
Dim DocumentO As HTMLDocumentStr
Dim Index As Long
Dim WMHTMLGetObjectMessage As Long

   WMHTMLGetObjectMessage = CheckForError(RegisterWindowMessageA("WM_HTML_GETOBJECT"))
   
   If Not WMHTMLGetObjectMessage = NO_MESSAGE Then
      DocumentList , , , Refresh:=True
      CheckForError EnumWindows(AddressOf HandleWindows, WMHTMLGetObjectMessage), ERROR_PROC_NOT_FOUND
      
      Index = 0
      Do
         DocumentO = DocumentList(, , Index:=Index)
         If DocumentO.DocumentO Is Nothing Then Exit Do
         If Not HasParent(DocumentO.DocumentO) Then CheckForFrames DocumentO
         Index = Index + 1
      Loop
   End If
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

