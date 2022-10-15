VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   3855
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   20.688
   ScaleMode       =   4  'Character
   ScaleWidth      =   32.125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox StatusBar 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4560
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid DocumentTable 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid ElementTable 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      WordWrap        =   -1  'True
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu ProgramSeparator1Menu 
         Caption         =   "-"
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu DocumentsMainMenu 
      Caption         =   "&Documents"
      Begin VB.Menu RefreshMenu 
         Caption         =   "&Refresh"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This window contains this program's main interface.
Option Explicit

'This procedure gives the command to scan for HTML documents.
Private Sub GetDocumentList()
On Error GoTo ErrorTrap
   Me.MousePointer = vbHourglass
   DocumentTable.Enabled = False
   ElementTable.Enabled = False
   ScanForDocuments
   FillDocumentTable DocumentTable
   DocumentTable.Enabled = True
   ElementTable.Enabled = True
   Me.MousePointer = vbDefault
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure updates the status bar.
Private Sub UpdateStatusBar()
On Error GoTo ErrorTrap
   StatusBar.Text = "Documents: " & CStr(DocumentTable.rows - 1) & " - Elements: " & CStr(ElementTable.rows - 1)
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure gives the command to retrieve the selected HTML document's elements and attributes.
Private Sub DocumentTable_Click()
On Error GoTo ErrorTrap
   If DocumentTable.Row > 0 Then
      Me.MousePointer = vbHourglass
      DocumentTable.Enabled = False
      ElementTable.Enabled = False
      FillElementTable ElementTable, DocumentList(, , Index:=DocumentTable.Row - 1).DocumentO
      DocumentTable.Enabled = True
      ElementTable.Enabled = True
      Me.MousePointer = vbDefault
   End If
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure gives the command to update the status bar.
Private Sub DocumentTable_RowColChange()
On Error GoTo ErrorTrap
   UpdateStatusBar
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure gives the command to update the status bar.
Private Sub ElementTable_RowColChange()
On Error GoTo ErrorTrap
   UpdateStatusBar
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   With App
      Me.Caption = .Title & ", v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With
   
   Me.Width = Screen.Width / 1.5
   Me.Height = Screen.Height / 1.5
   
   GetDocumentList
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure adjusts this window's controls to its new size.
Private Sub Form_Resize()
On Error Resume Next
   DocumentTable.Width = Me.ScaleWidth - 2
   DocumentTable.Height = (Me.ScaleHeight - 1) / 3
   DocumentTable.ColWidth(0) = (Me.Width / 2.1)
   DocumentTable.ColWidth(1) = (Me.Width / 2.1)
   
   ElementTable.Top = DocumentTable.Top + DocumentTable.Height + 1
   ElementTable.Width = Me.ScaleWidth - 2
   ElementTable.Height = (Me.ScaleHeight - 2) / 1.6
   ElementTable.ColWidth(0) = (Me.Width / 4)
   ElementTable.ColWidth(1) = (Me.Width / 1.5)
   
   StatusBar.Top = Me.ScaleHeight - 1.5
   StatusBar.Width = Me.ScaleWidth - 2
End Sub

'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   MsgBox App.Comments, vbInformation
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure closes this window.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

'This procedure gives the command to scan for HTML documents.
Private Sub RefreshMenu_Click()
On Error GoTo ErrorTrap
   GetDocumentList
EndProcedure:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndProcedure
End Sub

