VERSION 5.00
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3192
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4680
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit

'This procedure gives the command to set the most recently pressed key.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
   LastKeyStroke NewKeyStroke:=CLng(KeyCode)
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to set the most recently selected column of disks.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
   LastColumnSelected NewSelectedColumn:=(X \ SLOT_SIZE)
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

