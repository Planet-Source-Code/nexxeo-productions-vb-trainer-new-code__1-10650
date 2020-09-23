VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H80000007&
   Caption         =   "VB Trainer Example by Robert"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   5340
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   3720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ReadProcessMemory sample"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Timer tmrHotKeys 
      Interval        =   1
      Left            =   0
      Top             =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WriteProcessMemory Sample"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblabout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4800
      TabIndex        =   1
      Top             =   4440
      Width           =   420
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB Trainer Example by Robert of Cleris Gaming Inc.

'No explaination needed. Its straight forward

'Note: Stay with the dataypes when you use the functions. The datatype is in the name.
'WriteAByte writes a Byte(duh :)
'ReadAByte returns a byte
'etc

'For best size, use small graphics and compress the exe with PECompact

'PECompact-http://www.collakesoftware.com/

'I also wrote this same code in C, C++, and
'MASM compatible Assembly.
'Contact me at phantom2023@hotmail.com if you
'want the other sources.

'Enjoy

'Robert
'Cleris Gaming Inc.
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Dim f1holder As Integer

'Dim all of you address variables up here

Dim timer_pos As Long

'End Dim list

'API Declaration
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
'I had to alias the ReadProcessMemory API because VB6 thinks "ReadProcessMemory" is ambiguous
Private Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'End API Declaration
Private Sub Command1_Click()

Dim value As Byte
value = 1
Call WriteAByte("Trainer::Trainer", timer_pos, value)
End Sub

Private Sub Command2_Click()
Dim value As Byte
Call ReadAByte("Trainer::Trainer", timer_pos, value)
MsgBox value
End Sub

Private Sub Form_Load()
'Set the equates here

timer_pos = &H40C35C
f1holder = 0
'End equates
End Sub

Private Sub lblabout_Click()
frmabout.Show
End Sub


Private Sub Timer1_Timer()
Command1_Click
End Sub

Private Sub tmrHotKeys_Timer()
'HotKeys go here
'Look under Keyboard Constants for a full list
If GetKeyPress(vbKeyF12) Then
If (f1holder = 0) Then
Command1_Click
f1holder = 1
Timer1.Enabled = True
Exit Sub
End If
If (f1holder = 1) Then
f1holder = 0
Timer1.Enabled = False
Exit Sub
End If
End If
End Sub
Private Function WriteAByte(gamewindowtext As String, address As Long, value As Byte)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
MsgBox "Run the game first", vbCritical, "Error"
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
WriteProcessMemory phandle, address, value, 1, 0&
CloseHandle hProcess
End Function

Private Function WriteAnInt(gamewindowtext As String, address As Long, value As Integer)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
MsgBox "Run the game first", vbCritical, "Error"
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
WriteProcessMemory phandle, address, value, 2, 0&
CloseHandle hProcess
End Function

Private Function WriteALong(gamewindowtext As String, address As Long, value As Long)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
MsgBox "Run the game first", vbCritical, "Error"
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
WriteProcessMemory phandle, address, value, 4, 0&
CloseHandle hProcess
End Function

Private Function ReadAByte(gamewindowtext As String, address As Long, valbuffer As Byte)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
MsgBox "Run the game first", vbCritical, "Error"
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
ReadProcessMem phandle, address, valbuffer, 1, 0&
CloseHandle hProcess
End Function

Private Function ReadAnInt(gamewindowtext As String, address As Long, valbuffer As Integer)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
MsgBox "Run the game first", vbCritical, "Error"
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
ReadProcessMem phandle, address, valbuffer, 2, 0&
CloseHandle hProcess
End Function

Private Function ReadALong(gamewindowtext As String, address As Long, valbuffer As Long)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
MsgBox "Run the game first", vbCritical, "Error"
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
ReadProcessMem phandle, address, valbuffer, 4, 0&
CloseHandle hProcess
End Function

