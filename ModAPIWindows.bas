Attribute VB_Name = "ModAPIWindows"
' --------------------------------------------------------------------------------- '
' "ModAPIWindow.bas"
'
' Please note : Spelling mistakes is FREE! :)
'
' Coburt "Clawy" Jordaan.
' clawy@yahoo.com
' --------------------------------------------------------------------------------- '

Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_Minimize = 6
Public Const SW_Maximize = 3
Public Const SW_Normal = 1

Public TargetList As ListBox

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
' --------------------------------------------------------------------------------- '
' Puts all of the current window captions into a ListBox and the windows "hwnd" into
' that the listbox.itemdata field.
'
' This fucntion is from an example on "Windows API Guide" web site, then modifidy by
' me for put the info into a listbox.
' Windows API Guide - "http://www.vbapi.com"
'
' Put the following lines in your code when you need to get the list:
' ' -------------------------------------------- '
'    Set TargetList = "your list box name"
'    TargetList.Clear
'    EnumWindows AddressOf EnumWindowsProc, 0
' ' -------------------------------------------- '
' --------------------------------------------------------------------------------- '

' Display the title bar text of all top-level windows.  This
' task is given to the callback function, which will receive each handle individually.
' Note that if the window has no title bar text, it will not be displayed (for clarity's sake).

' *** This is the callback function. ***
' This function displays the title bar text of the window identified by hwnd.

  Dim SLength As Long, Buffer As String  ' title bar text length and buffer
  Dim RetVal As Long  ' return value
  Static WinNum As Integer  ' counter keeps track of how many windows have been enumerated

  WinNum = WinNum + 1  ' one more window enumerated....
  SLength = GetWindowTextLength(hWnd) + 1  ' get length of title bar text
  If SLength > 1 Then  ' if return value refers to non-empty string
    Buffer = Space(SLength)  ' make room in the buffer
    RetVal = GetWindowText(hWnd, Buffer, SLength)  ' Get title bar text.
    TargetList.AddItem Left(Buffer, SLength - 1)  '
    TargetList.ItemData(TargetList.NewIndex) = hWnd '
  End If

  EnumWindowsProc = 1  ' return value of 1 means continue enumeration
' --------------------------------------------------------------------------------- '
End Function

