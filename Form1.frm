VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Transparent Form Maker"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Trans As CTranslucentForm

Private Sub Form_Load()
On Error Resume Next
Set m_Trans = New CTranslucentForm
m_Trans.hWnd = frmMain.listWindows.ItemData(frmMain.listWindows.ListIndex)
m_Trans.Alpha = frmMain.Slider1.Value / 100 * 255
For i = 0 To frmMain.List1.ListCount - 1
frmMain.List1.ListIndex = i
If frmMain.List1.ItemData(i) = frmMain.listWindows.ItemData(frmMain.listWindows.ListIndex) Then frmMain.List1.RemoveItem i
Next i
frmMain.List1.AddItem frmMain.listWindows.Text & " (" & Int(frmMain.Slider1.Value / frmMain.Slider1.Max * 100) & "%)"
frmMain.List1.ItemData(frmMain.List1.NewIndex) = frmMain.listWindows.ItemData(frmMain.listWindows.ListIndex)
End Sub
