VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transparency Options"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Transparent Windows"
      Height          =   2445
      Left            =   45
      TabIndex        =   4
      Top             =   3735
      Width           =   5550
      Begin VB.CommandButton Command1 
         Caption         =   "Make Normal"
         Height          =   375
         Left            =   4230
         TabIndex        =   6
         Top             =   1935
         Width           =   1230
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   5370
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Applications Currently Running :"
      Height          =   3675
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   5550
      Begin MSComctlLib.Slider Slider1 
         Height          =   285
         Left            =   765
         TabIndex        =   8
         Top             =   3330
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   503
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   70
         TickStyle       =   3
         Value           =   70
      End
      Begin VB.CommandButton Trans 
         Caption         =   "Make Window Transparent"
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   2925
         Width           =   5325
      End
      Begin VB.CommandButton cmdListWindows 
         Caption         =   "List Active Applications"
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   2610
         Width           =   5325
      End
      Begin VB.ListBox listWindows 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   5355
      End
      Begin VB.Label Label1 
         Caption         =   "Opacity:"
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   3330
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdListWindows_Click()
    Set TargetList = frmMain.listWindows
    TargetList.Clear
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub

Private Sub cmdMinimize_Click()
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            ShowWindow listWindows.ItemData(listWindows.ListIndex), SW_Minimize
        End If
    End If
End Sub

Private Sub cmdNormal_Click()
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            ShowWindow listWindows.ItemData(listWindows.ListIndex), SW_Normal
        End If
    End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
m_Trans.Alpha = 100
End Sub

Private Sub Command1_Click()
On Error Resume Next
Set m_Trans = New CTranslucentForm
m_Trans.hWnd = List1.ItemData(List1.ListIndex)
m_Trans.Alpha = 255
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Form_Activate()
cmdListWindows_Click
End Sub

Private Sub Form_Click()
keybd_event vbKeySnapshot, 0, 0&, 0&
DoEvents
Me.Picture = Clipboard.GetData(vbCFBitmap)
SavePicture Me.Picture, "k:\screenshot.bmp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
temp = List1.ListIndex
Slider1.Value = Slider2.Value
Trans_Click
Slider2.Value = Slider1.Value
List1.ListIndex = temp
List1.Selected(temp) = True
End Sub

Private Sub Trans_Click()
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            Load New Form1
        End If
    End If
End Sub
