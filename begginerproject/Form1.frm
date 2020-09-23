VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmmain 
   Caption         =   "Newbies Learning"
   ClientHeight    =   5424
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8124
   LinkTopic       =   "Form1"
   ScaleHeight     =   5424
   ScaleWidth      =   8124
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Sroll Bars"
      Height          =   3012
      Left            =   5640
      TabIndex        =   23
      Top             =   2160
      Width           =   2412
      Begin VB.HScrollBar HScroll1 
         Height          =   252
         Left            =   360
         Max             =   100
         Min             =   1
         TabIndex        =   25
         Top             =   2400
         Value           =   1
         Width           =   1812
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1572
         Left            =   240
         Max             =   100
         Min             =   1
         TabIndex        =   24
         Top             =   600
         Value           =   1
         Width           =   252
      End
      Begin VB.Label lblHScroll 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   372
         Left            =   1440
         TabIndex        =   27
         Top             =   1920
         Width           =   372
      End
      Begin VB.Label lblVScroll 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   252
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   492
      End
   End
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7560
      Top             =   360
   End
   Begin VB.Frame Frame6 
      Caption         =   "Porgress bar"
      Height          =   1452
      Left            =   5520
      TabIndex        =   20
      Top             =   120
      Width           =   2532
      Begin VB.CommandButton cmdstart 
         Caption         =   "Start"
         Height          =   372
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   972
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   372
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   2172
         _ExtentX        =   3831
         _ExtentY        =   656
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Message Boxes"
      Height          =   3132
      Left            =   2760
      TabIndex        =   14
      Top             =   2040
      Width           =   2652
      Begin VB.CommandButton cmdCustom 
         Caption         =   "Custom"
         Height          =   372
         Left            =   480
         TabIndex        =   19
         Top             =   2280
         Width           =   1332
      End
      Begin VB.CommandButton cmdRentry 
         Caption         =   "VbRentry Ignore"
         Height          =   372
         Left            =   480
         TabIndex        =   18
         Top             =   1800
         Width           =   1332
      End
      Begin VB.CommandButton cmdOkCancel 
         Caption         =   "VbOkCancel"
         Height          =   372
         Left            =   480
         TabIndex        =   17
         Top             =   1320
         Width           =   1332
      End
      Begin VB.CommandButton cmdYesNo 
         Caption         =   "VbYesNo"
         Height          =   372
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   1332
      End
      Begin VB.CommandButton cmdVbOk 
         Caption         =   "VbOk"
         Height          =   372
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   1332
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Text"
      Height          =   3132
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2532
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Text"
         Height          =   372
         Left            =   1320
         TabIndex        =   13
         Top             =   2640
         Width           =   852
      End
      Begin VB.ListBox List1 
         Height          =   1200
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1692
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   360
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1680
         Width           =   1692
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Text"
         Height          =   372
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   852
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   240
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2280
         Width           =   2052
      End
   End
   Begin VB.Frame Frame3 
      Height          =   732
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   2652
      Begin VB.CommandButton cmdmouseover 
         Caption         =   "Mouse Over"
         Height          =   372
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mouse Over"
      Height          =   612
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   2652
      Begin VB.Label lblMouseOver 
         Caption         =   "Mouse Over"
         Height          =   252
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time and Date"
      Height          =   1812
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2532
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1920
         Top             =   120
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "&Stop"
         Height          =   372
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   972
      End
      Begin VB.CommandButton cmdTime 
         Caption         =   "&Start"
         Height          =   372
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Height          =   372
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1932
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    List1.AddItem Text1.Text
        Combo1.AddItem Text1.Text
End Sub

Private Sub cmdClear_Click()
    List1.Clear
        Combo1.Clear
End Sub

Private Sub cmdCustom_Click()
    frmmain.Hide
        frm2nd.Show
End Sub

Private Sub cmdDate_Click()
    Timer1.Enabled = False
        lblTime.Caption = ""
End Sub


Private Sub cmdMouseOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdmouseover.Caption = "click here"
End Sub

Private Sub cmdOkCancel_Click()
    MsgBox "This is a ok cancel box", vbOKCancel, "With a title"
End Sub

Private Sub cmdRentry_Click()
    MsgBox "This is rentry or ignore box"
End Sub

Private Sub cmdstart_Click()
    tmrProgress.Enabled = True
End Sub

Private Sub cmdTime_Click()
    Timer1.Enabled = True
End Sub

Private Sub cmdVbOk_Click()
    MsgBox "this is just using ok", vbCritical
End Sub

Private Sub cmdYesNo_Click()
   If MsgBox("This is a yes or no box", vbYesNo) = vbYes Then
        MsgBox "User clicked Yes"
            Else
                MsgBox "User clicked no"
                    End If
End Sub

Private Sub Command1_Click()
    Combo1.AddItem Text1.Text
        List1.AddItem Text1.Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdmouseover.Caption = "Mouse Over"
        lblMouseOver.ForeColor = &H0&
End Sub


Private Sub HScroll1_Change()
    lblHScroll.Caption = HScroll1.Value
End Sub

Private Sub lblMouseOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMouseOver.ForeColor = &HFF&
End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = Now
End Sub

Private Sub tmrProgress_Timer()
    ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = 100 Then
            ProgressBar1.Value = 0
            tmrProgress.Enabled = False
                End If
End Sub

Private Sub VScroll1_Change()
    lblVScroll.Caption = VScroll1.Value
End Sub
