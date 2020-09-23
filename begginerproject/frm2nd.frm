VERSION 5.00
Begin VB.Form frm2nd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A Custom Message Box"
   ClientHeight    =   1836
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4992
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1836
   ScaleWidth      =   4992
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFlash 
      Interval        =   150
      Left            =   240
      Top             =   1080
   End
   Begin VB.Shape Shape1 
      Height          =   492
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   852
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Getting fancy with a custom message box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4692
   End
End
Attribute VB_Name = "frm2nd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.BackColor = &HFFFFFF
    Shape1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmmain.Show
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.BackColor = &HFF0000
    Shape1.Visible = True
End Sub

Private Sub tmrFlash_Timer()
    If Label1.BackColor = &HFFFFFF Then
        Label1.BackColor = &HFF0000
            Else
                Label1.BackColor = &HFFFFFF
                    End If
    If Label1.ForeColor = &HFF0000 Then
        Label1.ForeColor = &H0&
            Else
                Label1.ForeColor = &HFF0000
                    End If
End Sub
