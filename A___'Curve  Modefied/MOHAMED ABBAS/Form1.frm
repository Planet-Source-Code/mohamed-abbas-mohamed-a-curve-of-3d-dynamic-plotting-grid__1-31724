VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Wave 1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&RUN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&End"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   325
      ScaleMode       =   0  'User
      ScaleWidth      =   429
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00400040&
         Height          =   4935
         Left            =   0
         ScaleHeight     =   4875
         ScaleWidth      =   6435
         TabIndex        =   2
         Top             =   0
         Width           =   6495
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   840
         Top             =   2640
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture2.Visible = False
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
End
End Sub
Private Sub Timer1_Timer()
Dim x, y, offset
Picture1.AutoRedraw = True
Picture1.Cls
Static inte
inte = inte + 0.043
offset = inte * (462 / 49#)
Picture1.Scale (-23, 23)-(23, -23)
Picture1.Cls
For x = -13 / 1.2 To 13 / 1.2 Step 1.2
For y = -13 / 1.2 To 11.9 / 1.2 Step 0.9
Picture1.Line (-x * (4.949 / 7#) + y * (4.949 / 7#), -x * (4.949 / 7#) - y * (4.949 / 7#) + 1.5 * Sin(Sqr(x ^ 2 + y ^ 2) - offset))-(-x * (4.949 / 7#) + (y + 0.9) * (4.949 / 7#), -x * (4.949 / 7#) - (y + 0.9) * (4.949 / 7#) + 1.5 * Sin(Sqr(x ^ 2 + (y + 0.9) ^ 2) - offset)), vb3DLight
Next y
Next x
For y = -13 / 1.2 To 13 / 1.2 Step 1.2
For x = -13 / 1.2 To 11.9 / 1.2 Step 0.9
Picture1.Line (-x * (4.949 / 7#) + y * (4.949 / 7#), -x * (4.949 / 7#) - y * (4.949 / 7#) + 1.5 * Sin(Sqr(x ^ 2 + y ^ 2) - offset))-(-(x + 0.9) * (4.949 / 7#) + (y) * (4.949 / 7#), -(x + 0.9) * (4.949 / 7#) - (y) * (4.949 / 7#) + 1.5 * Sin(Sqr((x + 0.9) ^ 2 + y ^ 2) - offset)), vbWhite
Next x
Next y
Picture1.AutoRedraw = True
End Sub
