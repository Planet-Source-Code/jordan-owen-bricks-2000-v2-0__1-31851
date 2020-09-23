VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brix 2000 v2.0"
   ClientHeight    =   5775
   ClientLeft      =   2235
   ClientTop       =   495
   ClientWidth     =   8760
   ForeColor       =   &H8000000E&
   Icon            =   "WallBall.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "WallBall.frx":030A
   ScaleHeight     =   5775
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   3960
   End
   Begin VB.PictureBox bouncer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3840
      MouseIcon       =   "WallBall.frx":035C
      MousePointer    =   99  'Custom
      Picture         =   "WallBall.frx":04AE
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   1000
      TabIndex        =   5
      Top             =   5450
      Width           =   1545
   End
   Begin VB.PictureBox ball 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4560
      MouseIcon       =   "WallBall.frx":1D50
      MousePointer    =   99  'Custom
      Picture         =   "WallBall.frx":1EA2
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   9440
      TabIndex        =   4
      Top             =   5160
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   3960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   -8520
      TabIndex        =   9
      Top             =   5760
      Width           =   8775
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Next Level..."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   2400
         Width           =   4335
      End
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   0
      MouseIcon       =   "WallBall.frx":22D3
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   8850
   End
   Begin VB.Label tempscores 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   21
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores2 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores3 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores4 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores5 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores6 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores7 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores8 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores9 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores10 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label scores1 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   5880
      Width           =   495
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   62
      Left            =   480
      Picture         =   "WallBall.frx":2425
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   61
      Left            =   7320
      Picture         =   "WallBall.frx":2BFE
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   60
      Left            =   8040
      Picture         =   "WallBall.frx":3337
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   59
      Left            =   5880
      Picture         =   "WallBall.frx":3A7D
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   58
      Left            =   6600
      Picture         =   "WallBall.frx":3FA5
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   57
      Left            =   4440
      Picture         =   "WallBall.frx":4706
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   56
      Left            =   5160
      Picture         =   "WallBall.frx":4F18
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   55
      Left            =   3000
      Picture         =   "WallBall.frx":56AF
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   54
      Left            =   3720
      Picture         =   "WallBall.frx":5D7A
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   53
      Left            =   1560
      Picture         =   "WallBall.frx":654C
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   52
      Left            =   2280
      Picture         =   "WallBall.frx":6D90
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   51
      Left            =   120
      Picture         =   "WallBall.frx":75D6
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   50
      Left            =   840
      Picture         =   "WallBall.frx":7DCE
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   49
      Left            =   8040
      Picture         =   "WallBall.frx":8582
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   48
      Left            =   8040
      Picture         =   "WallBall.frx":8DC6
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   47
      Left            =   8040
      Picture         =   "WallBall.frx":960A
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   46
      Left            =   2400
      Picture         =   "WallBall.frx":9E4E
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   45
      Left            =   4080
      Picture         =   "WallBall.frx":A692
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   44
      Left            =   5640
      Picture         =   "WallBall.frx":AED6
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   43
      Left            =   7200
      Picture         =   "WallBall.frx":B71A
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   42
      Left            =   1080
      Picture         =   "WallBall.frx":BF5E
      Top             =   480
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   41
      Left            =   2400
      Picture         =   "WallBall.frx":C7A2
      Top             =   480
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   40
      Left            =   4080
      Picture         =   "WallBall.frx":CFE6
      Top             =   480
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   39
      Left            =   5640
      Picture         =   "WallBall.frx":D82A
      Top             =   480
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   38
      Left            =   7080
      Picture         =   "WallBall.frx":E06E
      Top             =   480
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   37
      Left            =   0
      Picture         =   "WallBall.frx":E8B2
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   35
      Left            =   0
      Picture         =   "WallBall.frx":F0F6
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   34
      Left            =   0
      Picture         =   "WallBall.frx":F93A
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   36
      Left            =   840
      Picture         =   "WallBall.frx":1017E
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   33
      Left            =   6240
      Picture         =   "WallBall.frx":109C2
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   32
      Left            =   7200
      Picture         =   "WallBall.frx":111BD
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   31
      Left            =   7080
      Picture         =   "WallBall.frx":119B8
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   30
      Left            =   6360
      Picture         =   "WallBall.frx":121B3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   29
      Left            =   7200
      Picture         =   "WallBall.frx":129AE
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   28
      Left            =   6360
      Picture         =   "WallBall.frx":131A9
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   27
      Left            =   7080
      Picture         =   "WallBall.frx":139A4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   26
      Left            =   6720
      Picture         =   "WallBall.frx":1419F
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   25
      Left            =   6240
      Picture         =   "WallBall.frx":1499A
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   24
      Left            =   5040
      Picture         =   "WallBall.frx":15195
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   23
      Left            =   5040
      Picture         =   "WallBall.frx":15990
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   22
      Left            =   5040
      Picture         =   "WallBall.frx":1618B
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   21
      Left            =   5040
      Picture         =   "WallBall.frx":16986
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   20
      Left            =   3840
      Picture         =   "WallBall.frx":17181
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   19
      Left            =   3840
      Picture         =   "WallBall.frx":1797C
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   18
      Left            =   3600
      Picture         =   "WallBall.frx":18177
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   17
      Left            =   5040
      Picture         =   "WallBall.frx":18972
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   16
      Left            =   3720
      Picture         =   "WallBall.frx":1916D
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   15
      Left            =   3600
      Picture         =   "WallBall.frx":19968
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   14
      Left            =   2880
      Picture         =   "WallBall.frx":1A163
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   13
      Left            =   2880
      Picture         =   "WallBall.frx":1A95E
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   12
      Left            =   2880
      Picture         =   "WallBall.frx":1B159
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   11
      Left            =   1560
      Picture         =   "WallBall.frx":1B954
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   10
      Left            =   1560
      Picture         =   "WallBall.frx":1C14F
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   9
      Left            =   1800
      Picture         =   "WallBall.frx":1C94A
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   8
      Left            =   1800
      Picture         =   "WallBall.frx":1D145
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   7
      Left            =   2880
      Picture         =   "WallBall.frx":1D940
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   6
      Left            =   2880
      Picture         =   "WallBall.frx":1E13B
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   5
      Left            =   840
      Picture         =   "WallBall.frx":1E936
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   4
      Left            =   840
      Picture         =   "WallBall.frx":1F131
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   3
      Left            =   840
      Picture         =   "WallBall.frx":1F92C
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   2
      Left            =   840
      Picture         =   "WallBall.frx":20127
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   1
      Left            =   1560
      Picture         =   "WallBall.frx":20922
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image block 
      Height          =   360
      Index           =   0
      Left            =   840
      Picture         =   "WallBall.frx":2111D
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Level: 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label livesleft 
      BackColor       =   &H00000000&
      Caption         =   "Lives: 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label scorecard 
      BackColor       =   &H00000000&
      Caption         =   "Score: 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label bricksleft 
      BackColor       =   &H00000000&
      Caption         =   "Bricks: x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Game Paused - Right Click"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Left Click to Start!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Left Click For Next Life!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim bricks As Integer
Dim lives As Integer
Dim score As Integer
Dim level As Integer
Dim points As Integer
Dim a, MovUD, MovLR
Dim ballspeed As Integer

Public Function FileExists(FileNa As String) As Boolean
    Dim FRes As String
    On Error GoTo NotFound
    FRes = Dir$(FileNa)
    If FRes = "" Then FileExists = False Else FileExists = True
NotFound:
    If Err = 53 Then Resume Next
End Function
Public Function ReadINI(strsection As String, strkey As String, strfullpath As String) As String
   Dim strbuffer As String
   Let strbuffer$ = String$(750, Chr$(0&))
   Let ReadINI$ = Left$(strbuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), "", strbuffer, Len(strbuffer), strfullpath$))
End Function

Public Sub WriteINI(strsection As String, strkey As String, strkeyvalue As String, strfullpath As String)
    Call WritePrivateProfileString(strsection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Sub



Private Sub ball_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Label1.Visible = True Or Label2.Visible = True Then
        Label1.Visible = False
        Label2.Visible = False
        Timer1.Enabled = True
    Else: Exit Sub
    End If
End If
End Sub

Private Sub bouncer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Label1.Visible = True Or Label2.Visible = True Then
        Label1.Visible = False
        Label2.Visible = False
        Timer1.Enabled = True
    Else: Exit Sub
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
    leveldone
End If
If KeyCode = vbKeyF8 Then
    lives = 10
    livesleft.Caption = "Lives: " & lives
End If
If KeyCode = vbKeyEscape Then
    If Timer1.Enabled = False Then
        Unload Me: Form2.Show
        ReleaseCursor Me
    Else: Exit Sub
    End If
End If
End Sub
Public Sub Form_Load()
If FileExists(App.Path & "\brix.ini") Then
    ballspeed = ReadINI("Settings", "Ballspeed", App.Path & "\brix.ini")
    Form1.Caption = "Bricks 2000 v2.0,   Ballspeed: " & ballspeed
End If

    bricks = 50
    bricksleft.Caption = "Bricks: " & 50
    level = 1
    Label4.Caption = "Level: " & level
    score = 0
    scorecard.Caption = "Score: " & score
    lives = 2
    livesleft.Caption = "Lives: " & lives
    Label1.Visible = True
    bouncer.Left = 3600
    bouncer.Visible = True
    ball.Top = bouncer.Top - bouncer.Height + 50
    ball.Left = bouncer.Left - ball.Width / 2 + 770
    ball.Visible = True
    MovUD = -(ballspeed)
    MovLR = (ballspeed)
Dim X As String
X = ReadINI("Skins", "Ball", App.Path & "\brix.ini")
If X = "1200" Then ball.Picture = Form10.Image2.Picture
If X = "1800" Then ball.Picture = Form10.Image3.Picture
If X = "2400" Then ball.Picture = Form10.Image4.Picture
If X = "3000" Then ball.Picture = Form10.Image6.Picture
If X = "3600" Then ball.Picture = Form10.Image7.Picture
If X = "4200" Then ball.Picture = Form10.Image8.Picture
Dim Y As String
Y = ReadINI("Skins", "Paddle", App.Path & "\brix.ini")
If Y = "Paddle1" Then bouncer.Picture = Form10.Image1.Picture
If Y = "Paddle2" Then bouncer.Picture = Form10.Image5.Picture
If Y = "Paddle3" Then bouncer.Picture = Form10.Image9.Picture
If Y = "Paddle4" Then bouncer.Picture = Form10.Image10.Picture
If Y = "Paddle5" Then bouncer.Picture = Form10.Image11.Picture
If Y = "Paddle6" Then bouncer.Picture = Form10.Image12.Picture
Unload Form10

If ballspeed < 10 Then points = 5
If ballspeed > 10 Then points = 10
If ballspeed > 20 Then points = 15
If ballspeed > 30 Then points = 20
If ballspeed > 40 Then points = 25
If ballspeed > 50 Then points = 30
If ballspeed > 60 Then points = 35
If ballspeed = 10 Then points = 5
If ballspeed = 20 Then points = 10
If ballspeed = 30 Then points = 15
If ballspeed = 40 Then points = 20
If ballspeed = 50 Then points = 25
If ballspeed = 60 Then points = 30

Call levels(level)
FormOnTop Me

scores10.Caption = ReadINI("Scores", "Score10", App.Path & "\brix.ini")
scores9.Caption = ReadINI("Scores", "Score9", App.Path & "\brix.ini")
scores8.Caption = ReadINI("Scores", "Score8", App.Path & "\brix.ini")
scores7.Caption = ReadINI("Scores", "Score7", App.Path & "\brix.ini")
scores6.Caption = ReadINI("Scores", "Score6", App.Path & "\brix.ini")
scores5.Caption = ReadINI("Scores", "Score5", App.Path & "\brix.ini")
scores4.Caption = ReadINI("Scores", "Score4", App.Path & "\brix.ini")
scores3.Caption = ReadINI("Scores", "Score3", App.Path & "\brix.ini")
scores2.Caption = ReadINI("Scores", "Score2", App.Path & "\brix.ini")
scores1.Caption = ReadINI("Scores", "Score1", App.Path & "\brix.ini")
tempscores.Caption = ReadINI("Scores", "Tempscore", App.Path & "\brix.ini")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Timer1.Enabled = True Then Cancel = 1
    ReleaseCursor Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
ReleaseCursor Me
Timer1.Enabled = False
Form2.Show
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Label1.Visible = True Or Label2.Visible = True Then
        Label1.Visible = False
        Label2.Visible = False
        Timer1.Enabled = True
    Else: Exit Sub
    End If
End If
If Button = 2 Then
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
        Label3.Visible = True
    Exit Sub
    End If
    If Label3.Visible = True Then
        Timer1.Enabled = True
        Label3.Visible = False
    End If

End If
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.Visible = True Then Exit Sub
bouncer.Left = X - bouncer.Width / 2
If Timer1.Enabled = False Then
ball.Left = X - ball.Width / 2
End If
End Sub
Private Sub bouncer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ball.Tag = "" Then Exit Sub
MoveCursor ball.Tag, 300

End Sub

Public Sub Timer1_Timer()
Do
DoEvents
TrapCursor Me
ball.Top = ball.Top + MovUD
ball.Left = ball.Left + MovLR
For a = 0 To 49
'check bottom of the block
If MovUD < 0 Then
If block(a).Visible = True And ball.Top <= (block(a).Top + block(a).Height) And ball.Top >= block(a).Top And (ball.Left + ball.Width / 2) <= (block(a).Left + block(a).Width) And (ball.Left + ball.Width / 2) >= block(a).Left Then
    MovUD = (ballspeed)
    Call blockhit(block(a))
End If
End If
'check top of the block
If MovUD > 0 Then
If block(a).Visible = True And (ball.Top + ball.Height) >= block(a).Top And (ball.Top + ball.Height) <= (block(a).Top + block(a).Height) And (ball.Left + ball.Width / 2) <= (block(a).Left + block(a).Width) And (ball.Left + ball.Width / 2) >= block(a).Left Then
    MovUD = -(ballspeed)
    Call blockhit(block(a))
End If
End If
'check left of the block
If MovLR > 0 Then
If block(a).Visible = True And (ball.Left + ball.Width) >= block(a).Left And (ball.Left + ball.Width) <= (block(a).Left + block(a).Width) And (ball.Top + ball.Height / 2) <= (block(a).Top + block(a).Height) And (ball.Top + ball.Height / 2) >= block(a).Top Then
    MovLR = -(ballspeed)
    Call blockhit(block(a))
End If
End If
'check right of the block
If MovLR < 0 Then
If block(a).Visible = True And ball.Left <= (block(a).Left + block(a).Width) And ball.Left >= block(a).Left And (ball.Top + ball.Height / 2) <= (block(a).Top + block(a).Height) And (ball.Top + ball.Height / 2) >= block(a).Top Then
    MovLR = (ballspeed)
    Call blockhit(block(a))
End If
End If
Next a
If ball.Top < 0 Then MovUD = (ballspeed)
If ball.Left < 2 Then MovLR = (ballspeed)
If ball.Left > Form1.Width - ball.Width - 50 Then MovLR = -(ballspeed)

If (ball.Top + ball.Height) > bouncer.Top - 10 And (ball.Left + ball.Width) > bouncer.Left And ball.Left + ball.Width / 2 = (bouncer.Left + bouncer.Width / 2) Then
    MovUD = -(ballspeed)
    MovLR = (ballspeed)
End If

If (ball.Top + ball.Height) > bouncer.Top - 10 And (ball.Left + ball.Width) > bouncer.Left - 100 And ball.Left + ball.Width / 2 < (bouncer.Left + bouncer.Width / 4) Then
    MovUD = -(ballspeed)
    MovLR = -(ballspeed) - 10
End If

If (ball.Top + ball.Height) > bouncer.Top - 10 And (ball.Left + ball.Width) > bouncer.Left - 100 And ball.Left + ball.Width / 2 < (bouncer.Left + bouncer.Width / 4 * 2) And ball.Left + ball.Width / 2 > (bouncer.Left + bouncer.Width / 4) Then
    MovUD = -(ballspeed)
    MovLR = -(ballspeed)
End If

If (ball.Top + ball.Height) > bouncer.Top - 10 And ball.Left < (bouncer.Left + bouncer.Width) - 100 And ball.Left + ball.Width / 2 < (bouncer.Left + bouncer.Width / 4 * 3) And ball.Left + ball.Width / 2 > (bouncer.Left + bouncer.Width / 4 * 2) Then
    MovUD = -(ballspeed)
    MovLR = (ballspeed)
End If

If (ball.Top + ball.Height) > bouncer.Top - 10 And ball.Left < (bouncer.Left + bouncer.Width) - 100 And ball.Left + ball.Width / 2 < (bouncer.Left + bouncer.Width / 4 * 4) And ball.Left + ball.Width / 2 > (bouncer.Left + bouncer.Width / 4 * 3) Then
    MovUD = -(ballspeed)
    MovLR = (ballspeed) + 10
End If
If (ball.Top + ball.Height) > bouncer.Top + 300 Then deathhascome
If lives = -1 Then: endofgame: Exit Do
If bricks = 0 Then: leveldone: Exit Do
ReleaseCursor Me
Loop Until Timer1.Enabled = False
End Sub
Public Sub blockhit(block As Object)
    block.Visible = False
    bricks = bricks - 1
    bricksleft.Caption = "Bricks: " & bricks
    score = score + points
    scorecard.Caption = "Score: " & score
    
End Sub
Public Sub endofgame()
    bouncer.Visible = False
    ball.Visible = False
    tempscores.Caption = score
    livesleft.Caption = "Lives: 0"
    checkhighscore
    Timer1.Enabled = False
    Label2.Visible = False
    ReleaseCursor Me
    Form4.Show
    Call death(score)
    Form1.Enabled = False
End Sub
Public Sub leveldone()
    Timer1.Enabled = False
    ball.Visible = False
    score = score + 50
    scorecard.Caption = "Score: " & score
    level = level + 1
    Label4.Caption = "Level: " & level
    bricks = 50
    bricksleft.Caption = "Bricks: " & bricks
    ball.Top = bouncer.Top - bouncer.Height + 50
    ball.Left = bouncer.Left - ball.Width / 2 + 770
    Label2.Visible = False
    Label3.Visible = False
    Label1.Visible = False
    MovUD = -(ballspeed)
    MovLR = (ballspeed)
    ball.Visible = True
If level = 6 Then
    tempscores.Caption = score
    checkhighscore
    Timer1.Enabled = False
    Label2.Visible = False
    lives = 0
    livesleft.Caption = "Lives: " & lives
    ReleaseCursor Me
    Form4.Show
    Call death(score)
    Form1.Enabled = False
End If
    Call levels(level)
End Sub
Public Sub deathhascome()
    Label2.Visible = True
    Timer1.Enabled = False
    score = score - 50
    scorecard.Caption = "Score: " & score
    lives = lives - 1
    livesleft.Caption = "Lives: " & lives
    ball.Top = bouncer.Top - bouncer.Height + 50
    ball.Left = bouncer.Left - ball.Width / 2 + 770
    MovUD = -(ballspeed)
    MovLR = (ballspeed)
End Sub
Public Sub changebrickcount()
If level = 4 Then
bricks = 48
bricksleft.Caption = "Bricks: " & bricks
End If
End Sub
Public Sub checkhighscore()
Dim score1, score2, score3, score4, score5, score6, score7, score8, score9, score10, tempscore As Integer
Call WriteINI("Scores", "TempScore", tempscores.Caption, App.Path & "\brix.ini")
score10 = ReadINI("Scores", "Score10", App.Path & "\brix.ini")
score9 = ReadINI("Scores", "Score9", App.Path & "\brix.ini")
score8 = ReadINI("Scores", "Score8", App.Path & "\brix.ini")
score7 = ReadINI("Scores", "Score7", App.Path & "\brix.ini")
score6 = ReadINI("Scores", "Score6", App.Path & "\brix.ini")
score5 = ReadINI("Scores", "Score5", App.Path & "\brix.ini")
score4 = ReadINI("Scores", "Score4", App.Path & "\brix.ini")
score3 = ReadINI("Scores", "Score3", App.Path & "\brix.ini")
score2 = ReadINI("Scores", "Score2", App.Path & "\brix.ini")
score1 = ReadINI("Scores", "Score1", App.Path & "\brix.ini")
tempscore = ReadINI("Scores", "Tempscore", App.Path & "\brix.ini")
scores1.Caption = score1
scores2.Caption = score2
scores3.Caption = score3
scores4.Caption = score4
scores5.Caption = score5
scores6.Caption = score6
scores7.Caption = score7
scores8.Caption = score8
scores9.Caption = score9
scores10.Caption = score10
tempscores.Caption = tempscore
If tempscore = score1 Then Exit Sub
If tempscore = score2 Then Exit Sub
If tempscore = score3 Then Exit Sub
If tempscore = score4 Then Exit Sub
If tempscore = score5 Then Exit Sub
If tempscore = score6 Then Exit Sub
If tempscore = score7 Then Exit Sub
If tempscore = score8 Then Exit Sub
If tempscore = score9 Then Exit Sub
If tempscore = score10 Then Exit Sub
If tempscore = 0 Then Exit Sub
If tempscore < 0 Then Exit Sub

If tempscore >= score1 Then
Me.Tag = "score1"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score5", scores4.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score4", scores3.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score3", scores2.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score2", scores1.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score1", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score2 Then
Me.Tag = "score2"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score5", scores4.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score4", scores3.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score3", scores2.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score2", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score3 Then
Me.Tag = "score3"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score5", scores4.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score4", scores3.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score3", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score4 Then
Me.Tag = "score4"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score5", scores4.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score4", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score5 Then
Me.Tag = "score5"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score5", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score6 Then
Me.Tag = "score6"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score6", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score7 Then
Me.Tag = "score7"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score7", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score8 Then
Me.Tag = "score8"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score9 Then
Me.Tag = "score9"
Call WriteINI("Scores", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If

If tempscore >= score10 Then
Me.Tag = "score10"
Call WriteINI("Scores", "Score10", tempscores.Caption, App.Path & "\brix.ini")
Exit Sub
End If
End Sub

Private Sub Timer2_Timer()
    GetCursorPosition
    ball.Tag = XPos
End Sub
