VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brix 2000 v2.0"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      MaxLength       =   18
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label scores1 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores10 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores9 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores8 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores7 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores6 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores5 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores4 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores3 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label scores2 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label tempscores 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Score: x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Game Over!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2205
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "New Game!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2205
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "New Game!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public Sub Form_Load()
FormOnTop Me

If Form1.Tag <> "" Then
Label7.Top = 840
Label1.Caption = "New High Score!"
Label5.Visible = True
Text1.Visible = True
Label3.Visible = False
Label6.Visible = False
Label2.Visible = True
End If


Form4.AutoRedraw = True
Form4.ScaleMode = 3
Form4.Cls

T3D Form4, Label7, 5, T3dRaiseinset, T3dF1



End Sub

Private Sub Label4_Click()

FormNotOnTop Me
Unload Form1
Unload Me
Form2.Show
Form2.Form_Load
End Sub

Private Sub Label8_Click()
dohighscores
End Sub
Private Sub dohighscores()
Dim score1, score2, score3, score4, score5, score6, score7, score8, score9, score10, tempscore As Integer
score10 = ReadINI("Names", "Score10", App.Path & "\brix.ini")
score9 = ReadINI("Names", "Score9", App.Path & "\brix.ini")
score8 = ReadINI("Names", "Score8", App.Path & "\brix.ini")
score7 = ReadINI("Names", "Score7", App.Path & "\brix.ini")
score6 = ReadINI("Names", "Score6", App.Path & "\brix.ini")
score5 = ReadINI("Names", "Score5", App.Path & "\brix.ini")
score4 = ReadINI("Names", "Score4", App.Path & "\brix.ini")
score3 = ReadINI("Names", "Score3", App.Path & "\brix.ini")
score2 = ReadINI("Names", "Score2", App.Path & "\brix.ini")
score1 = ReadINI("Names", "Score1", App.Path & "\brix.ini")
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

If Text1.Visible = True Then
If Text1.Text = "" Then
MsgBox "Please enter you're name!", vbCritical
Exit Sub
Else

If Form1.Tag = "score1" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score5", scores4.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score4", scores3.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score3", scores2.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score2", scores1.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score1", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score2" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score5", scores4.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score4", scores3.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score3", scores2.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score2", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score3" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score5", scores4.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score4", scores3.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score3", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score4" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score5", scores4.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score4", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score5" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score6", scores5.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score5", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score6" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score7", scores6.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score6", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score7" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", scores7.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score7", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score8" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", scores8.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score9" Then
Call WriteINI("Names", "Score10", scores9.Caption, App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", Text1.Text, App.Path & "\brix.ini")
End If

If Form1.Tag = "score10" Then
Call WriteINI("Names", "Score10", Text1.Text, App.Path & "\brix.ini")
End If
End If
End If
Unload Me
Unload Form1
Form2.Visible = False
Form6.Show
End Sub
Private Sub Label9_Click()

Form1.Enabled = True
Form1.Form_Load
Unload Me
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.Visible = False
Label9.Visible = True
Label3.Visible = True
Label4.Visible = False
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Visible = False
Label8.Visible = True
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Visible = False
Label4.Visible = True
Label6.Visible = True
Label9.Visible = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Form1.Tag = "" Then
Label9.Visible = False
Label6.Visible = True
Label3.Visible = True
Label4.Visible = False
Else
Label2.Visible = True
Label8.Visible = False
End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
dohighscores
End If
End Sub
