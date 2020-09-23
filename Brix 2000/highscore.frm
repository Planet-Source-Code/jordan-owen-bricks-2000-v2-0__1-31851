VERSION 5.00
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Brix 2000 High Scores"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   Icon            =   "highscore.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleMode       =   0  'User
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6255
      Left            =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Caption         =   "Clear"
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
      TabIndex        =   1
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Caption         =   "Exit"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Brix 2000 High Scores"
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
      Left            =   720
      TabIndex        =   34
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label names1 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names6 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   32
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names10 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5640
      TabIndex        =   31
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names9 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   30
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names8 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names7 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names5 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names4 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2040
      TabIndex        =   26
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names3 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label names2 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label scores1 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores10 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5640
      TabIndex        =   22
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores9 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores8 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores7 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores6 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores5 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores4 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores3 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label scores2 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label level7 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   13
      Top             =   3600
      Width           =   3960
   End
   Begin VB.Label level8 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   12
      Top             =   4080
      Width           =   3960
   End
   Begin VB.Label level9 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   11
      Top             =   4560
      Width           =   3960
   End
   Begin VB.Label level10 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   10
      Top             =   5040
      Width           =   3960
   End
   Begin VB.Label level2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   9
      Top             =   1200
      Width           =   3960
   End
   Begin VB.Label level3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   8
      Top             =   1680
      Width           =   3960
   End
   Begin VB.Label level4 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   7
      Top             =   2160
      Width           =   3960
   End
   Begin VB.Label level5 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   6
      Top             =   2640
      Width           =   3960
   End
   Begin VB.Label level6 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   540
      TabIndex        =   5
      Top             =   3120
      Width           =   3960
   End
   Begin VB.Label level1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1. Jordan Owen - 200"
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
      Height          =   360
      Left            =   540
      TabIndex        =   4
      Top             =   720
      Width           =   3960
   End
   Begin VB.Image Image1 
      Height          =   5800
      Left            =   200
      Top             =   193
      Width           =   4640
   End
End
Attribute VB_Name = "Form6"
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
Form2.Enabled = False
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

names10.Caption = ReadINI("Names", "Score10", App.Path & "\brix.ini")
names9.Caption = ReadINI("Names", "Score9", App.Path & "\brix.ini")
names8.Caption = ReadINI("Names", "Score8", App.Path & "\brix.ini")
names7.Caption = ReadINI("Names", "Score7", App.Path & "\brix.ini")
names6.Caption = ReadINI("Names", "Score6", App.Path & "\brix.ini")
names5.Caption = ReadINI("Names", "Score5", App.Path & "\brix.ini")
names4.Caption = ReadINI("Names", "Score4", App.Path & "\brix.ini")
names3.Caption = ReadINI("Names", "Score3", App.Path & "\brix.ini")
names2.Caption = ReadINI("Names", "Score2", App.Path & "\brix.ini")
names1.Caption = ReadINI("Names", "Score1", App.Path & "\brix.ini")

level1.Caption = " 1. " & names1.Caption & " - " & scores1.Caption
level2.Caption = " 2. " & names2.Caption & " - " & scores2.Caption
level3.Caption = " 3. " & names3.Caption & " - " & scores3.Caption
level4.Caption = " 4. " & names4.Caption & " - " & scores4.Caption
level5.Caption = " 5. " & names5.Caption & " - " & scores5.Caption
level6.Caption = " 6. " & names6.Caption & " - " & scores6.Caption
level7.Caption = " 7. " & names7.Caption & " - " & scores7.Caption
level8.Caption = " 8. " & names8.Caption & " - " & scores8.Caption
level9.Caption = " 9. " & names9.Caption & " - " & scores9.Caption
level10.Caption = " 10. " & names10.Caption & " - " & scores10.Caption
    T3D Form6, Image1, 200, T3dRaiseinset, T3dF1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Enabled = True
Form2.Show
End Sub

Private Sub Label10_Click()
Form2.Form_Load
Unload Me
End Sub

Private Sub Label11_Click()
Label7.Visible = True
Label11.Visible = False
Dim Response As Long
Dim Response1 As Long

Response = MessageBox(Me.hwnd, "Are you sure you want to clear the scores?", _
            "Brix 2000", MB_YESNO Or _
                MB_ICONQUESTION Or MB_TASKMODAL)

Select Case Response
Case Is = IDYES
        Response1 = MessageBox(Me.hwnd, "Scores cleared", _
                    "Brix 2000", MB_OK Or MB_ICONINFORMATION)
            
If Response1 = IDOK Then
Call WriteINI("Scores", "Score1", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score2", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score3", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score4", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score5", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score6", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score7", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score8", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score9", "0", App.Path & "\brix.ini")
Call WriteINI("Scores", "Score10", "0", App.Path & "\brix.ini")

Call WriteINI("Names", "Score1", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score2", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score3", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score4", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score5", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score6", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score7", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score8", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score9", "Empty", App.Path & "\brix.ini")
Call WriteINI("Names", "Score10", "Empty", App.Path & "\brix.ini")
Form_Load
End If
            
Case Is = IDNO
       Exit Sub
End Select
End Sub


Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Visible = False
Label11.Visible = True
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.Visible = False
Label10.Visible = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Visible = True
Label11.Visible = False
Label10.Visible = False
Label8.Visible = True
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Visible = True
Label11.Visible = False
Label10.Visible = False
Label8.Visible = True
End Sub

Private Sub level10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Visible = True
Label11.Visible = False
Label10.Visible = False
Label8.Visible = True
End Sub
