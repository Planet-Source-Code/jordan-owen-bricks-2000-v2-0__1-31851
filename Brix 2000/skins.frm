VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bricks 2000 v2.0"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "skins.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "skins.frx":030A
   ScaleHeight     =   4035
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image12 
      Height          =   300
      Left            =   4200
      Picture         =   "skins.frx":045C
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Image Image11 
      Height          =   300
      Left            =   2160
      Picture         =   "skins.frx":1CFE
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Image Image10 
      Height          =   300
      Left            =   4200
      Picture         =   "skins.frx":2762
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   135
      Left            =   2640
      TabIndex        =   4
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   2160
      Picture         =   "skins.frx":2C2E
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   4200
      Picture         =   "skins.frx":364A
      Top             =   360
      Width           =   1545
   End
   Begin VB.Image Image8 
      Height          =   225
      Left            =   4320
      Picture         =   "skins.frx":3AFF
      Top             =   2640
      Width           =   225
   End
   Begin VB.Image Image7 
      Height          =   225
      Left            =   3720
      Picture         =   "skins.frx":3C0D
      Top             =   2640
      Width           =   225
   End
   Begin VB.Image Image6 
      Height          =   225
      Left            =   3120
      Picture         =   "skins.frx":3F1F
      Top             =   2640
      Width           =   225
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   2520
      Picture         =   "skins.frx":4039
      Top             =   2640
      Width           =   210
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1920
      Picture         =   "skins.frx":41C3
      Top             =   2640
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   495
      Left            =   1200
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1320
      Picture         =   "skins.frx":45F5
      Top             =   2640
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   555
      Left            =   4080
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2160
      Picture         =   "skins.frx":4A26
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Ball:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Paddle: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
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
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3360
      Width           =   855
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ReadINI(strsection As String, strkey As String, strfullpath As String) As String
   Dim strbuffer As String
   Let strbuffer$ = String$(750, Chr$(0&))
   Let ReadINI$ = Left$(strbuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), "", strbuffer, Len(strbuffer), strfullpath$))
End Function

Public Sub WriteINI(strsection As String, strkey As String, strkeyvalue As String, strfullpath As String)
    Call WritePrivateProfileString(strsection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Sub
Private Sub Form_Load()
FormOnTop Me
Label2.Caption = ReadINI("Skins", "Ball", App.Path & "\brix.ini")
Shape2.Left = Label2.Caption

Label2.Caption = ReadINI("Skins", "Paddle", App.Path & "\brix.ini")
If Label2.Caption = "Paddle1" Then
Shape1.Top = 240
Shape1.Left = 2040
End If
If Label2.Caption = "Paddle2" Then
Shape1.Top = 240
Shape1.Left = 4080
End If
If Label2.Caption = "Paddle3" Then
Shape1.Top = 960
Shape1.Left = 2040
End If
If Label2.Caption = "Paddle4" Then
Shape1.Top = 960
Shape1.Left = 4080
End If
If Label2.Caption = "Paddle5" Then
Shape1.Top = 1680
Shape1.Left = 2040
End If
If Label2.Caption = "Paddle6" Then
Shape1.Top = 1680
Shape1.Left = 4080
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.Visible = True
Label11.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

Form2.Enabled = True
FormOnTop Form2
Unload Me
End Sub

Private Sub Image1_Click()
Shape1.Top = 240
Shape1.Left = 2040
Call WriteINI("Skins", "Paddle", "Paddle1", App.Path & "\brix.ini")
End Sub

Private Sub Image10_Click()
Shape1.Top = 960
Shape1.Left = 4080
Call WriteINI("Skins", "Paddle", "Paddle4", App.Path & "\brix.ini")
End Sub

Private Sub Image11_Click()
Shape1.Top = 1680
Shape1.Left = 2040
Call WriteINI("Skins", "Paddle", "Paddle5", App.Path & "\brix.ini")
End Sub

Private Sub Image12_Click()
Shape1.Top = 1680
Shape1.Left = 4080
Call WriteINI("Skins", "Paddle", "Paddle6", App.Path & "\brix.ini")
End Sub

Private Sub Image5_Click()
Shape1.Top = 240
Shape1.Left = 4080
Call WriteINI("Skins", "Paddle", "Paddle2", App.Path & "\brix.ini")
End Sub

Private Sub Image9_Click()
Shape1.Left = 2040
Shape1.Top = 960
Call WriteINI("Skins", "Paddle", "Paddle3", App.Path & "\brix.ini")
End Sub

Private Sub Label11_Click()
Unload Me
Form2.Enabled = True
FormOnTop Form2
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.Visible = False
Label11.Visible = True
End Sub

Private Sub Image2_Click()
Shape2.Left = 1200
Call WriteINI("Skins", "Ball", "1200", App.Path & "\brix.ini")
End Sub

Private Sub Image3_Click()
Shape2.Left = 1800
Call WriteINI("Skins", "Ball", "1800", App.Path & "\brix.ini")
End Sub

Private Sub Image4_Click()
Shape2.Left = 2400
Call WriteINI("Skins", "Ball", "2400", App.Path & "\brix.ini")
End Sub

Private Sub Image6_Click()
Shape2.Left = 3000
Call WriteINI("Skins", "Ball", "3000", App.Path & "\brix.ini")
End Sub

Private Sub Image7_Click()
Shape2.Left = 3600
Call WriteINI("Skins", "Ball", "3600", App.Path & "\brix.ini")
End Sub

Private Sub Image8_Click()
Shape2.Left = 4200
Call WriteINI("Skins", "Ball", "4200", App.Path & "\brix.ini")
End Sub
