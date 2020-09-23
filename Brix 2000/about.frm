VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Brix 2000 v2.0"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox P1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      Left            =   120
      ScaleHeight     =   3240
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6480
      Top             =   0
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim thetop As Long
Dim p1hgt As Long
Dim p1wid As Long
Dim theleft As Long
Dim Tempstring As String
Sub Form_Load()
FormOnTop Me
        P1.AutoRedraw = True
        P1.Visible = False
        P1.FontSize = 12
        P1.ForeColor = &HFF00&
        P1.BackColor = BackColor
        P1.ScaleMode = 3
        ScaleMode = 3
        Open (App.Path & "\credits.txt") For Input As #1
        Line Input #1, Tempstring
        P1.Height = (Val(Tempstring) * P1.TextHeight("Test Height")) + 200
        Do Until EOF(1)
            Line Input #1, Tempstring
            PrintText Tempstring
        Loop
        Close #1
        theleft = 0
        thetop = ScaleHeight
        p1hgt = P1.ScaleHeight
        p1wid = P1.ScaleWidth
        Timer1.Enabled = True
        Timer1.Interval = 50
End Sub



Private Sub Form_Unload(Cancel As Integer)
Form2.Enabled = True
FormOnTop Form2
Timer1.Enabled = False
Unload Me
End Sub

Sub Timer1_Timer()
       X% = BitBlt(hDC, theleft, thetop, p1wid, p1hgt, P1.hDC, 0, 0, &HCC0020)
        thetop = thetop - 1
        If thetop < -p1hgt Then
        Timer1.Enabled = False
        CurrentY = ScaleHeight / 2
        CurrentX = (ScaleWidth - TextWidth(Txt$)) / 2
        theleft = 0
        thetop = ScaleHeight
        p1hgt = P1.ScaleHeight
        p1wid = P1.ScaleWidth
        Timer1.Enabled = True
 
        End If
End Sub

Sub PrintText(Text As String)
P1.CurrentX = (P1.ScaleWidth / 2) - (P1.TextWidth(Text) / 2)
P1.ForeColor = 0: X = P1.CurrentX: Y = P1.CurrentY
For i = 1 To 3
    P1.Print Text
    X = X + 1: Y = Y + 1: P1.CurrentX = X: P1.CurrentY = Y
Next i
P1.ForeColor = &HFF00&
P1.Print Text
End Sub
