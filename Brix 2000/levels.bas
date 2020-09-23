Attribute VB_Name = "Module2"
Dim a, b, c, d, e, f, g, h As Integer
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
Public Sub levels(level As Integer)
If level = 1 Then
Form1.Frame1.Left = 0
Form1.Frame1.Top = 0
Form1.Frame1.Visible = True
For b = 0 To 33
Form1.block(b).Picture = Form1.block(51).Picture
Next
For c = 34 To 49
Form1.block(c).Picture = Form1.block(53).Picture
Next
For a = 0 To 49
Form1.block(a).Left = Form7.block(a).Left
Form1.block(a).Top = Form7.block(a).Top
Form1.block(a).Visible = True

Next
Unload Form7
Form1.Frame1.Visible = False
End If

If level = 2 Then
Form1.Frame1.Left = 0
Form1.Frame1.Top = 0
Form1.Frame1.Visible = True
For a = 0 To 5
Form1.block(a).Picture = Form1.block(58).Picture
Next
For b = 6 To 21
Form1.block(b).Picture = Form1.block(60).Picture
Next
For c = 22 To 23
Form1.block(c).Picture = Form1.block(62).Picture
Next
For d = 24 To 30
Form1.block(d).Picture = Form1.block(61).Picture
Next
For e = 31 To 47
Form1.block(e).Picture = Form1.block(62).Picture
Next
For f = 48 To 49
Form1.block(f).Picture = Form1.block(60).Picture
Next
For g = 0 To 49
Form1.block(g).Visible = True
Form1.block(g).Left = Form5.block(g).Left
Form1.block(g).Top = Form5.block(g).Top
Next
Unload Form5
Form1.Label1.Visible = True
Form1.Frame1.Visible = False
End If

If level = 4 Then
Form1.Frame1.Left = 0
Form1.Frame1.Top = 0
Form1.Frame1.Visible = True
For a = 0 To 26
Form1.block(a).Picture = Form1.block(61).Picture
Next
For b = 27 To 30
Form1.block(b).Picture = Form1.block(55).Picture
Next
For c = 33 To 34
Form1.block(c).Picture = Form1.block(60).Picture
Next
For d = 43 To 49
Form1.block(d).Picture = Form1.block(55).Picture
Next
For e = 31 To 32
Form1.block(e).Picture = Form1.block(59).Picture
Next
For f = 35 To 36
Form1.block(f).Picture = Form1.block(59).Picture
Next
For g = 37 To 42
Form1.block(g).Picture = Form1.block(62).Picture
Next

Form1.block(18).Picture = Form1.block(55).Picture

For h = 0 To 49
Form1.block(h).Visible = True
Form1.block(h).Left = Form3.block(h).Left
Form1.block(h).Top = Form3.block(h).Top
Next
Unload Form3
Form1.Label1.Visible = True
Form1.block(27).Visible = False
Form1.block(28).Visible = False
Form1.changebrickcount
Form1.Frame1.Visible = False
End If

If level = 5 Then
Form1.Frame1.Left = 0
Form1.Frame1.Top = 0
Form1.Frame1.Visible = True
For a = 0 To 21
Form1.block(a).Picture = Form1.block(56).Picture
Next
For b = 22 To 37
Form1.block(b).Picture = Form1.block(60).Picture
Next
For c = 38 To 44
Form1.block(c).Picture = Form1.block(61).Picture
Next
For d = 45 To 46
Form1.block(d).Picture = Form1.block(62).Picture
Next
For f = 47 To 49
Form1.block(f).Picture = Form1.block(56).Picture
Next

For e = 0 To 49
Form1.block(e).Left = Form8.block(e).Left
Form1.block(e).Top = Form8.block(e).Top
Form1.block(e).Visible = True
If e = 49 Then
Unload Form8
End If
Next
Form1.Label1.Visible = True
Form1.Frame1.Visible = False
End If

If level = 3 Then
Form1.Frame1.Left = 0
Form1.Frame1.Top = 0
Form1.Frame1.Visible = True
For a = 0 To 9
Form1.block(a).Picture = Form1.block(62).Picture
Next
For b = 10 To 19
Form1.block(b).Picture = Form1.block(60).Picture
Next
For c = 20 To 29
Form1.block(c).Picture = Form1.block(61).Picture
Next
For d = 30 To 40
Form1.block(d).Picture = Form1.block(56).Picture
Next
For f = 42 To 48
Form1.block(f).Picture = Form1.block(56).Picture
Next
Form1.block(41).Picture = Form1.block(61).Picture
Form1.block(49).Picture = Form1.block(60).Picture
For e = 0 To 49
Form1.block(e).Left = Form9.block(e).Left
Form1.block(e).Top = Form9.block(e).Top
Form1.block(e).Visible = True
If e = 49 Then
Unload Form9
End If
Next
Form1.Label1.Visible = True
Form1.Frame1.Visible = False
End If
End Sub
Public Sub death(score As Integer)

Form4.Label7.Caption = "Score: " & score

End Sub
