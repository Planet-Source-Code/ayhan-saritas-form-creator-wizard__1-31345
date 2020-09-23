VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   420
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   30
      Width           =   345
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Komut Butonu"
      Height          =   405
      Index           =   0
      Left            =   3360
      TabIndex        =   10
      Top             =   90
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1980
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   60
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CheckBox chk 
      Caption         =   "Check1"
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   780
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.OptionButton opt 
      Caption         =   "Option1"
      Height          =   285
      Index           =   0
      Left            =   1260
      TabIndex        =   6
      Top             =   780
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   0
      Left            =   3660
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   780
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ListBox lst 
      Height          =   1425
      Index           =   0
      Left            =   3660
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.HScrollBar hscr 
      Height          =   255
      Index           =   0
      Left            =   3630
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.VScrollBar vscr 
      Height          =   1455
      Index           =   0
      Left            =   5070
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox pic 
      Height          =   795
      Index           =   0
      Left            =   2340
      ScaleHeight     =   735
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   30
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   345
   End
   Begin VB.Label lbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   1140
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Line lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   1110
      X2              =   1650
      Y1              =   1200
      Y2              =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X1 As Single
Dim Y1 As Single

Private Sub cmb_GotFocus(Index As Integer)

    Set xobj = cmb(Index)

End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
100     frmsec.Show vbModeless, Me
        '<EhFooter>
        Exit Sub
Command1_Click_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.Command1_Click " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub Form_Activate()

    mysel = False

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Set myobject = Form1
        'myobject.Name = Form1.Name
102     start2 = True
        '<EhFooter>
        Exit Sub
Form_Load_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.Form_Load " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

        Set xobj = Nothing
        'If Button <> 1 Then start2 = False
        '<EhHeader>
        On Error GoTo Form_MouseDown_Err
        '</EhHeader>
        Dim nParam As Long

100     If Button = 1 And mymod > 0 And duz_dev_ediyor = False Then

            'shp1.Picture = txt(0)
102         myx = X: myy = Y
104         nesne_yap
            'shp1.Visible = True
106         GoTo a
108         shp1.Top = Y
110         shp1.Left = X
112         shp1.Width = 50
114         shp1.Height = 50
a:
116         duz_dev_ediyor = True
118         MOV = True
120         start = True
122         Form1.MousePointer = vbNormal
124         myobject.MousePointer = 8
            Exit Sub
126         nParam = 17

128         If Mid(myobject.Name, 1, 3) = "lbl" Then Exit Sub

130         Call ReleaseCapture
132         Call SendMessage(myobject.hwnd, WM_NCLBUTTONDOWN, nParam, 0)

        End If

        '<EhFooter>
        Exit Sub
Form_MouseDown_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.Form_MouseDown " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo Form_MouseMove_Err
        '</EhHeader>
100     Me.MousePointer = vbNormal
        '<EhFooter>
        Exit Sub
Form_MouseMove_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.Form_MouseMove " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub nesne_yap()

        '<EhHeader>
        On Error GoTo nesne_yap_Err
        '</EhHeader>

100     Select Case mymod

            Case 1
102             Load lbl(MaxId(mymod) + 1)
104             lbl(MaxId(mymod) + 1).Top = myy
106             lbl(MaxId(mymod) + 1).Left = myx
                'lbl(MaxId(mymod) + 1).Height = -300
                'lbl(MaxId(mymod) + 1).Width = -800
108             lbl(MaxId(mymod) + 1).Visible = True
110             lbl(MaxId(mymod) + 1).Caption = "Etiket_" & MaxId(mymod)
112             Set myobject = lbl(MaxId(mymod) + 1)
114             MaxId(mymod) = MaxId(mymod) + 1

116         Case 2
118             Load txt(MaxId(mymod) + 1)
120             txt(MaxId(mymod) + 1).Top = myy
122             txt(MaxId(mymod) + 1).Left = myx
                'txt(MaxId(mymod) + 1).Height = 50
                'txt(MaxId(mymod) + 1).Width = 50
124             txt(MaxId(mymod) + 1).Visible = True: ' shp1.Visible = False
126             txt(MaxId(mymod) + 1).Text = "Tekst_" & MaxId(mymod)
128             Set myobject = txt(MaxId(mymod) + 1)
130             Set xobj = txt(MaxId(mymod) + 1)
                MaxId(mymod) = MaxId(mymod) + 1
                frmsec.txtyukle

132         Case 3

134         Case Else

        End Select

        '<EhFooter>
        Exit Sub
nesne_yap_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.nesne_yap " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err
        '</EhHeader>
100     frmOptions.Show vbModeless, Me
        '<EhFooter>
        Exit Sub
Command2_Click_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.Command2_Click " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        Form1.Tag = lbl(Index)
        Set xobj = lbl(Index)
        '<EhHeader>
        On Error GoTo lbl_MouseDown_Err
        '</EhHeader>

        If MOV = False And start2 = False Then lbl(Index).MousePointer = vbNormal: Exit Sub

100     setmyobject = lbl(Index)
102     X1 = X
104     Y1 = Y
106     start = True
        '<EhFooter>
        Exit Sub
lbl_MouseDown_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.lbl_MouseDown " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo lbl_MouseMove_Err
        '</EhHeader>

100     If start2 = False Then Exit Sub

102     If myobject.Name <> lbl(Index).Name And myobject.Name <> Form1.Name Then Exit Sub

        'If Button = 0 Then MOV = False: lbl(Index).MousePointer = vbNormal
104     Set myobject = lbl(Index)
        Dim X2, Y2 As Single
        On Error GoTo hell:

112     If ((Y < lbl(Index).Height + 150 And Y > lbl(Index).Height - 150) And (X < lbl(Index).Width + 150 And X > lbl(Index).Width - 150)) Or MOV = True Then

            'If ((Y > lbl(Index).Height - 150) And (X > lbl(Index).Width - 150)) Or MOV = True Then
114         lbl(Index).MousePointer = 8

        Else

116         lbl(Index).MousePointer = 5

        End If

118     If Button = 1 Then

120         X2 = X
122         Y2 = Y

124         If lbl(Index).MousePointer = 8 Then

                '  lbl(Index).Appearance = 0
126             MOV = True
128             lbl(Index).Width = X2: lbl(Index).Height = Y2

130         ElseIf lbl(Index).MousePointer = 5 Then

132             With lbl(Index)

134                 .Move lbl(Index).Left - X1 + X2, lbl(Index).Top - Y1 + Y2

                End With

                'MOV = False

            End If

        End If

hell:
        Exit Sub
        '<EhFooter>
        Exit Sub
lbl_MouseMove_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.lbl_MouseMove " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo lbl_MouseUp_Err
        '</EhHeader>

100     If start2 = False Then Exit Sub

102     start = False: MOV = False: lbl(Index).MousePointer = vbNormal: Set myobject = Form1
104     duz_dev_ediyor = False: mymod = 0: ' txt(Index).Appearance = 1
        '<EhFooter>
        Exit Sub
lbl_MouseUp_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.lbl_MouseUp " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo txt_MouseDown_Err
        '</EhHeader>
        On Error Resume Next

        If frmsec.Visible = True And (xobj <> txt(Index)) Then

            On Error GoTo txt_MouseDown_Err
            Set xobj = txt(Index)
            frmsec.txtyukle

        End If

100     If start2 = False Then Exit Sub

102     Set myobject = txt(Index)
        Set xobj = txt(Index)
104     X1 = X
106     Y1 = Y
108     start = True
        '<EhFooter>
        Exit Sub
txt_MouseDown_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.txt_MouseDown " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        If MOV = False And start2 = False Then txt(Index).MousePointer = vbNormal: Exit Sub

        '<EhHeader>
        On Error GoTo txt_MouseMove_Err
        '</EhHeader>

100     If start2 = False Then Exit Sub

102     If myobject.Name <> txt(Index).Name And myobject.Name <> Form1.Name Then Exit Sub

        'If Button = 0 Then MOV = False: txt(Index).MousePointer = vbNormal
104     Set myobject = txt(Index)
        
        Dim X2, Y2 As Single
        On Error GoTo hell:

112     If ((Y < txt(Index).Height + 150 And Y > txt(Index).Height - 150) And (X < txt(Index).Width + 150 And X > txt(Index).Width - 150)) Or MOV = True Then

            'If ((Y > txt(Index).Height - 150) And (X > txt(Index).Width - 150)) Or MOV = True Then
114         txt(Index).MousePointer = 8

        Else

116         txt(Index).MousePointer = 5

        End If

118     If Button = 1 Then

120         X2 = X
122         Y2 = Y

124         If txt(Index).MousePointer = 8 Then

                '  txt(Index).Appearance = 0
126             MOV = True
128             txt(Index).Width = X2: txt(Index).Height = Y2

130         ElseIf txt(Index).MousePointer = 5 Then

132             With txt(Index)

134                 .Move txt(Index).Left - X1 + X2, txt(Index).Top - Y1 + Y2

                End With

                'MOV = False

            End If

        End If

hell:
        Exit Sub
        '<EhFooter>
        Exit Sub
txt_MouseMove_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.txt_MouseMove " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo txt_MouseUp_Err
        '</EhHeader>

100     If start2 = False Then Exit Sub

102     start = False: MOV = False: txt(Index).MousePointer = vbNormal: Set myobject = Form1
104     duz_dev_ediyor = False: mymod = 0: ' txt(Index).Appearance = 1
        '<EhFooter>
        Exit Sub
txt_MouseUp_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.Form1.txt_MouseUp " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

