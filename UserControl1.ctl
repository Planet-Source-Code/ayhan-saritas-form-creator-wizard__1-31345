VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl UserControl1 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   DrawStyle       =   5  'Transparent
   LockControls    =   -1  'True
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   3720
   ScaleWidth      =   4815
   Begin VB.VScrollBar vs 
      Height          =   1215
      Left            =   3810
      Max             =   0
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1020
      Width           =   300
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Height          =   3720
      Left            =   0
      ScaleHeight     =   3660
      ScaleMode       =   0  'User
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   450
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   0
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   450
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "color"
         Top             =   -5000
         Width           =   975
      End
      Begin VB.TextBox lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   -5000
         Width           =   975
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   3180
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontName        =   "Arial"
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_deger
Private m_myform As String
Private m_myvalue As String
Private m_myprop As String
Private Const m_def_rowcount = 0
Private Const m_def_myform = "Form1"
Private Const m_def_myvalue = "Name"
Private Const m_def_myprop = "Name"
'Private m_def_myform   As Form
Private m_rowcount As Long
Private xx As Integer
Public myFontSize As Integer
Public myFontBold As Boolean
Public myFontItalic As Boolean

Public Sub refresh2()

    UserControl_Resize

End Sub

Public Property Get myvalue() As String

    myvalue = m_myvalue

End Property

Public Property Let myvalue(ByVal new_myvalue As String)

    m_myvalue = new_myvalue
    PropertyChanged "m_myvalue"

End Property

Public Property Get myprop() As String

    myprop = m_myprop

End Property

Public Property Let myprop(ByVal new_myprop As String)

    m_myprop = new_myprop
    PropertyChanged "m_myprop"

End Property

Public Property Get myform() As String

    myform = m_myform

End Property

Private Property Get deger() As String

    deger = m_deger

End Property

Public Property Let myform(ByVal new_myform As String)

    m_myform = new_myform
    PropertyChanged "m_myform"

End Property

Public Property Get rowcount() As Long

    rowcount = m_rowcount

End Property

Public Sub deleterows()

    Dim i As Integer
    On Error Resume Next

    For i = 1 To m_rowcount

        Unload lbl(i)
        Unload cmb(i)
        Unload txt(i)

    Next

    vs.Visible = False: m_rowcount = 0:
    vs.Tag = "0": vs.Min = 0: vs.Max = 0: vs.Value = 0
    pic.Height = 0: UserControl.Height = 0: pic.Top = 0
    On Error GoTo 0

End Sub

Public Function addnewrow(rowtype As String, name As String, Optional defval As String, Optional ROWVAL As String) As Long

    Dim z() As String
    Dim i As Integer

    If rowtype = "0" Then 'TEXT

        m_rowcount = m_rowcount + 1
        Load txt(m_rowcount)
        Load lbl(m_rowcount)
        txt(m_rowcount).Text = defval
        lbl(m_rowcount).Text = name
        txt(m_rowcount).Tag = ""
        YERLESTIR lbl(m_rowcount), True, m_rowcount

    ElseIf rowtype = "1" Then 'COMBO

        m_rowcount = m_rowcount + 1
        Load cmb(m_rowcount)
        z = Split(ROWVAL, ";")

        For i = LBound(z) To UBound(z)

            cmb(m_rowcount).AddItem (z(i))

        Next

        Erase z
        Load lbl(m_rowcount)
        lbl(m_rowcount).Text = name
        YERLESTIR lbl(m_rowcount), False, m_rowcount
        On Error Resume Next
        cmb(m_rowcount).Text = defval
        On Error GoTo 0

    ElseIf rowtype = "2" Then 'FILE

        m_rowcount = m_rowcount + 1
        Load txt(m_rowcount)
        txt(m_rowcount).Text = defval
        txt(m_rowcount).Tag = "FILE"
        Load lbl(m_rowcount)
        lbl(m_rowcount).Text = name
        YERLESTIR lbl(m_rowcount), True, m_rowcount

    ElseIf rowtype = "3" Then 'COLOR

        m_rowcount = m_rowcount + 1
        Load txt(m_rowcount)
        Load lbl(m_rowcount)
        txt(m_rowcount).Text = defval
        lbl(m_rowcount).Text = name
        txt(m_rowcount).Tag = "COLOR"
        YERLESTIR lbl(m_rowcount), True, m_rowcount

    ElseIf rowtype = "4" Then 'PRINT

        m_rowcount = m_rowcount + 1
        Load txt(m_rowcount)
        txt(m_rowcount).Tag = "PRINT"
        Load lbl(m_rowcount)
        txt(m_rowcount).Text = defval
        lbl(m_rowcount).Text = name
        YERLESTIR lbl(m_rowcount), True, m_rowcount

    ElseIf rowtype = "5" Then 'FONT

        m_rowcount = m_rowcount + 1
        Load txt(m_rowcount)
        txt(m_rowcount).Tag = "FONT"
        Load lbl(m_rowcount)
        txt(m_rowcount).Text = defval
        lbl(m_rowcount).Text = name
        YERLESTIR lbl(m_rowcount), True, m_rowcount

    End If

    vs.Value = vs.Max

End Function

Private Sub cmb_Click(Index As Integer)

    myvalue = cmb(Index).Text
    myprop = lbl(Index).Text

End Sub

Private Sub cmd_Click()

    If cmd.Tag = "" Then cmd.Visible = False

    If txt(Val(cmd.Tag)).Tag = "FILE" Then

        cd.ShowOpen
        txt(Val(cmd.Tag)).Text = cd.FileName

    ElseIf txt(Val(cmd.Tag)).Tag = "PRINT" Then

        On Error GoTo err
        cd.ShowPrinter
        txt(Val(cmd.Tag)).Text = Printer.DeviceName

    ElseIf txt(Val(cmd.Tag)).Tag = "COLOR" Then

        cd.CancelError = True
        On Error GoTo err
        cd.ShowColor
        txt(Val(cmd.Tag)).Text = cd.Color
err:

    ElseIf txt(Val(cmd.Tag)).Tag = "FONT" Then

        cd.Flags = cdlCFBoth
        On Error GoTo err
        cd.ShowFont
        txt(Val(cmd.Tag)).Text = cd.FontName

        If cd.FontBold = True Then myFontBold = True Else myFontBold = False

        If cd.FontItalic = True Then myFontItalic = True Else myFontItalic = False

        myFontSize = cd.FontSize

    End If

End Sub

Private Sub lbl_GotFocus(Index As Integer)

    On Error Resume Next
    txt(Index).SetFocus
    cmb(Index).SetFocus
    On Error GoTo 0

End Sub

Private Sub txt_Change(Index As Integer)

    myvalue = txt(Index).Text
    myprop = lbl(Index).Text

End Sub

Private Sub txt_GotFocus(Index As Integer)

    If txt(Index).Tag <> "" Then

        cmd.Width = 300
        cmd.Top = txt(Index).Top
        cmd.Height = txt(Index).Height
        cmd.Visible = True: cmd.Tag = Index

        If xx > 0 Then

            cmd.Left = txt(Index).Left + txt(Index).Width - (xx)

        Else

            cmd.Left = txt(Index).Left + txt(Index).Width - 300

        End If

    Else

        cmd.Visible = False

    End If

End Sub

Private Sub txt_LostFocus(Index As Integer)

    myvalue = txt(Index).Text
    myprop = lbl(Index).Text

End Sub

Private Sub UserControl_Initialize()

    'Set m_def_myform = Nothing
    lbl(0).Width = 800
    lbl(0).Height = 315
    txt(0).Width = 800
    txt(0).Height = 315
    pic.Width = lbl(0).Width + txt(0).Width: pic.Visible = True
    pic.Left = 0: pic.Top = 0: UserControl.Height = lbl(0).Height * 8
    pic.Height = Int(m_rowcount * 315) ' UserControl.Height

End Sub

Private Sub m_form(formname As Form)

End Sub

Private Sub YERLESTIR(NESNE As Object, tip As Boolean, Optional IND As Long)

    Dim i As Integer, ypos As Long
    On Error Resume Next
    i = m_rowcount
    NESNE.Left = 0

    If IND > 1 Then

        ypos = lbl(IND - 1).Top + 315

    Else

        ypos = 0

    End If

    NESNE.Top = ypos
    NESNE.Width = 800: NESNE.Height = 300
    NESNE.Visible = True

    If tip = True Then

        txt(IND).Left = 800
        txt(IND).Top = ypos
        txt(IND).Width = 800: txt(IND).Height = 300
        txt(IND).Visible = True

    ElseIf tip = False Then

        cmb(IND).Left = 800
        cmb(IND).Top = ypos
        cmb(IND).Width = 800: cmb(IND).Height = 300
        cmb(IND).Visible = True

    End If

    vs.Max = m_rowcount - Int(pic.Height / 315)

    If vs.Max <= 0 Then

        vs.Max = 0: vs.Min = 0

    Else

        vs.Min = 0: vs.Visible = True: vs.Tag = vs.Value

    End If

    Call UserControl_Resize

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_rowcount = PropBag.ReadProperty("rowcount", m_def_rowcount)
    m_myform = PropBag.ReadProperty("myform", m_def_myform)
    m_myvalue = PropBag.ReadProperty("myvalue", m_def_myvalue)
    m_myprop = PropBag.ReadProperty("myprop", m_def_myprop)

End Sub

Private Sub UserControl_Resize()
    If pic.Visible = True Then pic.SetFocus
    Dim i As Integer
    cmd.Visible = False
    If (m_rowcount * 315) > UserControl.Height Then

        xx = 300
        'If pic.Width = UserControl.Width Then pic.Width = pic.Width - 300
        vs.Visible = True
        vs.Height = UserControl.Height
        vs.Top = 0
        vs.Width = xx
        vs.Left = pic.Width - xx

    Else

        xx = 0
        'If pic.Width < UserControl.Width Then pic.Width = pic.Width + 300
        vs.Visible = False:

    End If

    vs.Max = m_rowcount - Int(UserControl.Height / 315)

    If vs.Max <= 0 Then vs.Max = 0

    vs.Min = 0

    If UserControl.Height < 315 Then UserControl.Height = 315

    If UserControl.Width < 400 Then UserControl.Width = 400

    UserControl.Height = Int(UserControl.Height / 315) * 315
    pic.Width = UserControl.Width
    pic.Height = Int(m_rowcount * 315) 'UserControl.Height '

    For i = 1 To m_rowcount

        If vs.Visible = True Then xx = 300 Else xx = 0

        On Error Resume Next
        lbl(i).Width = (pic.Width / 2) - (xx / 2)
        txt(i).Left = lbl(i).Left + lbl(i).Width + 5
        txt(i).Width = (pic.Width / 2) - (xx / 2) - 5
        cmb(i).Left = lbl(i).Left + lbl(i).Width + 5
        cmb(i).Width = (pic.Width / 2) - (xx / 2) - 5
        'pic.Width = UserControl.Width

    Next

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("rowcount", m_rowcount, m_def_rowcount)
    Call PropBag.WriteProperty("myform", m_myform, m_def_myform)
    Call PropBag.WriteProperty("myvalue", m_myvalue, m_def_myvalue)
    Call PropBag.WriteProperty("myprop", m_myprop, m_def_myprop)

End Sub

Private Sub vs_Change()

    Dim x As Long
    cmd.Visible = False
    Dim i As Integer
    On Error Resume Next
    x = Abs(vs.Value - Val(vs.Tag))

    If vs.Value > Val(vs.Tag) Then

        pic.Top = pic.Top - 315 * x

    Else

        pic.Top = pic.Top + 315 * x

    End If

    vs.Tag = vs.Value
    Exit Sub

    If vs.Value > Val(vs.Tag) Then

        For i = 1 To m_rowcount

            lbl(i).Top = lbl(i).Top - 315 * x
            txt(i).Top = txt(i).Top - 315 * x
            cmb(i).Top = cmb(i).Top - 315 * x

        Next

    Else

        For i = 1 To m_rowcount

            lbl(i).Top = lbl(i).Top + 315 * x
            txt(i).Top = txt(i).Top + 315 * x
            cmb(i).Top = cmb(i).Top + 315 * x

        Next

    End If

    vs.Tag = vs.Value

End Sub

