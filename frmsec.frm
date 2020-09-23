VERSION 5.00
Object = "*\Apropertiesocx.vbp"
Begin VB.Form frmsec 
   AutoRedraw      =   -1  'True
   Caption         =   "Properties Page"
   ClientHeight    =   3825
   ClientLeft      =   2580
   ClientTop       =   1515
   ClientWidth     =   3555
   Icon            =   "frmsec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin ocxprj.UserControl1 g2 
      Align           =   4  'Align Right
      Height          =   3780
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   6668
      myform          =   ""
      myvalue         =   ""
      myprop          =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   350
      Left            =   1410
      Top             =   1050
   End
End
Attribute VB_Name = "frmsec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

    mysel = True

End Sub

Private Sub Form_Load()
On Error GoTo err
    If Mid(xobj.Name, 1, 3) = "txt" Then txtyukle
    birak = True
    'g2.Refresh2
Exit Sub

err:
Unload Me
End Sub

Public Sub txtyukle()
    'm_rowcount = 0
    'vs.Max = 0: vs.Min = 0: vs.Tag = "0"
    On Error GoTo 0
    'Dim aa As PropertyBag
    g2.Visible = False
    Dim x_obj As TextBox
    'Dim myfr As Form
    '<EhHeader>
    On Error GoTo txtyukle_Err
    '</EhHeader>
    g2.deleterows
    g2.addnewrow "1", "Appereance", xobj.Appearance, "0;1"
    g2.addnewrow "3", "Back Color", xobj.BackColor
    g2.addnewrow "1", "Alignment", xobj.Alignment, "0;1;2"
    g2.addnewrow "1", "Border Style", xobj.BorderStyle, "0;1"
    g2.addnewrow "1", "Data Field", "", "TestField1;TestField2;TestField3;TestField4"
    g2.addnewrow "1", "Enabled", xobj.Enabled, "True;False"
    g2.addnewrow "5", "Font", xobj.Font
    g2.addnewrow "1", "Locked", xobj.Locked, "True;False"
    g2.addnewrow "3", "Fore Color", xobj.ForeColor
    g2.addnewrow "0", "Height", xobj.Height
    g2.addnewrow "0", "Width", xobj.Width
    g2.addnewrow "0", "Top", xobj.Top
    g2.addnewrow "0", "Left", xobj.Left
    g2.addnewrow "0", "Max Lenght", xobj.MaxLength
    g2.addnewrow "1", "Visible", xobj.Visible, "True;False"
    g2.addnewrow "1", "TabStop", xobj.TabStop, "True;False"
    g2.addnewrow "0", "TabIndex", xobj.TabIndex
    g2.Visible = True
    Exit Sub
    '<EhFooter>
    Exit Sub
txtyukle_Err:
    MsgBox err.Description & vbCrLf & _
       "in Project1.frmsec.txtyukle " & _
       "at line " & Erl
    Resume Next
    '</EhFooter>

End Sub

Private Sub Form_Resize()
frmsec.Width = 3600
End Sub

Private Sub Timer1_Timer()

    On Error GoTo err

    If mysel = True Then ' load from properties form

        If g2.myprop = "Appereance" Then

            xobj.Appearance = Val(g2.myvalue)

        ElseIf g2.myprop = "Back Color" Then xobj.BackColor = g2.myvalue

        ElseIf g2.myprop = "Alignment" Then xobj.Alignment = Val(g2.myvalue)

        ElseIf g2.myprop = "Border Style" Then xobj.BorderStyle = g2.myvalue

        ElseIf g2.myprop = "Enabled" Then xobj.Enabled = g2.myvalue

        ElseIf g2.myprop = "Locked" Then xobj.Locked = g2.myvalue

        ElseIf g2.myprop = "Font" Then

            xobj.Font = g2.myvalue
            xobj.FontBold = g2.MYFontBold
            xobj.FontItalic = g2.myFontItalic
            xobj.FontSize = g2.myFontSize

        ElseIf g2.myprop = "Fore Color" Then xobj.ForeColor = g2.myvalue

        ElseIf g2.myprop = "Height" Then xobj.Height = g2.myvalue

        ElseIf g2.myprop = "Width" Then xobj.Width = g2.myvalue

        ElseIf g2.myprop = "Top" Then xobj.Top = g2.myvalue

        ElseIf g2.myprop = "Left" Then xobj.Left = g2.myvalue

        ElseIf g2.myprop = "Max Lenght" Then xobj.MaxLength = g2.myvalue

        ElseIf g2.myprop = "Visible" Then xobj.Visible = g2.myvalue

        ElseIf g2.myprop = "TabStop" Then xobj.TabStop = g2.myvalue

        ElseIf g2.myprop = "TabIndex" Then xobj.TabIndex = g2.myvalue

        ElseIf g2.myprop = "Data Field" Then xobj.DataField = g2.myvalue

        End If

    Else ' save to properties form

    End If

    Exit Sub
err:
    
    If myerr <> err.Description & g2.myprop Then
        myerr = err.Description & g2.myprop
        MsgBox "Invalid Properties Value" & Chr$(13) & g2.myprop, vbExclamation, "Properties Settings"
        Debug.Print err.Description & " -> " & g2.myprop
        End If
    

End Sub

