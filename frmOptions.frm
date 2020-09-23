VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toolbox"
   ClientHeight    =   2115
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   2820
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2055
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         DownPicture     =   "frmOptions.frx":000C
         Height          =   350
         Index           =   0
         Left            =   120
         Picture         =   "frmOptions.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   330
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   1
         Left            =   630
         Picture         =   "frmOptions.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   330
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   2
         Left            =   1140
         Picture         =   "frmOptions.frx":092A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   330
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   3
         Left            =   1650
         Picture         =   "frmOptions.frx":0C34
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   330
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   4
         Left            =   2160
         Picture         =   "frmOptions.frx":0F3E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   330
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   5
         Left            =   120
         Picture         =   "frmOptions.frx":1248
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   690
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   6
         Left            =   630
         Picture         =   "frmOptions.frx":1552
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   690
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   7
         Left            =   1140
         Picture         =   "frmOptions.frx":185C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   690
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   8
         Left            =   1650
         Picture         =   "frmOptions.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   690
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   9
         Left            =   2160
         Picture         =   "frmOptions.frx":1E70
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   690
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   10
         Left            =   120
         Picture         =   "frmOptions.frx":217A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1050
         UseMaskColor    =   -1  'True
         Width           =   500
      End
      Begin VB.CommandButton c1 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   11
         Left            =   630
         Picture         =   "frmOptions.frx":2484
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1050
         Width           =   500
      End
      Begin VB.CommandButton c2 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   0
         Left            =   150
         Picture         =   "frmOptions.frx":278E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   500
      End
      Begin VB.CommandButton c2 
         Appearance      =   0  'Flat
         Height          =   350
         Index           =   1
         Left            =   660
         Picture         =   "frmOptions.frx":2A98
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   500
      End
      Begin VB.CommandButton c2 
         BackColor       =   &H000000FF&
         Height          =   345
         Index           =   2
         Left            =   1200
         MaskColor       =   &H8000000A&
         Picture         =   "frmOptions.frx":2DA2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   500
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2640
         Y1              =   1470
         Y2              =   1470
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub c1_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo c1_Click_Err
        '</EhHeader>
100     duz_dev_ediyor = False
102     mymod = Index
        '<EhFooter>
        Exit Sub
c1_Click_Err:
        MsgBox err.Description & vbCrLf & _
           "in Project1.frmOptions.c1_Click " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub c2_Click(Index As Integer)

    If Index = 0 Then MOV = True: start2 = True

    If Index = 1 Then MOV = False: start2 = False

    If Index = 2 Then

        If myobject.Name = Form1.Name Then Exit Sub

        If myobject.Index > 0 Then

            If MsgBox(myobject.Name & " " & myobject.Index - 1 & " isimli nesneyi silmek istiyor musunuz?", vbYesNo, "Silme Onayý") = vbYes Then

                Unload myobject
                Set myobject = Form1

            End If

        End If

    End If

End Sub

Private Sub Form_Activate()

    mysel = False

End Sub

