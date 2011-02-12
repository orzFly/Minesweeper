VERSION 5.00
Begin VB.Form frmSelectGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmSelectGame"
   ClientHeight    =   1335
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3840
   Icon            =   "frmSelectGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "2"
   Begin orzMinesweeper.ctlLongNumberTextBox txtGameID 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "cmdCancel"
      Height          =   360
      Left            =   3240
      TabIndex        =   3
      Tag             =   "5"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "cmdOK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2040
      TabIndex        =   2
      Tag             =   "4"
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "lblTitle"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Tag             =   "3"
      Top             =   120
      Width           =   4560
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSelectGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    SelectGameResult = -1
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SelectGameResult = txtGameID.value
    Unload Me
End Sub

Private Sub Form_Load()
    txtGameID.MaxValue = GameIDMax
    txtGameID.MinValue = GameIDMin
    txtGameID.Random

    Me.lblTitle.Width = Me.ScaleWidth - 240

    Dim obj As Control
    Me.Caption = LoadResString(Val(Me.Tag))
    For Each obj In Me.Controls
        On Error GoTo hErr
        obj.Text = LoadResString(Val(obj.Tag))
        obj.Caption = LoadResString(Val(obj.Tag))
        obj.Text = LoadResString(Val(obj.HelpContextID))
        obj.Caption = LoadResString(Val(obj.HelpContextID))
    Next

    Me.lblTitle.Caption = Replace(Replace(Me.lblTitle.Caption, "%l", Format(txtGameID.MinValue)), "%u", Format(txtGameID.MaxValue))

    Me.lblTitle.Move 120, 120, Me.ScaleWidth - 240
    Me.txtGameID.Move 120, Me.lblTitle.Top + Me.lblTitle.Height + 120, Me.ScaleWidth - 240
    Me.cmdOK.Move Me.ScaleWidth - 240 - Me.cmdCancel.Width * 2, Me.txtGameID.Height + Me.txtGameID.Top + 120
    Me.cmdCancel.Move Me.cmdOK.Left + Me.cmdOK.Width + 120, Me.cmdOK.Top
    basForm.SetClientRect Me, Me.txtGameID.Width + 240, Me.cmdCancel.Top + Me.cmdCancel.Height + 120
    Exit Sub
hErr:
    Err.Clear
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then SelectGameResult = -1
End Sub
