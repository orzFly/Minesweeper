VERSION 5.00
Begin VB.Form frmCustomGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmCustomGame"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   Icon            =   "frmCustomGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "6"
   Begin VB.CommandButton cmdOK 
      Caption         =   "cmdOK"
      Default         =   -1  'True
      Height          =   360
      Left            =   1920
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "cmdCancel"
      Height          =   360
      Left            =   3480
      TabIndex        =   1
      Tag             =   "5"
      Top             =   2760
      Width           =   1455
   End
   Begin orzMinesweeper.ctlLongNumberTextBox txtWidth 
      Height          =   300
      Left            =   120
      TabIndex        =   0
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
   Begin orzMinesweeper.ctlLongNumberTextBox txtHeight 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   1080
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
   Begin orzMinesweeper.ctlLongNumberTextBox txtMaxMinesPerCell 
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   1680
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
   Begin orzMinesweeper.ctlLongNumberTextBox txtMines 
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   2280
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
   Begin VB.Label lblMines 
      AutoSize        =   -1  'True
      Caption         =   "lblMines"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Tag             =   "10"
      Top             =   2040
      Width           =   4560
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMaxMinesPerCell 
      AutoSize        =   -1  'True
      Caption         =   "lblMaxMinesPerCell"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Tag             =   "9"
      Top             =   1440
      Width           =   4560
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      Caption         =   "lblHeight"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Tag             =   "8"
      Top             =   840
      Width           =   4560
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "lblWidth"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Tag             =   "7"
      Top             =   120
      Width           =   4560
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCustomGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    CustomGameResult = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    CustomGameResult = Format(txtWidth.value) & "," & Format(txtHeight.value) & "," & Format(txtMines.value) & "," & Format(txtMaxMinesPerCell.value)
    Unload Me
End Sub

Private Sub Form_Load()
    txtWidth.MinValue = BoardWidthMin
    txtWidth.MaxValue = BoardWidthMax
    txtHeight.MinValue = BoardHeightMin
    txtHeight.MaxValue = BoardHeightMax
    txtMaxMinesPerCell.MinValue = BoardMaxMinesPerCellMin
    txtMaxMinesPerCell.MaxValue = BoardMaxMinesPerCellMax
    txtMines.MinValue = BoardMinesMin
    txtWidth.value = frmGame.Board.Width
    txtHeight.value = frmGame.Board.Height
    txtMaxMinesPerCell.value = frmGame.Board.MaxMinesPerCell
    txtMines.value = frmGame.Board.Mines

    Me.lblWidth.Width = Me.ScaleWidth - 240
    Me.lblHeight.Width = Me.ScaleWidth - 240
    Me.lblMaxMinesPerCell.Width = Me.ScaleWidth - 240
    Me.lblMines.Width = Me.ScaleWidth - 240

    Dim obj As Control
    Me.Caption = LoadResString(Val(Me.Tag))
    For Each obj In Me.Controls
        On Error GoTo hErr
        obj.Text = LoadResString(Val(obj.Tag))
        obj.Caption = LoadResString(Val(obj.Tag))
        obj.Text = LoadResString(Val(obj.HelpContextID))
        obj.Caption = LoadResString(Val(obj.HelpContextID))
    Next

    UpdatePrompt Me.lblWidth, Me.txtWidth.MinValue, Me.txtWidth.MaxValue
    UpdatePrompt Me.lblHeight, Me.txtHeight.MinValue, Me.txtHeight.MaxValue
    UpdatePrompt Me.lblMaxMinesPerCell, Me.txtMaxMinesPerCell.MinValue, Me.txtMaxMinesPerCell.MaxValue
    UpdatePrompt Me.lblMines, Me.txtMines.MinValue, Me.txtMines.MaxValue

    Me.lblWidth.Move 120, 120, Me.ScaleWidth - 240
    Me.txtWidth.Move 120, Me.lblWidth.Top + Me.lblWidth.Height + 120, Me.ScaleWidth - 240
    Me.lblHeight.Move 120, Me.txtWidth.Top + Me.txtWidth.Height + 120, Me.ScaleWidth - 240
    Me.txtHeight.Move 120, Me.lblHeight.Top + Me.lblHeight.Height + 120, Me.ScaleWidth - 240
    Me.lblMaxMinesPerCell.Move 120, Me.txtHeight.Top + Me.txtHeight.Height + 120, Me.ScaleWidth - 240
    Me.txtMaxMinesPerCell.Move 120, Me.lblMaxMinesPerCell.Top + Me.lblMaxMinesPerCell.Height + 120, Me.ScaleWidth - 240
    Me.lblMines.Move 120, Me.txtMaxMinesPerCell.Top + Me.txtMaxMinesPerCell.Height + 120, Me.ScaleWidth - 240
    Me.txtMines.Move 120, Me.lblMines.Top + Me.lblMines.Height + 120, Me.ScaleWidth - 240
    Me.cmdOK.Move Me.ScaleWidth - 240 - Me.cmdCancel.Width * 2, Me.txtMines.Height + Me.txtMines.Top + 120
    Me.cmdCancel.Move Me.cmdOK.Left + Me.cmdOK.Width + 120, Me.cmdOK.Top
    basForm.SetClientRect Me, Me.txtWidth.Width + 240, Me.cmdCancel.Top + Me.cmdCancel.Height + 120
    Exit Sub
hErr:
    Err.Clear
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then CustomGameResult = ""
End Sub

Private Sub txtHeight_Change()
    txtMines.MaxValue = BoardMinesMax(txtWidth.value, txtHeight.value, txtMaxMinesPerCell.value)
    UpdatePrompt Me.lblMines, Me.txtMines.MinValue, Me.txtMines.MaxValue
End Sub

Private Sub txtMaxMinesPerCell_Change()
    txtMines.MaxValue = BoardMinesMax(txtWidth.value, txtHeight.value, txtMaxMinesPerCell.value)
    UpdatePrompt Me.lblMines, Me.txtMines.MinValue, Me.txtMines.MaxValue
End Sub

Private Sub txtWidth_Change()
    txtMines.MaxValue = BoardMinesMax(txtWidth.value, txtHeight.value, txtMaxMinesPerCell.value)
    UpdatePrompt Me.lblMines, Me.txtMines.MinValue, Me.txtMines.MaxValue
End Sub

Private Sub UpdatePrompt(ByRef lblLabel As Label, ByVal intMinValue As Integer, ByVal intMaxValue As Integer)
    lblLabel.Caption = Replace(Replace(LoadResString(Val(lblLabel.Tag)), "%l", Format(intMinValue)), "%u", Format(intMaxValue))
End Sub
