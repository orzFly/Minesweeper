VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmAbout"
   ClientHeight    =   5325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8745
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "11"
   Begin VB.TextBox txtAbout 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAbout.frx":000C
      Top             =   120
      Width           =   7935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "cmdOK"
      Default         =   -1  'True
      Height          =   360
      Left            =   7200
      TabIndex        =   1
      Tag             =   "4"
      Top             =   4920
      Width           =   1455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: orzMinesweeper
'Code license: GNU General Public License v3
'Author      : Yeechan Lu a.k.a. orzFly <i@orzfly.com>

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim obj As Control
    Me.Caption = LoadResString(Val(Me.Tag))
    For Each obj In Me.Controls
        On Error GoTo hErr
        obj.Text = LoadResString(Val(obj.Tag))
        obj.Caption = LoadResString(Val(obj.Tag))
        obj.Text = LoadResString(Val(obj.HelpContextID))
        obj.Caption = LoadResString(Val(obj.HelpContextID))
        obj.FontName = LoadResString(198)
        On Error GoTo 0
    Next

    Me.txtAbout.Move 120, 120
    Me.cmdOK.Move Me.ScaleWidth - Me.cmdOK.Width - 120, Me.txtAbout.Height + Me.txtAbout.Top + 120
    basForm.SetClientRect Me, Me.txtAbout.Width + 240, Me.cmdOK.Top + Me.cmdOK.Height + 120
    Me.cmdOK.Move Me.ScaleWidth - Me.cmdOK.Width - 120, Me.txtAbout.Height + Me.txtAbout.Top + 120
    
    Dim strAbout As String
    strAbout = ""
    strAbout = strAbout & IIf(LoadResString(1) = "orzMinesweeper", "orzMinesweeper", LoadResString(1) & "(orzMinesweeper)") & " " & Format(App.Major) & "." & Format(App.Minor) & "." & Format(App.Revision) & vbCrLf
    strAbout = strAbout & "Yet another clone of Minesweeper with many functions extended." & vbCrLf
    strAbout = strAbout & "" & vbCrLf
    strAbout = strAbout & "Copyright (C) 2011 Yeechan Lu a.k.a. orzFly (Founder of orzTech) <i@orzfly.com>" & vbCrLf
    strAbout = strAbout & "" & vbCrLf
    strAbout = strAbout & "This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version." & vbCrLf
    strAbout = strAbout & "" & vbCrLf
    strAbout = strAbout & "This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details." & vbCrLf
    strAbout = strAbout & "" & vbCrLf
    strAbout = strAbout & "You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>." & vbCrLf
    strAbout = strAbout & "" & vbCrLf
    strAbout = strAbout & "The image materials used in this program are licensed under a Creative Commons Attribution-ShareAlike 3.0 Unported License <http://creativecommons.org/licenses/by-sa/3.0/deed.en>." & vbCrLf
    strAbout = strAbout & "" & vbCrLf
    strAbout = strAbout & "You could get a copy of the source code of orzMinesweeper at <http://code.google.com/p/orz-minesweeper/source/checkout>" & vbCrLf
    Me.txtAbout.Text = strAbout
    
    Exit Sub
    
hErr:
    Err.Clear
    Resume Next
End Sub

Private Sub txtAbout_GotFocus()
    HideCaret txtAbout.hwnd
End Sub

Private Sub txtAbout_KeyDown(KeyCode As Integer, Shift As Integer)
    HideCaret txtAbout.hwnd
End Sub

Private Sub txtAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideCaret txtAbout.hwnd
End Sub
