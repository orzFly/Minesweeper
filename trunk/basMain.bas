Attribute VB_Name = "basMain"
'Project name: orzMinesweeper
'Code license: GNU General Public License v3
'Author      : Yeechan Lu a.k.a. orzFly <i@orzfly.com>

Public SelectGameResult As Long
Public CustomGameResult As String
Public ScreenSaverMode As Boolean

Sub Main()
    ScreenSaverMode = LCase(Command) = "s"
    Load frmGame
    frmGame.Show
End Sub
