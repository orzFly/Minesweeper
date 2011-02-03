Attribute VB_Name = "basMain"
Public SelectGameResult As Long
Public CustomGameResult As String
Public ScreenSaverMode As Boolean

Sub Main()
    ScreenSaverMode = LCase(Command) = "s"
    Load frmGame
    frmGame.Show
End Sub
