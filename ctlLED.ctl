VERSION 5.00
Begin VB.UserControl ctlLED 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ctlLED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project name: orzMinesweeper
'Code license: GNU General Public License v3
'Author      : Yeechan Lu a.k.a. orzFly <i@orzfly.com>

Public Enum enumLEDForeColor
    ledfcRed = 0
    ledfcGreen = 1
    ledfcBlue = 2
    ledfcYellow = 3
    ledfcCyan = 4
    ledfcPurple = 5
    ledfcWhite = 6
End Enum

Public Enum enumLEDBorderStyle
    ledbsNone = 0
    ledbsFixedSingle = 1
End Enum

Const m_def_ForeColor = enumLEDForeColor.ledfcRed
Const m_def_Text = "    "
Const m_def_DefaultText = " "
Const m_def_MaxLength = 4
Const m_LEDPictureChar = "- 9876543210:."
Dim m_ForeColor As enumLEDForeColor
Dim m_Text As String
Dim m_MaxLength As Byte
Dim m_DefaultText As String
Dim m_LEDPicture As StdPicture
Dim ForcePaint As Boolean
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event DblClick()
Event Resize()

Public Property Get BorderStyle() As enumLEDBorderStyle
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As enumLEDBorderStyle)
    UserControl.BorderStyle() = New_BorderStyle
    UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    UserControl.AutoRedraw = True
    Set m_LEDPicture = LoadResPicture(1, vbResBitmap)
    ForcePaint = True
    UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get ForeColor() As enumLEDForeColor
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As enumLEDForeColor)
    m_ForeColor = New_ForeColor
    UserControl_Resize
    PropertyChanged "ForeColor"
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Dim I As Integer, blnMinusSign As Boolean
    m_Text = ""
    For I = 1 To Len(New_Text)
        If InStr(1, m_LEDPictureChar, Mid(New_Text, I, 1)) > 0 Then
            m_Text = m_Text & Mid(New_Text, I, 1)
        End If
    Next I
    If m_DefaultText = "0" And Left(m_Text, 1) = "-" Then
        blnMinusSign = True
        m_Text = Mid(m_Text, 2)
        m_Text = "-" & Right(String(m_MaxLength, DefaultText) & m_Text, m_MaxLength - 1)
    Else
        m_Text = Right(String(m_MaxLength, DefaultText) & m_Text, m_MaxLength)
    End If
    UserControl_Paint
    PropertyChanged "Text"
End Property

Public Property Get DefaultText() As String
    If Len(m_DefaultText) = 0 Then
        DefaultText = m_def_DefaultText
    Else
        DefaultText = Left(m_DefaultText, 1)
    End If
    m_DefaultText = DefaultText
End Property

Public Property Let DefaultText(ByVal New_DefaultText As String)
    m_DefaultText = Left(New_DefaultText, 1)
    If InStr(1, m_LEDPictureChar, m_DefaultText) < 1 Then
        m_DefaultText = m_def_DefaultText
    End If
    Text = Text
    PropertyChanged "Text"
End Property

Public Property Get MaxLength() As Byte
    MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Byte)
    m_MaxLength = New_MaxLength
    Text = Text
    UserControl_Resize
    PropertyChanged "MaxLength"
End Property

Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    m_Text = m_def_Text
    m_MaxLength = m_def_MaxLength
    ForcePaint = True
    UserControl_Paint
End Sub

Private Sub UserControl_Paint()
    Static strLastText As String
    If ForcePaint = False Then
        If strLastText = m_Text Then Exit Sub
    Else
        ForcePaint = False
        For I = 1 To Len(m_Text)
            UserControl.PaintPicture m_LEDPicture, I * 13 * Screen.TwipsPerPixelX - 13 * Screen.TwipsPerPixelX, 0, 13 * Screen.TwipsPerPixelX, 23 * Screen.TwipsPerPixelY, m_ForeColor * 13 * Screen.TwipsPerPixelX, (InStr(1, m_LEDPictureChar, Mid(m_Text, I, 1)) - 1) * 23 * Screen.TwipsPerPixelY, 13 * Screen.TwipsPerPixelX, 23 * Screen.TwipsPerPixelY
        Next I
        Exit Sub
    End If
    If Len(strLastText) <> Len(m_Text) Then
        For I = 1 To Len(m_Text)
            UserControl.PaintPicture m_LEDPicture, I * 13 * Screen.TwipsPerPixelX - 13 * Screen.TwipsPerPixelX, 0, 13 * Screen.TwipsPerPixelX, 23 * Screen.TwipsPerPixelY, m_ForeColor * 13 * Screen.TwipsPerPixelX, (InStr(1, m_LEDPictureChar, Mid(m_Text, I, 1)) - 1) * 23 * Screen.TwipsPerPixelY, 13 * Screen.TwipsPerPixelX, 23 * Screen.TwipsPerPixelY
        Next I
        strLastText = m_Text
    Else
        For I = 1 To Len(m_Text)
            If Mid(m_Text, I, 1) <> Mid(strLastText, I, 1) Then UserControl.PaintPicture m_LEDPicture, I * 13 * Screen.TwipsPerPixelX - 13 * Screen.TwipsPerPixelX, 0, 13 * Screen.TwipsPerPixelX, 23 * Screen.TwipsPerPixelY, m_ForeColor * 13 * Screen.TwipsPerPixelX, (InStr(1, m_LEDPictureChar, Mid(m_Text, I, 1)) - 1) * 23 * Screen.TwipsPerPixelY, 13 * Screen.TwipsPerPixelX, 23 * Screen.TwipsPerPixelY
        Next I
    End If
    strLastText = m_Text
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", enumLEDBorderStyle.ledbsNone)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_MaxLength = PropBag.ReadProperty("MaxLength", m_def_MaxLength)
    m_DefaultText = PropBag.ReadProperty("DefaultText", m_def_DefaultText)
    ForcePaint = True
    UserControl_Paint
End Sub

Private Sub UserControl_Resize()
    Static lastWidth As Integer, lastHeight As Integer
    UserControl.Height = UserControl.Height - UserControl.ScaleHeight + Screen.TwipsPerPixelY * 23
    UserControl.Width = UserControl.Width - UserControl.ScaleWidth + Screen.TwipsPerPixelX * 13 * MaxLength
    If lastWidth <> UserControl.Width Or lastHeight <> UserControl.Height Then
        RaiseEvent Resize
    End If
    UserControl_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, enumLEDBorderStyle.ledbsNone)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("MaxLength", m_MaxLength, m_def_MaxLength)
    Call PropBag.WriteProperty("DefaultText", m_DefaultText, m_def_DefaultText)
End Sub

Public Sub SetLED(ByVal bytMaxLength As Byte, ByVal strText As String, Optional ByVal strDefaultText As String = " ")
    ForcePaint = True
    If MaxLength <> bytMaxLength Then MaxLength = bytMaxLength
    ForcePaint = True
    If DefaultText <> strDefaultText Then DefaultText = strDefaultText
    ForcePaint = True
    If Text <> strText Then Text = strText
End Sub

Public Sub SetTime()
    SetLED 9, Format(Now(), "hh:mm:ss")
End Sub

Public Sub SetDate()
    SetLED 9, Format(Now(), "yy-mm-dd")
End Sub

Public Sub SetTimerMinuteSecond(ByVal sngSecond As Single)
    Dim bytSecond As Integer
    Dim bytMinute As Byte
    bytSecond = sngSecond Mod 60
    bytMinute = sngSecond \ 60
    SetLED 9, Format(bytMinute, "00") & ":" & Format(bytSecond, "00")
End Sub

Public Sub SetTimerSecond(ByVal sngSecond As Single)
    SetLED 9, Format(sngSecond), 0
End Sub

Public Sub SetTimerMinuteSecondMilesecond(ByVal sngSecond As Single)
    Dim bytSecond As Integer
    Dim bytMinute As Integer
    Dim bytMilesecond As Integer
    bytMinute = Int(sngSecond \ 60)
    bytSecond = Int(sngSecond Mod 60)
    bytMilesecond = Round(sngSecond - (Int(sngSecond)), 3) * 1000
    SetLED 9, Format(bytMinute, "00") & ":" & Format(bytSecond, "00") & "." & Format(bytMilesecond, "000")
End Sub

Public Sub SetTimerSecondMilesecond(ByVal sngSecond As Single)
    SetLED 9, Format(sngSecond, "#.000"), 0
End Sub

Public Sub Refresh()
    ForcePaint = True
    UserControl_Paint
End Sub
