VERSION 5.00
Begin VB.UserControl ctlLEDBoard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.VScrollBar vsbScroll 
      Height          =   3615
      LargeChange     =   480
      Left            =   3960
      SmallChange     =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin orzMinesweeper.ctlLED led 
      Height          =   345
      Index           =   0
      Left            =   2160
      Top             =   360
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      Text            =   "74520.617"
      MaxLength       =   9
   End
   Begin VB.PictureBox picIcons 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   1935
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "ctlLEDBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Items As Variant
Private ledBoardImages As StdPicture
Private ScrollBarCheck As Boolean
Event Resize()

Private Sub UserControl_Initialize()
    Set ledBoardImages = LoadResPicture(6, vbResBitmap)
    ScrollBarCheck = False
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    Dim LEDWidthMax As Single
    Dim I As Integer
    For I = led.LBound To led.UBound
        If led(I).Width > LEDWidthMax Then LEDWidthMax = led(I).Width
    Next I
    UserControl.Width = ScaleX(ledBoardImages.Width, vbHimetric, vbTwips) + LEDWidthMax + 240 + IIf(vsbScroll.Visible, vsbScroll.Width, 0)
    For I = led.LBound To led.UBound
        led(I).Left = UserControl.ScaleWidth - 120 - led(I).Width - IIf(vsbScroll.Visible, vsbScroll.Width, 0)
        led(I).Top = 120 + (led(I).Height + 180) * I - IIf(vsbScroll.Visible, vsbScroll.value, 0)
    Next I
    
    picIcons.Move 0, -IIf(vsbScroll.Visible, vsbScroll.value, 0), UserControl.ScaleWidth - IIf(vsbScroll.Visible, vsbScroll.Width, 0), 120 + (led(0).Height + 180) * I
    If vsbScroll.Visible Then
        vsbScroll.Move UserControl.ScaleWidth - vsbScroll.Width, 0, vsbScroll.Width, UserControl.ScaleHeight
        If vsbScroll.Max <> picIcons.Height - UserControl.ScaleHeight Then vsbScroll.Max = picIcons.Height - UserControl.ScaleHeight
        If vsbScroll.LargeChange <> UserControl.ScaleHeight Then vsbScroll.LargeChange = UserControl.ScaleHeight
        If vsbScroll.SmallChange <> UserControl.ScaleHeight \ 4 Then vsbScroll.SmallChange = UserControl.ScaleHeight \ 4
    End If
    If ScrollBarCheck = False Then
        If picIcons.Height > UserControl.ScaleHeight Then
            If vsbScroll.Visible = False Then
                vsbScroll.Visible = True
                vsbScroll.value = 0
                ScrollBarCheck = True
                UserControl_Resize
            End If
        Else
            If vsbScroll.Visible = True Then
                vsbScroll.Visible = False
                ScrollBarCheck = True
                UserControl_Resize
            End If
        End If
    Else
        ScrollBarCheck = False
    End If
    RaiseEvent Resize
End Sub

Public Sub SetItems(Items As Variant)
    Dim I As Integer
    UnloadAllLED
    m_Items = Items
    For I = LBound(Items) To UBound(Items)
        LoadLED I
        With led(I) '("timer", 9, "00:00.000", " ", 1)
            .MaxLength = Items(I)(1)
            .DefaultText = Items(I)(3)
            .Text = Items(I)(2)
            If UBound(Items(I)) >= 5 Then
                If Items(I)(5) <> -1 Then
                    .ForeColor = Items(I)(5)
                Else
                    .ForeColor = ledfcRed
                End If
            Else
                .ForeColor = ledfcRed
            End If
            .Refresh
            .ZOrder 0
        End With
    Next
    UserControl_Resize
    picIcons.Cls
    For I = LBound(Items) To UBound(Items)
        If UBound(Items(I)) >= 4 Then
            If Items(I)(4) <> -1 Then
                picIcons.PaintPicture ledBoardImages, 120, 120 + (led(I).Height + 180) * I, ScaleX(ledBoardImages.Width, vbHimetric, vbTwips), led(I).Height, 0, led(I).Height * Items(I)(4), ScaleX(ledBoardImages.Width, vbHimetric, vbTwips), led(I).Height
            End If
            picIcons.ForeColor = &H808080
            picIcons.Line (105, 120 + (led(I).Height + 180) * I - 15)-(picIcons.ScaleWidth - 120, 120 + (led(I).Height + 180) * I - 15), , BF
            picIcons.Line (95, 120 + (led(I).Height + 180) * I - 30)-(picIcons.ScaleWidth - 105, 120 + (led(I).Height + 180) * I - 30), , BF
            picIcons.Line (105, 120 + (led(I).Height + 180) * I)-(105, 120 + (led(I).Height + 180) * I + led(I).Height), , BF
            picIcons.Line (95, 120 + (led(I).Height + 180) * I - 15)-(95, 120 + (led(I).Height + 180) * I + led(I).Height + 15), , BF
            picIcons.ForeColor = vbWhite
            picIcons.Line (95, 120 + (led(I).Height + 180) * I + led(I).Height + 15)-(picIcons.ScaleWidth - 105, 120 + (led(I).Height + 180) * I + led(I).Height + 15), , BF
            picIcons.Line (105, 120 + (led(I).Height + 180) * I + led(I).Height)-(picIcons.ScaleWidth - 120, 120 + (led(I).Height + 180) * I + led(I).Height), , BF
            picIcons.Line (picIcons.ScaleWidth - 120, 120 + (led(I).Height + 180) * I)-(picIcons.ScaleWidth - 120, 120 + (led(I).Height + 180) * I + led(I).Height), , BF
            picIcons.Line (picIcons.ScaleWidth - 105, 120 + (led(I).Height + 180) * I - 15)-(picIcons.ScaleWidth - 105, 120 + (led(I).Height + 180) * I + led(I).Height + 15), , BF
        End If
    Next
    UserControl_Resize
End Sub

Private Sub LoadLED(ByVal Index As Integer)
    On Error Resume Next
    Load led(Index)
    Err.Clear
    led(Index).Visible = True
    On Error GoTo 0
End Sub

Private Sub UnloadAllLED()
    On Error Resume Next
    Dim I As Integer
    For I = led.LBound To led.UBound
        Unload led(I)
        Err.Clear
    Next I
End Sub

Public Property Get LEDs(Optional ByVal Index As Integer = -1, Optional ByVal Key As String = "") As ctlLED
    Dim I As Integer
    If Index = -1 And Key = "" Then Err.Raise 5: Exit Property
    If Index <> -1 And Key <> "" Then Err.Raise 5: Exit Property
    If Index <> -1 Then
        Set LEDs = led(Index)
        Exit Property
    ElseIf Key <> "" Then
        For I = LBound(m_Items) To UBound(m_Items)
            If LCase(m_Items(I)(0)) = LCase(Key) Then
                Set LEDs = led(I)
                Exit Property
            End If
        Next
    End If
End Property

Private Sub vsbScroll_Change()
    ScrollBarCheck = True
    UserControl_Resize
End Sub

Private Sub vsbScroll_Scroll()
    ScrollBarCheck = True
    UserControl_Resize
    DoEvents
End Sub
