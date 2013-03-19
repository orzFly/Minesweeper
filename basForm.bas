Attribute VB_Name = "basForm"
'Project name: orzMinesweeper
'Code license: GNU General Public License v3
'Author      : Yeechan Lu a.k.a. orzFly <i@orzfly.com>

Public Function GetFormPaddingSize(ByVal hwnd As Long) As RECT
    Dim OuterBorder As RECT
    Dim InnerBorder As RECT
    basWin32API.GetWindowRect hwnd, OuterBorder
    basWin32API.GetClientRect hwnd, InnerBorder
    With GetFormPaddingSize
        .Left = 0
        .Top = 0
        .Right = OuterBorder.Right - OuterBorder.Left - (InnerBorder.Right - InnerBorder.Left)
        .Bottom = OuterBorder.Bottom - OuterBorder.Top - (InnerBorder.Bottom - InnerBorder.Top)
    End With
End Function

Public Function RectPixelToTwip(lpRect As RECT) As RECT
    With RectPixelToTwip
        .Left = lpRect.Left * Screen.TwipsPerPixelX
        .Top = lpRect.Top * Screen.TwipsPerPixelY
        .Right = lpRect.Right * Screen.TwipsPerPixelX
        .Bottom = lpRect.Bottom * Screen.TwipsPerPixelY
    End With
End Function

Public Function SetClientRect(objForm As Form, lngWidth As Long, lngHeight As Long)
    Dim PaddingSize As RECT
    With objForm
        PaddingSize = RectPixelToTwip(GetFormPaddingSize(.hwnd))
        .Width = lngWidth + PaddingSize.Right - PaddingSize.Left
        .Height = lngHeight + PaddingSize.Bottom - PaddingSize.Top
    End With
End Function

Public Function SetFormCentered(objForm As Form)
    With objForm
        .Left = Screen.Width / 2 - .Width / 2
        .Top = Screen.Height / 2 - .Height / 2
    End With
End Function

Public Function MoveLine(lineObject As Line, ByVal sngX1 As Single, ByVal sngY1 As Single, ByVal sngX2 As Single, ByVal sngY2 As Single)
    With lineObject
        .X1 = sngX1
        .Y1 = sngY1
        .X2 = sngX2
        .Y2 = sngY2
    End With
End Function
