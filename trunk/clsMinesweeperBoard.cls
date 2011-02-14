VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMinesweeperBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project name: orzMinesweeper
'Code license: GNU General Public License v3
'Author      : Yeechan Lu a.k.a. orzFly <i@orzfly.com>

Option Explicit

Dim m_Board() As Integer
Dim m_Width As Integer
Dim m_Height As Integer
Dim m_Mines As Integer
Dim m_MaxMinesPerCell As Integer
Dim m_GameID As Long
Dim m_BoardType As Integer
Dim m_MentionedCells As String

Public Property Get GameID() As Long
    GameID = m_GameID
End Property

Public Property Get Board2D(ByVal X As Integer, ByVal Y As Integer, Optional ByVal Error As Boolean = False) As Integer
    Dim Pos As Integer
    Pos = m_Width * Y + X
    If Pos >= LBound(m_Board) And Pos <= UBound(m_Board) And X <= m_Width - 1 And X >= 0 And Y <= m_Height - 1 And Y >= 0 Then
        Board2D = m_Board(Pos)
    ElseIf Error Then
        Board2D = -1
    End If
End Property

Public Property Get Board1D(ByVal Pos As Integer) As Integer
    If Pos >= LBound(m_Board) And Pos <= UBound(m_Board) Then
        Board1D = m_Board(Pos)
    End If
End Property

Public Property Let Board1D(ByVal Pos As Integer, ByVal value As Integer)
    If Pos >= LBound(m_Board) And Pos <= UBound(m_Board) Then
        m_Board(Pos) = value
    End If
End Property

Public Property Let Board2D(ByVal X As Integer, ByVal Y As Integer, Optional ByVal Error As Boolean = False, ByVal value As Integer)
    Dim Pos As Integer
    Pos = m_Width * Y + X
    If Pos >= LBound(m_Board) And Pos <= UBound(m_Board) And X <= m_Width - 1 And X >= 0 And Y <= m_Height - 1 And Y >= 0 Then
        m_Board(Pos) = value
    End If
End Property

Public Property Get Width() As Integer
    Width = m_Width
End Property

Public Property Get Height() As Integer
    Height = m_Height
End Property

Public Property Get Mines() As Integer
    Mines = m_Mines
End Property

Public Property Get BoardType() As Integer
    BoardType = m_BoardType
End Property

Public Property Get MaxMinesPerCell() As Integer
    MaxMinesPerCell = m_MaxMinesPerCell
End Property

Public Property Let MaxMinesPerCell(value As Integer)
    m_MaxMinesPerCell = value
End Property

Public Sub Initialize(ByVal intWidth As Integer, ByVal intHeight As Integer, ByVal intMines As Integer, ByVal intBoardType As Integer, Optional ByVal intGameID As Long = -1)
    Dim I As Integer, J As Integer, Pos As Integer, Num As Integer

    m_Width = intWidth
    m_Height = intHeight
    m_Mines = intMines
    m_BoardType = intBoardType

    If intGameID = -1 Then
        Randomize
        intGameID = Int(Rnd() * 499999999 + 1)
    End If

    Do While m_Width < BoardWidthMin Or m_Width > BoardWidthMax
        m_Width = Int(BoardWidthMax * MyRnd(intGameID))
    Loop

    Do While m_Height < BoardHeightMin Or m_Height > BoardHeightMax
        m_Height = Int(BoardHeightMax * MyRnd(intGameID))
    Loop

    Do While m_MaxMinesPerCell < BoardMaxMinesPerCellMin Or m_MaxMinesPerCell > BoardMaxMinesPerCellMax
        m_MaxMinesPerCell = Int(10 * MyRnd(intGameID))
    Loop

    Do While m_Mines > BoardMinesMax(m_Width, m_Height, m_MaxMinesPerCell)
        m_Mines = m_Mines - m_Height
    Loop

    Do While m_Mines < BoardMinesMin
        m_Mines = Int(BoardMinesMax(m_Width, m_Height, m_MaxMinesPerCell) * MyRnd(intGameID))
    Loop


    '初始化随机数
    m_GameID = intGameID

    '初始化棋盘数组
    ReDim m_Board(m_Width * m_Height - 1)
    Dim AvailableCells As New Collection
    For I = LBound(m_Board) To UBound(m_Board)
        m_Board(I) = 0
        AvailableCells.Add I, Format(I + 1)
    Next I
    '放炸弹
    I = 0
    Do
        Pos = AvailableCells(Int(MyRnd(intGameID) * AvailableCells.Count + 1))
        '没放过炸弹的格子，里面值为周围炸弹数（这个可以抹掉，毕竟有炸弹的格子只是炸弹不显示周围八格的炸弹数。）
        '随机产生炸弹数，小于剩余炸弹数，也小于每格的限制
        Num = Min(m_Mines - I, m_MaxMinesPerCell, m_MaxMinesPerCell - IIf(Board1D(Pos) < 0, Abs(Board1D(Pos)), 0))
        If Num > 0 Then Num = Int(MyRnd(intGameID) * Num + 1)
        '棋盘里负数的格子代表有炸弹格，其绝对值代表炸弹数。
        If Board1D(Pos) <= 0 Then
            Board1D(Pos) = Board1D(Pos) - Num
        Else
            Board1D(Pos) = -Num
        End If
        If Board1D(Pos) = -m_MaxMinesPerCell Then
            AvailableCells.Remove (Format(Pos + 1))
        End If
        I = I + Num
    Loop Until I = m_Mines

    For I = LBound(m_Board) To UBound(m_Board)
        If Board1D(I) >= 0 Then
            Dim MinesCount As Integer
            MinesCount = 0
            MentionedCellsInit I Mod Width, I \ Width
            '根据规则填充周围八格的数据，遇到放了炸弹的就略过。
            For J = 1 To MentionedCellsCount()
                Dim X2 As Integer, Y2 As Integer
                X2 = I Mod Width + MentionedCellsDeltaX(J - 1)
                Y2 = I \ Width + MentionedCellsDeltaY(J - 1)

                If Board2D(X2, Y2) <= 0 Then
                    MinesCount = MinesCount + Abs(Board2D(X2, Y2))
                End If
            Next J
            Board1D(I) = MinesCount
        End If
    Next I
End Sub


Private Function MyRnd(intGameID As Long) As Double
    Static LastNumber As Double
    Static MyGameID As Long
    If MyGameID = 0 Then MyGameID = Val("&H" & Right(MD5(intGameID), 7)) + 1
again:
    If LastNumber = 0 Then LastNumber = 0.314159265358979 + Val("&H" & Right(MD5(intGameID), 7)) / 3333333333# + 1
    LastNumber = LastNumber - Int(LastNumber)
    MyRnd = (((MyGameID * MyRndInner(MyGameID * (LastNumber + 0.3141592653589)) * (LastNumber + 0.958924274663138) Mod 32719) / 32719) + 1) + Sin(MyGameID) + Sin(LastNumber)
    MyRnd = MyRnd - Int(MyRnd)
    LastNumber = MyRnd
    If LastNumber = 0 Then GoTo again
End Function

Private Function MyRndInner(intGameID As Long) As Double
    Static LastNumber As Double
    If LastNumber = 0 Then LastNumber = 0.628318530717958
    MyRndInner = Sqr(Abs(((LastNumber * intGameID * 11035 + 12345) / 65536) Mod 32768))
    MyRndInner = MyRndInner - Int(MyRndInner)
    LastNumber = MyRndInner
End Function

Public Function Pos2Dto1D(ByVal X As Integer, ByVal Y As Integer) As Integer
    Pos2Dto1D = m_Width * Y + X
    If Not (Pos2Dto1D >= LBound(m_Board) And Pos2Dto1D <= UBound(m_Board) And X <= m_Width - 1 And X >= 0 And Y <= m_Height - 1 And Y >= 0) Then Pos2Dto1D = -1
End Function

Public Property Get MentionedCellsRaw(ByVal intBoardType As Integer, ByVal intX As Integer, ByVal intY As Integer) As String
    '棋盘类型影响单元格
    '    长度偶数。两个字符一组。XY。
    '    X和Y=123456789ABCDEF
    '    0  0 8  0
    '    1 +1 F -1
    '    2 +2 E -2
    '    3 +3 D -3
    '    4 +4 C -4
    '    5 +5 B -5
    '    6 +6 A -6
    '    7 +7 9 -7
    Select Case intBoardType
        Case 1
            MentionedCellsRaw = IIf((intY) Mod 2 = 1, "0F1FF0100111", "FF0FF010F101")
        Case 2
            MentionedCellsRaw = IIf((intY) Mod 2 = 1, "F0100111020F1F0E", "F010FF0F0EF10102")
        Case Else
            MentionedCellsRaw = "FF0F1FF010F10111"    '矩形周围两层 "EEFE0E1E2EEFFF0F1F2FE0F01020E1F1011121E2F2021222"
    End Select
End Property

Public Function MentionedCellsInit(ByVal intX As Integer, ByVal intY As Integer)
    m_MentionedCells = MentionedCellsRaw(BoardType, intX, intY)
End Function

Public Property Get MentionedCellsCount() As Byte
    MentionedCellsCount = Len(m_MentionedCells) / 2
End Property

Public Property Get MentionedCellsDeltaX(ByVal intIndex As Byte) As Integer
    MentionedCellsDeltaX = Choose(Hex2Dec(Mid(m_MentionedCells, intIndex * 2 + 1, 1)) + 1, 0, 1, 2, 3, 4, 5, 6, 7, 0, -7, -6, -5, -4, -3, -2, -1)
End Property

Public Property Get MentionedCellsDeltaY(ByVal intIndex As Byte) As Integer
    MentionedCellsDeltaY = Choose(Hex2Dec(Mid(m_MentionedCells, intIndex * 2 + 2, 1)) + 1, 0, 1, 2, 3, 4, 5, 6, 7, 0, -7, -6, -5, -4, -3, -2, -1)
End Property
