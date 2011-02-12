VERSION 5.00
Begin VB.Form frmGame 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmGame"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10545
   Icon            =   "frmGame.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   StartUpPosition =   1  '所有者中心
   Begin orzMinesweeper.ctlLEDBoard ledBoard 
      Height          =   615
      Left            =   6960
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
      _ExtentX        =   5503
      _ExtentY        =   1085
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer tmrScreenSaverWaiter 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer tmrScreenSaverAutoplayer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picBoard 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5400
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Shape shapeBackground 
      BackColor       =   &H80000000&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   7800
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Menu mnuGame 
      Caption         =   "mnuGame"
      HelpContextID   =   100
      Begin VB.Menu mnuGameNewGame 
         Caption         =   "mnuGameNewGame"
         HelpContextID   =   101
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameSelectGame 
         Caption         =   "mnuGameSelectGame"
         HelpContextID   =   102
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameDummyA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameStatistics 
         Caption         =   "mnuGameStatistics"
         HelpContextID   =   103
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "mnuOptions"
      HelpContextID   =   200
      Begin VB.Menu mnuOptionsDifficulty 
         Caption         =   "mnuOptionsDifficulty"
         HelpContextID   =   201
         Begin VB.Menu mnuOptionsDifficultyItems 
            Caption         =   "mnuOptionsDifficultyItems(0)"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOptionsDifficultyDummyA 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsDifficultyCustom 
            Caption         =   "mnuOptionsDifficultyCustom"
            HelpContextID   =   500
         End
      End
      Begin VB.Menu mnuOptionsBoardType 
         Caption         =   "mnuOptionsBoardType"
         HelpContextID   =   202
         Begin VB.Menu mnuOptionsBoardTypeNormal 
            Caption         =   "mnuOptionsBoardTypeNormal"
            HelpContextID   =   601
         End
         Begin VB.Menu mnuOptionsBoardTypeHexagon 
            Caption         =   "mnuOptionsBoardTypeHexagon"
            HelpContextID   =   602
         End
         Begin VB.Menu mnuOptionsBoardTypeDiamond 
            Caption         =   "mnuOptionsBoardTypeDiamond"
            HelpContextID   =   603
         End
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuActions 
         Caption         =   "mnuActions"
         HelpContextID   =   400
         Begin VB.Menu mnuActionsNone 
            Caption         =   "mnuActionsNone"
            HelpContextID   =   401
         End
         Begin VB.Menu mnuActionsFlags 
            Caption         =   "mnuActionsFlags(0)"
            Index           =   0
         End
         Begin VB.Menu mnuActionsQuestions 
            Caption         =   "mnuActionsQuestions(0)"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Board As clsMinesweeperBoard
Dim MouseBoardImage As StdPicture
Dim BoardImage As StdPicture
Dim BoardImageWidth As Integer
Dim BoardImageHeight As Integer
Dim SmallDigitalImage As StdPicture
Dim m_BoardData() As Integer
Dim MouseX As Integer
Dim MouseY As Integer
Dim LastPushedX As Integer
Dim LastPushedY As Integer
Dim PopupMenuX As Integer
Dim PopupMenuY As Integer
Dim MouseLeftPushed As Boolean
Dim MouseRightPushed As Boolean
Dim MouseLeftReleased As Boolean
Dim MouseRightReleased As Boolean
Dim GameMaxMinesPerCell As Integer
Dim GameBoardWidth As Integer
Dim GameBoardHeight As Integer
Dim GameBoardMines As Integer
Dim GameBoardType As Integer
Dim GameBoardMineCells As Integer
Dim GameBoard3BV As Integer
Dim UnopenedCells As Collection
Dim BlankCells As Collection
Dim BlankCellsGroups As Collection
Dim mHandle As Long, sHandle As Long, dHandle As Long
Attribute sHandle.VB_VarUserMemId = 1073938459
Attribute dHandle.VB_VarUserMemId = 1073938459
Dim GameStartTime As Single
Attribute GameStartTime.VB_VarUserMemId = 1073938462
Dim GameEndTime As Single
Attribute GameEndTime.VB_VarUserMemId = 1073938463
Dim FirstClick As Boolean
Attribute FirstClick.VB_VarUserMemId = 1073938464
Dim FlagsCount() As Integer
Dim MinesCount As Integer
Dim FlaggedFlagsCount() As Integer
Dim FlaggedMinesCount As Integer
Dim ProcessingDblClick As Boolean

Private Sub NewGame(Optional ByVal GameID As Long = -1)
    Dim I As Integer, J As Integer, K As Integer, Flag As Boolean

    Me.picBoard.Enabled = True And (Not ScreenSaverMode)
    Me.picBoard.Cls

    Set Board = New clsMinesweeperBoard
    If ScreenSaverMode Then
        Randomize
        LoadBoard Int(Rnd() * 3)
        GameMaxMinesPerCell = 0
        GameBoardWidth = 0
        GameBoardHeight = 0
        GameBoardMines = 0
    End If
    Board.MaxMinesPerCell = GameMaxMinesPerCell

    Board.Initialize GameBoardWidth, GameBoardHeight, GameBoardMines, GameBoardType, GameID
    GameBoardWidth = Board.Width
    GameBoardHeight = Board.Height
    GameBoardMines = Board.Mines
    GameBoardMineCells = 0

    GameMaxMinesPerCell = Board.MaxMinesPerCell
    ReDim m_BoardData(Board.Width * Board.Height - 1)
    Set UnopenedCells = New Collection
    Set BlankCells = New Collection
    Set BlankCellsGroups = New Collection
    ReDim FlagsCount(Board.MaxMinesPerCell - 1)
    ReDim FlaggedFlagsCount(Board.MaxMinesPerCell - 1)
    
    Me.picBoard.Width = CalcBoardWidth
    Me.picBoard.Height = CalcBoardHeight
    Form_Resize

    Dim X As Integer
    Dim Y As Integer
    For Y = 0 To Board.Height - 1
        For X = 0 To Board.Width - 1
            PaintBoard X, Y, 0
            UnopenedCells.Add Format(X) & "#" & Format(Y), Format(X) & "#" & Format(Y)
            If Board.Board2D(X, Y) = 0 Then BlankCells.Add Format(X) & "#" & Format(Y), Format(X) & "#" & Format(Y)
            If Board.Board2D(X, Y) < 0 Then GameBoardMineCells = GameBoardMineCells + 1
            If Board.Board2D(X, Y) < 0 Then
                FlagsCount(Abs(Board.Board2D(X, Y)) - 1) = FlagsCount(Abs(Board.Board2D(X, Y)) - 1) + 1
            End If
        Next X
    Next Y

    'For Y = 0 To Board.Height - 1
    'For X = 0 To Board.Width - 1
    'If Board.Board2D(X, Y) >= 0 Then OpenBoard X, Y
    'Next X
    'Next Y

    If BlankCells.Count > 0 Then
        Dim Z As Integer
        For Z = 0 To 1
            For I = 0 To BlankCells.Count - 1
                If BlankCellsGroups.Count = 0 Then
                    BlankCellsGroups.Add "!" & BlankCells(I + 1) & "!"
                Else
                    Dim point
                    point = Split(BlankCells(I + 1), "#")
                    Flag = False
                    For J = 0 To BlankCellsGroups.Count - 1
                        For K = 0 To 3
                            Dim strPoint As String
                            strPoint = "!" & Format(Val(point(0)) + Choose(K + 1, -1, 1, 0, 0)) & "#" & Format(Val(point(1)) + Choose(K + 1, 0, 0, -1, 1)) & "!"
                            If InStr(1, BlankCellsGroups(J + 1), strPoint) > 0 Then
                                strPoint = CStr(BlankCellsGroups(J + 1) & "!" & BlankCells(I + 1) & "!") & strPoint
                                BlankCellsGroups.Remove J + 1
                                BlankCellsGroups.Add strPoint
                                Flag = True
                            End If
                        Next K
                    Next J
                    If Flag = False Then
                        BlankCellsGroups.Add "!" & BlankCells(I + 1) & "!"
                    End If
                End If
            Next I
            Dim colTemp As Collection
            Dim strTemp As String
            Dim groupMembers As Variant, colDup As Collection
            Set colTemp = New Collection
            For I = 0 To BlankCellsGroups.Count - 1
                groupMembers = Split(Replace(Replace(BlankCellsGroups(I + 1), "!!", "orz"), "!", ""), "orz")
                If UBound(groupMembers) > 0 Then
                    Set colDup = New Collection
                    On Error Resume Next
                    For J = 0 To UBound(groupMembers)
                        colDup.Add groupMembers(J), groupMembers(J)
                        Err.Clear
                    Next J
                    On Error GoTo 0
                    strTemp = ""
                    For J = 0 To colDup.Count - 1
                        strTemp = strTemp & "!" & colDup.Item(J + 1) & "!"
                    Next J
                    colTemp.Add strTemp
                Else
                    colTemp.Add BlankCellsGroups(I + 1)
                End If
            Next I
            Set BlankCellsGroups = colTemp
            If BlankCellsGroups.Count > 1 Then
                Dim CombinableBlankCellsGroups As New ctlIntegerStack
                For I = 0 To BlankCells.Count - 1
                    CombinableBlankCellsGroups.Clear
                    For J = 0 To BlankCellsGroups.Count - 1
                        If InStr(1, BlankCellsGroups(J + 1), "!" & BlankCells(I + 1) & "!") > 0 Then
                            CombinableBlankCellsGroups.Push J + 1
                        End If
                    Next J
                    If CombinableBlankCellsGroups.Count > 1 Then
                        strTemp = ""
                        For J = 1 To CombinableBlankCellsGroups.Count
                            strTemp = strTemp & BlankCellsGroups(CombinableBlankCellsGroups.Peek())
                            BlankCellsGroups.Remove CombinableBlankCellsGroups.Pop()
                        Next J
                        strTemp = Replace(strTemp, "!" & BlankCells(I + 1) & "!", "") & "!" & BlankCells(I + 1) & "!"
                        BlankCellsGroups.Add strTemp
                    End If
                Next
            End If
            Set colTemp = New Collection
            For I = 0 To BlankCellsGroups.Count - 1
                groupMembers = Split(Replace(Replace(BlankCellsGroups(I + 1), "!!", "orz"), "!", ""), "orz")
                If UBound(groupMembers) > 0 Then
                    Set colDup = New Collection
                    On Error Resume Next
                    For J = 0 To UBound(groupMembers)
                        colDup.Add groupMembers(J), groupMembers(J)
                        Err.Clear
                    Next J
                    On Error GoTo 0
                    strTemp = ""
                    For J = 0 To colDup.Count - 1
                        strTemp = strTemp & "!" & colDup.Item(J + 1) & "!"
                    Next J
                    colTemp.Add strTemp
                Else
                    colTemp.Add BlankCellsGroups(I + 1)
                End If
            Next I
            Set BlankCellsGroups = colTemp
        Next
    End If

    'If GameBoardType = 0 Then
        Flag = False
        GameBoard3BV = BlankCellsGroups.Count
        For Y = 0 To Board.Height - 1
            For X = 0 To Board.Width - 1
                If Board.Board2D(X, Y) > 0 Then
                    Flag = False
                    Board.MentionedCellsInit X, Y
                    For I = 0 To Board.MentionedCellsCount - 1
                        If Board.Board2D(X + Board.MentionedCellsDeltaX(I), Y + Board.MentionedCellsDeltaY(I), True) = 0 Then
                            Flag = True
                            Exit For
                        End If
                    Next
                    If Flag = False Then
                        GameBoard3BV = GameBoard3BV + 1
                    End If
                End If
            Next X
        Next Y
        If GameBoard3BV > BlankCellsGroups.Count Then GameBoard3BV = GameBoard3BV - 1
    'Else
    '    GameBoard3BV = -1
    'End If

    If Not ScreenSaverMode Then
        Dim img As Long, Count As Integer
        If Me.mnuActionsFlags.UBound > Me.mnuActionsFlags.LBound Then
            For I = Me.mnuActionsFlags.LBound + 1 To Me.mnuActionsFlags.UBound
                Unload Me.mnuActionsFlags(I)
            Next I
        End If
        If Me.mnuActionsQuestions.UBound > Me.mnuActionsQuestions.LBound Then
            For I = Me.mnuActionsQuestions.LBound + 1 To Me.mnuActionsQuestions.UBound
                Unload Me.mnuActionsQuestions(I)
            Next I
        End If
        For I = 1 To GameMaxMinesPerCell
            Load Me.mnuActionsFlags(I)
            With Me.mnuActionsFlags(I)
                If I = 1 Then .Caption = LoadResString(402) Else .Caption = Replace(LoadResString(403), "%s", Format(I))
                .Visible = True
            End With
            Load Me.mnuActionsQuestions(I)
            With Me.mnuActionsQuestions(I)
                If I = 1 Then .Caption = LoadResString(404) Else .Caption = Replace(LoadResString(405), "%s", Format(I))
                .Visible = True
            End With
        Next I

        InsertMenu mHandle, 2, MF_BYPOSITION Or MF_POPUP Or MF_STRING, sHandle, ""
        img = GetBitMapHandle(0): SetMenuItemBitmaps dHandle, 0, MF_BYPOSITION, img, img
        Count = 1
        For I = 1 To GameMaxMinesPerCell
            img = GetBitMapHandle(I): SetMenuItemBitmaps dHandle, Count, MF_BYPOSITION, img, img
            Count = Count + 1
        Next I
        For I = 1 To GameMaxMinesPerCell
            img = GetBitMapHandle(-I): SetMenuItemBitmaps dHandle, Count, MF_BYPOSITION, img, img
            Count = Count + 1
        Next I
        RemoveMenu mHandle, 2, MF_BYPOSITION
        DrawMenuBar Me.hWnd
    End If

    FirstClick = True
    MinesCount = Board.Mines
    FlaggedMinesCount = 0

    tmrGame.Enabled = False
    GameStartTime = -1
    GameEndTime = -1
    
    tmrClock.Enabled = True
    Dim ledItems As New Collection
    ledItems.Add Array("timer", 9, "00:00.000", " ", 0, enumLEDForeColor.ledfcRed)
    ledItems.Add Array("mines", 9, Format(MinesCount), "0", 3, enumLEDForeColor.ledfcRed)
    For I = 1 To Board.MaxMinesPerCell
        ledItems.Add Array("flags" & Format(I), 9, Format(FlagsCount(I - 1)), "0", 3 + I, enumLEDForeColor.ledfcRed)
    Next I
    ledItems.Add Array("time", 9, "", "", 2, enumLEDForeColor.ledfcGreen)
    ledItems.Add Array("date", 9, "", "", 1, enumLEDForeColor.ledfcGreen)
    ledItems.Add Array("3BV", 9, IIf(GameBoard3BV = -1, "---------", Format(GameBoard3BV)), "0", 15, enumLEDForeColor.ledfcYellow)
    ledItems.Add Array("gameID", 9, Format(Board.GameID), "0", 14, enumLEDForeColor.ledfcYellow)
    Dim arrayLedItems
    ReDim arrayLedItems(ledItems.Count - 1)
    For I = LBound(arrayLedItems) To UBound(arrayLedItems)
        arrayLedItems(I) = ledItems(I + 1)
    Next I
    Set ledItems = Nothing
    ledBoard.SetItems arrayLedItems
    
    tmrClock_Timer
    tmrGame_Timer
    DoEvents
End Sub

Private Function GetBitMapHandle(ByVal Num As Integer)
    If ScreenSaverMode Then GetBitMapHandle = 0: Exit Function
    Dim dstWidth As Long, dstHeight As Long
    Dim hDc4 As Long, hDc5 As Long, I As Long
    Dim hBitmap As Long
    Dim hDstDc As Long

    hDc4 = CreateCompatibleDC(0)
    hDc5 = CreateCompatibleDC(0)
    Call SelectObject(hDc4, SmallDigitalImage.Handle)
    Call SelectObject(hDc5, MouseBoardImage.Handle)

    I = GetMenuCheckMarkDimensions
    dstWidth = I Mod 2 ^ 16
    dstHeight = I / 2 ^ 16
    hBitmap = CreateCompatibleBitmap(Me.hdc, dstWidth, dstHeight)
    hDstDc = CreateCompatibleDC(Me.hdc)
    SelectObject hDstDc, hBitmap

    If Num = 0 Then
        Call StretchBlt(hDstDc, 0, 0, dstWidth, dstHeight, hDc5, 0, 0, 16, 16, SRCCOPY)
    ElseIf Num > 0 Then
        If Num > 0 And Num <= 4 Then
            Call StretchBlt(hDstDc, 0, 0, dstWidth, dstHeight, hDc5, Num * 16 - 16, 1 * 16, 16, 16, SRCCOPY)
        Else
            Call StretchBlt(hDstDc, 0, 0, dstWidth, dstHeight, hDc5, 64, 1 * 16, 16, 16, SRCCOPY)
            Call StretchBlt(hDstDc, 9 / 16 * dstWidth, 2 / 16 * dstHeight, 6 / 16 * dstWidth, 13 / 16 * dstHeight, hDc4, 30, 13 * (1 + (10 - Num)), 6, 13, vbMergePaint)
            Call StretchBlt(hDstDc, 9 / 16 * dstWidth, 2 / 16 * dstHeight, 6 / 16 * dstWidth, 13 / 16 * dstHeight, hDc4, 30, 13 * (1 + (10 - Num)), 6, 13, vbSrcAnd)
        End If
    ElseIf Num < 0 Then
        Num = Abs(Num)
        If Num > 0 And Num <= 4 Then
            Call StretchBlt(hDstDc, 0, 0, dstWidth, dstHeight, hDc5, Num * 16 - 16, 2 * 16, 16, 16, SRCCOPY)
        Else
            Call StretchBlt(hDstDc, 0, 0, dstWidth, dstHeight, hDc5, 64, 2 * 16, 16, 16, SRCCOPY)
            Call StretchBlt(hDstDc, 9 / 16 * dstWidth, 2 / 16 * dstHeight, 6 / 16 * dstWidth, 13 / 16 * dstHeight, hDc4, 30, 13 * (1 + (10 - Num)), 6, 13, vbMergePaint)
            Call StretchBlt(hDstDc, 9 / 16 * dstWidth, 2 / 16 * dstHeight, 6 / 16 * dstWidth, 13 / 16 * dstHeight, hDc4, 30, 13 * (1 + (10 - Num)), 6, 13, vbSrcAnd)
        End If
    End If
    GetBitMapHandle = hBitmap
    Call DeleteDC(hDc4)
    Call DeleteDC(hDc5)
    Call DeleteDC(hDstDc)
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim obj As Object, mnu As Menu
    Set MouseBoardImage = LoadResPicture(3, vbResBitmap)
    Set SmallDigitalImage = LoadResPicture(2, vbResBitmap)
    If ScreenSaverMode Then
        For Each obj In Me.Controls
            If TypeOf obj Is Menu Then
                On Error Resume Next
                Set mnu = obj
                mnu.Visible = False
                Err.Clear
                On Error GoTo 0
            End If
        Next

        Dim lStyle As Long
        Dim tR As RECT
        GetWindowRect Me.hWnd, tR
        lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
        Me.Tag = Me.Caption
        Me.Caption = " "
        lStyle = lStyle And Not WS_SYSMENU
        lStyle = lStyle And Not WS_MAXIMIZEBOX
        lStyle = lStyle And Not WS_MINIMIZEBOX
        lStyle = lStyle And Not WS_CAPTION
        SetWindowLong Me.hWnd, GWL_STYLE, lStyle
        SetWindowPos Me.hWnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED

        Me.WindowState = 2

        Randomize

        Show
        NewGame

        Me.tmrScreenSaverAutoplayer.Enabled = True
    Else
        Me.mnuActionsFlags(0).Visible = False
        Me.mnuActionsQuestions(0).Visible = False
        mHandle = GetMenu(hWnd)
        sHandle = GetSubMenu(mHandle, 2)
        dHandle = GetSubMenu(sHandle, 0)
        RemoveMenu mHandle, 2, MF_BYPOSITION

        For Each obj In Me.Controls
            If TypeOf obj Is Menu Then
                Set mnu = obj
                If mnu.HelpContextID <> 0 Then
                    On Error Resume Next
                    mnu.Caption = LoadResString(mnu.HelpContextID)
                    Err.Clear
                    On Error GoTo 0
                    mnu.HelpContextID = 0
                End If
            End If
        Next

        AddDifficultyItem 501, 5, 5, 3, 1
        AddDifficultyItem 502, 7, 7, 5, 1
        AddDifficultyItem 503, 8, 8, 10, 1
        AddDifficultyItem 504, 8, 8, 20, 2
        AddDifficultyItem 505, 16, 16, 40, 1
        AddDifficultyItem 506, 16, 16, 80, 2
        AddDifficultyItem 507, 30, 16, 100, 1
        AddDifficultyItem 508, 30, 16, 200, 2
        AddDifficultyItem 509, 30, 16, 300, 3
        AddDifficultyItem 510, 30, 16, 400, 4

        Me.Caption = LoadResString(1)

        GameMaxMinesPerCell = 1
        GameBoardWidth = 8
        GameBoardHeight = 8
        GameBoardMines = 10
        LoadBoard 0

        Show
        NewGame
    End If
End Sub

Private Sub ledBoard_Resize()
    Form_Resize
End Sub

Private Sub mnuActionsFlags_Click(Index As Integer)
    CheckMinesCounter BoardData(Board.Pos2Dto1D(PopupMenuX, PopupMenuY)), -Index
    BoardData(Board.Pos2Dto1D(PopupMenuX, PopupMenuY)) = -Index
    PaintBoard PopupMenuX, PopupMenuY, 1, Index
End Sub

Private Sub mnuActionsNone_Click()
    CheckMinesCounter BoardData(Board.Pos2Dto1D(PopupMenuX, PopupMenuY)), 0
    BoardData(Board.Pos2Dto1D(PopupMenuX, PopupMenuY)) = 0
    PaintBoard PopupMenuX, PopupMenuY, 0, 0
End Sub

Private Sub mnuActionsQuestions_Click(Index As Integer)
    CheckMinesCounter BoardData(Board.Pos2Dto1D(PopupMenuX, PopupMenuY)), -500 - Index
    BoardData(Board.Pos2Dto1D(PopupMenuX, PopupMenuY)) = -500 - Index
    PaintBoard PopupMenuX, PopupMenuY, 2, Index
End Sub

Private Sub mnuGameNewGame_Click()
    NewGame -1
End Sub

Private Sub mnuGameSelectGame_Click()
    frmSelectGame.Show 1, Me
    If basMain.SelectGameResult > 0 Then NewGame basMain.SelectGameResult
End Sub

Private Sub mnuOptionsBoardTypeDiamond_Click()
    LoadBoard 2
    NewGame
End Sub

Private Sub mnuOptionsBoardTypeHexagon_Click()
    LoadBoard 1
    NewGame
End Sub

Private Sub mnuOptionsBoardTypeNormal_Click()
    LoadBoard 0
    NewGame
End Sub

Private Sub mnuOptionsDifficultyCustom_Click()
    frmCustomGame.Show 1, Me
    If basMain.CustomGameResult <> "" Then
        Dim config As Variant
        config = Split(basMain.CustomGameResult, ",")
        GameBoardWidth = Val(config(0))
        GameBoardHeight = Val(config(1))
        GameBoardMines = Val(config(2))
        GameMaxMinesPerCell = Val(config(3))
        NewGame
    End If
End Sub

Private Sub mnuOptionsDifficultyItems_Click(Index As Integer)
    Dim config As Variant
    config = Split(mnuOptionsDifficultyItems(Index).Tag, ",")
    GameBoardWidth = Val(config(0))
    GameBoardHeight = Val(config(1))
    GameBoardMines = Val(config(2))
    GameMaxMinesPerCell = Val(config(3))
    NewGame
End Sub

Private Sub picBoard_DblClick()
    ProcessingDblClick = True
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBoard_MouseMove Button, Shift, X, Y
    If (Button = 1) Then MouseLeftPushed = True
    If (Button = 2) And MouseLeftPushed Then
        MouseRightPushed = True
    ElseIf Button = 2 Then
        Dim MousePos As Variant
        MousePos = CalcBoardMouseXY(X, Y)
        PopupMenuX = MousePos(0)
        PopupMenuY = MousePos(1)
        Dim offset As Integer, data As Integer, h As Long, temp As Long
        data = BoardData(Board.Pos2Dto1D(PopupMenuX, PopupMenuY))
        If data = 1 Then Exit Sub
        temp = GetMenuCheckMarkDimensions
        h = temp / 2 ^ 16
        Dim rc As RECT
        GetMenuItemRect Me.hWnd, dHandle, 0, rc
        h = rc.Bottom - rc.Top
        Dim def As Menu
        Select Case data
            Case 0
                offset = h
                Set def = Me.mnuActionsFlags(1)
            Case Is < -500
                If Abs(data + 500) = GameMaxMinesPerCell Then
                    offset = 0
                    Set def = Me.mnuActionsNone
                Else
                    offset = h * (GameMaxMinesPerCell + Abs(data + 500)) + h
                    Set def = Me.mnuActionsQuestions(Abs(data + 500) + 1)
                End If
            Case Is < 0
                offset = h * Abs(data) + h
                If Abs(data) = Me.mnuActionsFlags.UBound Then
                    Set def = Me.mnuActionsQuestions(1)
                Else
                    Set def = Me.mnuActionsFlags(Abs(data) + 1)
                End If
        End Select
        PopupMenu Me.mnuActions, 2 + 4, Me.picBoard.Left + X, Me.picBoard.Top + Y - h / 2 - offset, def
    End If
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MousePos As Variant
    MousePos = CalcBoardMouseXY(X, Y)
    MouseX = MousePos(0)
    MouseY = MousePos(1)
    If (basWin32API.GetKeyState(basWin32API.VK_LBUTTON) And &HF0000000) Or Button = 4 Then
        Dim I As Integer
        Board.MentionedCellsInit LastPushedX, LastPushedY
        For I = 1 To Board.MentionedCellsCount
            PopBoard LastPushedX + Board.MentionedCellsDeltaX(I - 1), LastPushedY + Board.MentionedCellsDeltaY(I - 1)
        Next I
        PopBoard LastPushedX, LastPushedY

        PushBoard MouseX, MouseY
        LastPushedX = MouseX
        LastPushedY = MouseY
        If (basWin32API.GetKeyState(basWin32API.VK_RBUTTON) And &HF0000000) Or Button = 4 Then
            Board.MentionedCellsInit MouseX, MouseY
            For I = 1 To Board.MentionedCellsCount
                PushBoard MouseX + Board.MentionedCellsDeltaX(I - 1), MouseY + Board.MentionedCellsDeltaY(I - 1)
            Next I
            PushBoard MouseX, MouseY
        End If
    End If
End Sub

Private Sub PushBoard(ByVal X As Integer, ByVal Y As Integer)
    If BoardData(Board.Pos2Dto1D(X, Y)) = 0 Then PaintBoard X, Y, 6
End Sub

Private Sub PopBoard(ByVal X As Integer, ByVal Y As Integer)
    If BoardData(Board.Pos2Dto1D(X, Y)) = 0 Then PaintBoard X, Y, 0
End Sub

Private Sub PaintBoard(ByVal X As Integer, ByVal Y As Integer, ByVal Status As Integer, Optional ByVal Num As Integer = 0)
    If Not (X <= Board.Width - 1 And X >= 0 And Y <= Board.Height - 1 And Y >= 0) Then Exit Sub
    Me.picBoard.PaintPicture BoardImage, CalcBoardX(X, Y), CalcBoardY(X, Y), BoardImageWidth, BoardImageHeight, 0, 7 * BoardImageHeight, BoardImageWidth, BoardImageHeight, vbMergePaint
    Select Case Status
        Case 0, 6
            Me.picBoard.PaintPicture BoardImage, CalcBoardX(X, Y), CalcBoardY(X, Y), BoardImageWidth, BoardImageHeight, 0, Status * BoardImageHeight, BoardImageWidth, BoardImageHeight, vbSrcAnd
            If Status = 6 And Num > 0 Then PaintSmallDigital X, Y, Num, -1
        Case 1 To 5
            If Num <= 4 Then
                Me.picBoard.PaintPicture BoardImage, CalcBoardX(X, Y), CalcBoardY(X, Y), BoardImageWidth, BoardImageHeight, BoardImageWidth * Num - BoardImageWidth, Status * BoardImageHeight, BoardImageWidth, BoardImageHeight, vbSrcAnd
            Else
                Me.picBoard.PaintPicture BoardImage, CalcBoardX(X, Y), CalcBoardY(X, Y), BoardImageWidth, BoardImageHeight, BoardImageWidth * 4, Status * BoardImageHeight, BoardImageWidth, BoardImageHeight, vbSrcAnd
                PaintSmallDigital X, Y, Num, Choose(Status, 5, 5, 12, 0, 5)
            End If
        Case 6
            Me.picBoard.PaintPicture BoardImage, CalcBoardX(X, Y), CalcBoardY(X, Y), BoardImageWidth, BoardImageHeight, 0, Status * BoardImageHeight, BoardImageWidth, BoardImageHeight, vbSrcAnd
    End Select
End Sub

Private Sub PaintSmallDigital(ByVal X As Integer, ByVal Y As Integer, ByVal value As Integer, Optional ByVal Color As Integer = 0)
    Dim Value1 As Integer
    Dim Value2 As Integer
    Value1 = value \ 10
    Value2 = value Mod 10
    Dim DigitalWidth As Integer, DigitalHeight As Integer, DigitalX As Integer, DigitalY As Integer
    DigitalWidth = 13
    DigitalHeight = 13
    DigitalX = (BoardImageWidth - DigitalWidth) / 2
    DigitalY = (BoardImageHeight - DigitalHeight) / 2
    Dim absColor As Integer
    If Value1 > 0 Then
        If Color = -1 Then absColor = Choose(Value1 + 1, 9, 2, 1, 0, 13, 3, 7, 8, 6, 10) Else absColor = Color
        Me.picBoard.PaintPicture SmallDigitalImage, CalcBoardX(X, Y) + DigitalX, CalcBoardY(X, Y) + DigitalY, 6, 13, 30, 13 * (1 + (10 - Value1)), 6, 13, vbMergePaint
        Me.picBoard.PaintPicture SmallDigitalImage, CalcBoardX(X, Y) + DigitalX, CalcBoardY(X, Y) + DigitalY, 6, 13, 6 * absColor, 13 * (1 + (10 - Value1)), 6, 13, vbSrcAnd
    End If
    If Color = -1 Then absColor = Choose(Value2 + 1, 9, 2, 1, 0, 13, 3, 7, 8, 6, 10) Else absColor = Color
    Me.picBoard.PaintPicture SmallDigitalImage, CalcBoardX(X, Y) + DigitalX + 7, CalcBoardY(X, Y) + DigitalY, 6, 13, 30, 13 * (1 + (10 - Value2)), 6, 13, vbMergePaint
    Me.picBoard.PaintPicture SmallDigitalImage, CalcBoardX(X, Y) + DigitalX + 7, CalcBoardY(X, Y) + DigitalY, 6, 13, 6 * absColor, 13 * (1 + (10 - Value2)), 6, 13, vbSrcAnd
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MousePos As Variant
    picBoard_MouseMove Button, Shift, X, Y
    MousePos = CalcBoardMouseXY(X, Y)
    MouseX = MousePos(0)
    MouseY = MousePos(1)
    If (MouseLeftPushed And MouseRightPushed) Or Button = 4 Or ProcessingDblClick = True Then
        ProcessingDblClick = False
        If Button = 1 Then MouseLeftReleased = True
        If Button = 2 Then MouseRightReleased = True
        If MouseLeftReleased And MouseRightReleased Then
            MouseLeftPushed = False
            MouseRightPushed = False
            MouseLeftReleased = False
            MouseRightReleased = False
            LastPushedX = -100
            LastPushedY = -100
        End If
        For X = 0 To Board.Width
            For Y = 0 To Board.Height
                PopBoard X, Y
            Next Y
        Next X
        DoubleClickBoard MouseX, MouseY
        Exit Sub
    Else
        MouseLeftPushed = False
        MouseRightPushed = False
        MouseLeftReleased = False
        MouseRightReleased = False
        LastPushedX = -100
        LastPushedY = -100
    End If
    If Button = 1 Then OpenBoard MouseX, MouseY
End Sub

Private Function DoubleClickBoard(ByVal X As Integer, ByVal Y As Integer)
    If Board.Board2D(X, Y) > 0 And BoardData(Board.Pos2Dto1D(X, Y)) = 1 Then
        Dim I As Integer, intCount As Integer, intCell As Integer
        Board.MentionedCellsInit X, Y
        intCount = 0
        For I = 1 To Board.MentionedCellsCount
            intCell = BoardData(Board.Pos2Dto1D(X + Board.MentionedCellsDeltaX(I - 1), Y + Board.MentionedCellsDeltaY(I - 1)))
            If intCell < 0 And intCell > -500 Then
                intCount = intCount - intCell
            End If
        Next I
        If intCount = Board.Board2D(X, Y) Then
            For I = 1 To Board.MentionedCellsCount
                intCell = BoardData(Board.Pos2Dto1D(X + Board.MentionedCellsDeltaX(I - 1), Y + Board.MentionedCellsDeltaY(I - 1)))
                If Not (intCell < 0 And intCell > -500) Then
                    OpenBoard X + Board.MentionedCellsDeltaX(I - 1), Y + Board.MentionedCellsDeltaY(I - 1)
                End If
            Next I
        End If
    End If
End Function

Private Function OpenBoard(ByVal X As Integer, ByVal Y As Integer, Optional ByVal CallInOpenBoard As Boolean = False)
    If X <= Board.Width - 1 And X >= 0 And Y <= Board.Height - 1 And Y >= 0 And (BoardData(Board.Pos2Dto1D(X, Y)) = 0 Or BoardData(Board.Pos2Dto1D(X, Y)) < -500) Then
        If Board.Board2D(X, Y) < 0 And CallInOpenBoard = True Then
            Exit Function
        End If
        If FirstClick = True Then
            FirstClick = False
            GameStartTime = Timer
            tmrGame.Enabled = True
            tmrGame_Timer
        End If
        BoardData(Board.Pos2Dto1D(X, Y)) = 1
        UnopenedCells.Remove (Format(X) & "#" & Format(Y))
        If Board.Board2D(X, Y) < 0 Then
            GameEnd X, Y
            Exit Function
        Else
            PaintBoard X, Y, 6, Board.Board2D(X, Y)
            If UnopenedCells.Count = GameBoardMineCells Then
                GameWin X, Y
            End If
        End If
        OpenBoard = Board.Board2D(X, Y)

        If OpenBoard = 0 Then
            Dim I As Integer
            Board.MentionedCellsInit X, Y
            For I = 1 To Board.MentionedCellsCount
                OpenBoard X + Board.MentionedCellsDeltaX(I - 1), Y + Board.MentionedCellsDeltaY(I - 1), True
            Next I
        End If
    End If
End Function

Private Sub GameEnd(ByVal LastX As Integer, ByVal LastY As Integer)
    Dim Y As Integer
    Dim X As Integer
    Me.picBoard.Enabled = False
    Me.tmrGame.Enabled = False
    GameEndTime = Timer
    tmrGame_Timer
    For Y = 0 To Board.Height
        For X = 0 To Board.Width
            Dim data As Integer
            data = BoardData(Board.Pos2Dto1D(X, Y))
            If data < 0 And data > -500 Then
                data = Abs(data)
                If Board.Board2D(X, Y) < 0 Then
                    If data = -Board.Board2D(X, Y) Then
                        PaintBoard X, Y, 1, data
                    Else
                        PaintBoard X, Y, 4, -Board.Board2D(X, Y)
                    End If
                Else
                    PaintBoard X, Y, 4, data
                End If
            Else
                If Board.Board2D(X, Y) < 0 Then
                    PaintBoard X, Y, 5, -Board.Board2D(X, Y)
                End If
            End If
        Next X
    Next Y
    PaintBoard LastX, LastY, 3, -Board.Board2D(LastX, LastY)
    If ScreenSaverMode Then
        tmrScreenSaverAutoplayer.Enabled = False
        tmrScreenSaverWaiter.Enabled = True
    End If
End Sub

Private Sub GameWin(ByVal LastX As Integer, ByVal LastY As Integer)
    Dim Y As Integer
    Dim X As Integer
    Me.picBoard.Enabled = False
    Me.tmrGame.Enabled = False
    GameEndTime = Timer
    tmrGame_Timer
    For Y = 0 To Board.Height
        For X = 0 To Board.Width
            If Board.Board2D(X, Y) < 0 Then
                PaintBoard X, Y, 1, -Board.Board2D(X, Y)
            End If
        Next X
    Next Y
    If ScreenSaverMode Then
        tmrScreenSaverAutoplayer.Enabled = False
        tmrScreenSaverWaiter.Enabled = True
    End If
End Sub

Private Property Let BoardData(ByVal Pos As Integer, ByVal value As Integer)
    If Pos >= LBound(m_BoardData) And Pos <= UBound(m_BoardData) Then
        m_BoardData(Pos) = value
    End If
End Property

Private Property Get BoardData(ByVal Pos As Integer) As Integer
    If Pos >= LBound(m_BoardData) And Pos <= UBound(m_BoardData) Then
        BoardData = m_BoardData(Pos)
    End If
End Property

Private Sub AddDifficultyItem(ByVal CaptionResourceID As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Mines As Integer, ByVal MaxMinesPerCell As Integer)
    Load Me.mnuOptionsDifficultyItems(Me.mnuOptionsDifficultyItems.UBound + 1)
    With Me.mnuOptionsDifficultyItems(Me.mnuOptionsDifficultyItems.UBound)
        .Visible = True
        .Caption = LoadResString(CaptionResourceID)
        .Checked = False
        .Tag = Width & "," & Height & "," & Mines & "," & MaxMinesPerCell
    End With
End Sub

Private Function CalcBoardWidth() As Integer
    Select Case GameBoardType
        Case 1
            CalcBoardWidth = 16 * Board.Width + 8
        Case 2
            CalcBoardWidth = 22 * Board.Width + 10
        Case Else
            CalcBoardWidth = 16 * Board.Width
    End Select
End Function

Private Function CalcBoardHeight() As Integer
    Select Case GameBoardType
        Case 1
            CalcBoardHeight = 14 * Board.Height + 4
        Case 2
            CalcBoardHeight = 11 * Board.Height + 11
        Case Else
            CalcBoardHeight = 16 * Board.Height
    End Select
End Function

Private Function CalcBoardMouseXY(ByVal X As Integer, ByVal Y As Integer)
    Select Case GameBoardType
        Case 1
            CalcBoardMouseXY = Array((X - IIf((Y \ 14) Mod 2 = 0, 0, 8)) \ BoardImageWidth, Y \ 14)
        Case 2
            Dim tX As Integer, tY As Integer, pX As Integer, pY As Integer
            X = X + 1
            Y = Y + 1
            tX = (X \ 22) * 2 + IIf(X Mod 22 < 11, 0, 1)
            pX = (X Mod 22) Mod 11
            tY = (Y \ 11)
            pY = (Y Mod 11)
            Y = tY - 1
            Select Case (tX + tY) Mod 2
                Case 0
                    If tX = 0 And pY <= 11 - pX Then X = -100
                    If pY > 11 - pX Then Y = Y + 1
                Case 1
                    If tX = 0 And pY > pX Then X = -100
                    If pY >= pX Then Y = Y + 1
            End Select
            If X <> -100 Then X = (tX - (Y Mod 2)) \ 2 Else X = -1
            CalcBoardMouseXY = Array(X, Y)
        Case Else
            CalcBoardMouseXY = Array(X \ BoardImageWidth, Y \ BoardImageHeight)
    End Select
End Function

Private Function CalcBoardX(ByVal X As Integer, ByVal Y As Integer) As Integer
    Select Case GameBoardType
        Case 1
            CalcBoardX = X * BoardImageWidth + IIf(Y Mod 2 = 0, 0, 8)
        Case 2
            CalcBoardX = X * 22 + IIf(Y Mod 2 = 0, 0, 11)
        Case Else
            CalcBoardX = X * BoardImageWidth
    End Select
End Function

Private Function CalcBoardY(ByVal X As Integer, ByVal Y As Integer) As Integer
    Select Case GameBoardType
        Case 1
            CalcBoardY = Y * (BoardImageHeight - 4)
        Case 2
            CalcBoardY = Y * 11
        Case Else
            CalcBoardY = Y * BoardImageHeight
    End Select
End Function

Private Sub LoadBoard(ByVal intGameBoardType As Integer)
    GameBoardType = intGameBoardType
    Select Case GameBoardType
        Case 1
            Set BoardImage = LoadResPicture(4, vbResBitmap)
        Case 2
            Set BoardImage = LoadResPicture(5, vbResBitmap)
        Case Else
            Set BoardImage = LoadResPicture(3, vbResBitmap)
            GameBoardType = 0
    End Select
    BoardImageWidth = ScaleX(BoardImage.Width, vbHimetric, vbPixels) / 5
    BoardImageHeight = ScaleY(BoardImage.Height, vbHimetric, vbPixels) / 8
End Sub

Private Sub tmrClock_Timer()
    On Error Resume Next
    Me.ledBoard.LEDs(, "date").SetDate
    Err.Clear
    Me.ledBoard.LEDs(, "time").SetTime
    Err.Clear
End Sub

Private Sub tmrGame_Timer()
    Dim sngStart As Single
    Dim sngEnd As Single
    Dim sngPeriod As Single
    sngStart = IIf(GameStartTime = -1, Timer, GameStartTime)
    sngEnd = IIf(GameEndTime = -1, Timer, GameEndTime)
    sngPeriod = sngEnd - sngStart
    If sngPeriod < 0 Then sngPeriod = sngPeriod + 86400
    On Error Resume Next
    Me.ledBoard.LEDs(, "timer").SetTimerMinuteSecondMilesecond sngPeriod
    Err.Clear
End Sub

Private Sub tmrScreenSaverAutoplayer_Timer()
    Dim N As Integer, tempArray As Variant, X As Integer, Y As Integer, Z As Integer
    On Error GoTo hErr
again:
    Randomize
    N = Int(Rnd() * (UnopenedCells.Count) + 1)
    tempArray = Split(UnopenedCells.Item(N), "#")
    X = Val(tempArray(0))
    Y = Val(tempArray(1))
    Z = Int(Rnd() * 100)

    If Board.Board2D(X, Y) < 0 Then
        If Z < 5 Then
            OpenBoard X, Y
        Else
            Dim intNum As Integer
            Randomize
            If Int(Rnd() * 100) > 95 Then
                intNum = Abs(Board.Board2D(X, Y) + Int(Rnd() * 2)) Mod GameMaxMinesPerCell
                If intNum = 0 Then intNum = Int(Rnd() * (GameMaxMinesPerCell - 1)) + 1
            Else
                intNum = -Board.Board2D(X, Y)
            End If
            CheckMinesCounter BoardData(Board.Pos2Dto1D(X, Y)), -intNum
            BoardData(Board.Pos2Dto1D(X, Y)) = -intNum
            PaintBoard X, Y, 1, intNum
        End If
    Else
        OpenBoard X, Y
    End If
    Exit Sub
hErr:
    Err.Clear
    Resume Next
End Sub

Private Sub tmrScreenSaverWaiter_Timer()
    NewGame
    tmrScreenSaverAutoplayer.Enabled = True
    tmrScreenSaverWaiter.Enabled = False
End Sub

Private Sub Form_Resize()
    Me.ledBoard.Move 8, 8, Me.ledBoard.Width, Me.ScaleHeight - 16
    Me.shapeBackground.Move Me.ledBoard.Width + ledBoard.Left + 8 + 2, 8, Me.ScaleWidth - Me.ledBoard.Width - ledBoard.Left - 8 - 2 - 8, Me.ScaleHeight - 16
    Me.picBoard.Move Me.shapeBackground.Left + (Me.shapeBackground.Width - Me.picBoard.Width) / 2, Me.shapeBackground.Top + (Me.shapeBackground.Height - Me.picBoard.Height) / 2
    If Me.WindowState <> 0 Then Exit Sub
    Static lngWidth As Long, lngHeight As Long
    lngWidth = Me.shapeBackground.Left + Me.picBoard.Width + 8 + 2 + 16
    lngHeight = Me.shapeBackground.Top + Me.picBoard.Height + 8 + 2 + 16
    If Me.ScaleWidth <> lngWidth Or Me.ScaleHeight <> lngHeight Then
        basForm.SetClientRect Me, lngWidth * Screen.TwipsPerPixelX, lngHeight * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub DrawBackground(ByVal sngX As Single, ByVal sngY As Single, ByVal sngWidth As Single, ByVal sngHeight As Single, Optional blnSunken As Boolean = False)
    ForeColor = IIf(blnSunken, &HFFFFFF, &H808080)
    Line (sngX - 1, sngY - 1)-(sngX + sngWidth - 1, sngY - 1), , BF
    Line (sngX - 2, sngY - 2)-(sngX + sngWidth, sngY - 2), , BF
    Line (sngX - 1, sngY - 1)-(sngX - 1, sngY + sngHeight - 1), , BF
    Line (sngX - 2, sngY - 1)-(sngX - 2, sngY + sngHeight), , BF
    ForeColor = IIf(Not blnSunken, &HFFFFFF, &H808080)
    Line (sngX - 1, sngY + sngHeight)-(sngX + sngWidth + 1, sngY + sngHeight), , BF
    Line (sngX - 2, sngY + sngHeight + 1)-(sngX + sngWidth + 1, sngY + sngHeight + 1), , BF
    Line (sngX + sngWidth + 1, sngY - 1)-(sngX + sngWidth + 1, sngY + sngHeight - 1), , BF
    Line (sngX + sngWidth, sngY)-(sngX + sngWidth, sngY + sngHeight), , BF
End Sub

Private Sub Form_Paint()
    DrawBackground 2, 2, Me.ScaleWidth - 4, Me.ScaleHeight - 4, True
    ForeColor = &HC0C0C0
    Line (2, 2)-(Me.ScaleWidth - 4, Me.ScaleHeight - 4), , BF
    DrawBackground ledBoard.Left, ledBoard.Top, ledBoard.Width, ledBoard.Height, False
    DrawBackground shapeBackground.Left, shapeBackground.Top, shapeBackground.Width, shapeBackground.Height, False
End Sub

Private Sub UpdateMinesCounter()
    Dim I As Integer
    UpdateMinesCounterInner "mines", MinesCount - FlaggedMinesCount
    For I = 1 To Board.MaxMinesPerCell
        UpdateMinesCounterInner "flags" & Format(I), FlagsCount(I - 1) - FlaggedFlagsCount(I - 1)
    Next I
End Sub

Private Sub UpdateMinesCounterInner(ByVal strKey As String, ByVal lngNewValue As Long)
    On Error Resume Next
    If Val(ledBoard.LEDs(, strKey).Text) <> lngNewValue Then
        ledBoard.LEDs(, strKey).Text = Format(lngNewValue)
    End If
End Sub

Private Sub CheckMinesCounter(ByVal lngOldValue As Long, ByVal lngNewValue As Long)
    Select Case lngOldValue
        Case 0
        Case Is < -500
        Case Is < 0
            FlaggedMinesCount = FlaggedMinesCount - Abs(lngOldValue)
            FlaggedFlagsCount(Abs(lngOldValue) - 1) = FlaggedFlagsCount(Abs(lngOldValue) - 1) - 1
    End Select
    Select Case lngNewValue
        Case 0
        Case Is < -500
        Case Is < 0
            FlaggedMinesCount = FlaggedMinesCount + Abs(lngNewValue)
            FlaggedFlagsCount(Abs(lngNewValue) - 1) = FlaggedFlagsCount(Abs(lngNewValue) - 1) + 1
    End Select
    UpdateMinesCounter
End Sub
