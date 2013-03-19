Attribute VB_Name = "basConst"
'Project name: orzMinesweeper
'Code license: GNU General Public License v3
'Author      : Yeechan Lu a.k.a. orzFly <i@orzfly.com>

'游戏局数区间
Public Const GameIDMin = 1#
Public Const GameIDMax = 999999999#
Public Const BoardWidthMin = 4
Public Const BoardWidthMax = 30
Public Const BoardHeightMin = 4
Public Const BoardHeightMax = 30
Public Const BoardMaxMinesPerCellMin = 1
Public Const BoardMaxMinesPerCellMax = 9
Public Const BoardMinesMin = 1

Public Property Get BoardMinesMax(ByVal intWidth As Integer, ByVal intHeight As Integer, ByVal intMaxMinesPerCell As Integer)
    BoardMinesMax = Int(basMath.Min(intWidth * intHeight * (1 + 0.314159 * intMaxMinesPerCell), (intWidth * intHeight * intMaxMinesPerCell) - intMaxMinesPerCell))
End Property
