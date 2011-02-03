VERSION 5.00
Begin VB.UserControl ctlLongNumberTextBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtBase 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "ctlLongNumberTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum enumLNTBBorderStyle
    lntbbsNone = 0
    lntbbsFixedSingle = 1
End Enum

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Private m_MinValue As Long
Private m_MaxValue As Long
Private m_Step As Long
Private m_LongStep As Long
Private m_MouseX As Integer
Private m_MouseY As Integer
Private m_MouseValue As Long

Event Click()
Event DblClick()
Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub txtBase_Change()
    Dim en As String
    en = txtBase.Text
    If Val(en) < m_MinValue Then en = Format(m_MinValue)
    If Val(en) > m_MaxValue Then en = Format(m_MaxValue)
    en = en
    If txtBase.Text <> en Then
        txtBase.Text = en
        SelectAll
        Beep
    End If
    RaiseEvent Change
End Sub

Private Sub txtBase_GotFocus()
    SelectAll
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = UserControl.Height - UserControl.ScaleHeight + 300
    UserControl.txtBase.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = txtBase.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtBase.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = txtBase.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtBase.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = txtBase.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtBase.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = txtBase.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtBase.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BorderStyle() As enumLNTBBorderStyle
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = txtBase.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As enumLNTBBorderStyle)
    txtBase.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "强制完全重画一个对象。"
    txtBase.Refresh
End Sub

Private Sub txtBase_Click()
    RaiseEvent Click
End Sub

Private Sub txtBase_DblClick()
    Random
    SelectAll
End Sub

Private Sub txtBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyUp Then
            KeyCode = 0
            value = Val(txtBase.Text) + m_Step
            SelectAll
        ElseIf KeyCode = vbKeyDown Then
            KeyCode = 0
            value = value - m_Step
            SelectAll
        ElseIf KeyCode = vbKeyPageUp Then
            KeyCode = 0
            value = Val(txtBase.Text) + m_LongStep
            SelectAll
        ElseIf KeyCode = vbKeyPageDown Then
            KeyCode = 0
            value = value - m_LongStep
            SelectAll
        End If
    End If
    If KeyCode <> 0 Then
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub txtBase_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            RaiseEvent KeyPress(KeyAscii)
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBase_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SetCapture Me.hWnd
        m_MouseX = X
        m_MouseY = Y
        m_MouseValue = value
    End If
End Sub

Private Sub txtBase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim DeltaX As Integer, DeltaY As Integer
        DeltaX = ((X - m_MouseX) / 15 \ 20) * LongStep
        DeltaY = -((Y - m_MouseY) / 15 \ 5) * Step
        Me.value = m_MouseValue + DeltaX + DeltaY
        SelectAll
    End If
End Sub

Private Sub txtBase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
End Sub

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "返回/设置选定的字符数。"
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = txtBase.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtBase.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "返回/设置选定文本的起始点。"
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txtBase.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtBase.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "返回/设置包含当前选定文本的字符串。"
Attribute SelText.VB_MemberFlags = "400"
    SelText = txtBase.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtBase.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Public Property Get value() As Long
Attribute value.VB_Description = "返回/设置控件中包含的文本。"
    value = Val(txtBase.Text)
End Property

Public Property Let value(ByVal New_Value As Long)
    txtBase.Text = Format(New_Value, "#")
    PropertyChanged "Value"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "返回一个句柄到(from Microsoft Windows)一个对象的窗口。"
    hWnd = txtBase.hWnd
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "决定控件是否可编辑。"
    Locked = txtBase.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtBase.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get MinValue() As Long
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As Long)
    m_MinValue = New_MinValue
    PropertyChanged "MinValue"
End Property

Public Property Get MaxValue() As Long
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Long)
    m_MaxValue = New_MaxValue
    PropertyChanged "MaxValue"
End Property

Public Property Get Step() As Long
    Step = m_Step
End Property

Public Property Let Step(ByVal New_Step As Long)
    m_Step = New_Step
    PropertyChanged "Step"
End Property

Public Property Get LongStep() As Long
    LongStep = m_LongStep
End Property

Public Property Let LongStep(ByVal New_LongStep As Long)
    m_LongStep = New_LongStep
    PropertyChanged "LongStep"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtBase.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtBase.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtBase.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtBase.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtBase.BorderStyle = PropBag.ReadProperty("BorderStyle", enumLNTBBorderStyle.lntbbsFixedSingle)
    txtBase.Locked = PropBag.ReadProperty("Locked", False)
    m_MaxValue = PropBag.ReadProperty("MaxValue", 100)
    m_MinValue = PropBag.ReadProperty("MinValue", 0)
    m_Step = PropBag.ReadProperty("Step", 1)
    m_LongStep = PropBag.ReadProperty("LongStep", 10)
    value = PropBag.ReadProperty("Value", "0")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", txtBase.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtBase.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtBase.Enabled, True)
    Call PropBag.WriteProperty("Font", txtBase.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", txtBase.BorderStyle, enumLNTBBorderStyle.lntbbsFixedSingle)
    Call PropBag.WriteProperty("Locked", txtBase.Locked, False)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, 100)
    Call PropBag.WriteProperty("MinValue", m_MinValue, 0)
    Call PropBag.WriteProperty("Step", m_Step, 1)
    Call PropBag.WriteProperty("LongStep", m_LongStep, 10)
    Call PropBag.WriteProperty("Value", Me.value, 0)
End Sub

Public Sub SelectAll()
    txtBase.SelStart = 0
    txtBase.SelLength = Len(txtBase)
End Sub

Public Sub Random()
    Randomize
    value = Int(Rnd() * (MaxValue - MinValue) + MinValue)
End Sub
