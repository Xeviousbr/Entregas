VERSION 5.00
Begin VB.UserControl Chart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   4  'Dash-Dot
      Height          =   1455
      Left            =   1080
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Menu mnuMain 
      Caption         =   "mnuMain"
      Begin VB.Menu mnuLegend 
         Caption         =   "Legenda"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2.9.8 Previsão de erro no gráfico
'2.9.7 Relatório de Totais Resumido

Option Explicit
'Default Property Values:
Const m_def_MinValue = 0
Const m_def_MaxValue = 5
Const m_def_Rows = 0
Const m_def_Cols = 0
'Property Variables:
Private m_MinValue As Integer
Private m_MaxValue As Integer
Private m_Rows As Integer
Private m_Cols As Integer

Private RowOffset As Integer, ColOffset As Integer
Private LegendX1 As Integer, LegendY1 As Integer, LegendX2 As Integer, LegendY2 As Integer
Private IsMovingLegend As Boolean
Private tmpOffsetY As Integer, tmpOffsetX As Integer

Private CX(120) As Single

Public Sub DrawGraph(LinesArray() As String, ColorArray() As Long, RowCaption() As String)
Dim i As Integer, RowSize As Integer, ColSize As Integer, ColValue As Integer
Dim StepSize As Single, ArrayIndex As Integer, LineDimensions() As String
Dim LineColor As Long
Dim CurrXBak As Integer
Dim FirstPoint As Long, SecondPoint As Long

'2.9.8 Previsão de erro no gráfico
If MaxValue = 0 Then MaxValue = 1

RowOffset = 500
ColOffset = 500
With UserControl
    .Cls
    ' izracunamo maksimalan broj stupaca i kolona koje stanu u polje za crtanje
    RowSize = (.Width - ColOffset) / Rows
    ColSize = (.Height - RowOffset) / MaxValue
    'ColSize = .Height / MaxValue - RowOffset
    ' iscrtamo border oko djela gdje ce bit linije
    .BackColor = RGB(255, 255, 246)
    .DrawStyle = vbSolid
    Line (ColOffset, 0)-(.Width - 10, .Height - RowOffset), vbBlack, B
    
    .DrawStyle = vbDot
    For i = 1 To Rows - 1 ' crtamo linije za stupce
        Line (ColOffset + (i * RowSize), 0)-(ColOffset + (i * RowSize), .Height - RowOffset), RGB(192, 192, 192)
    Next i
    For i = 1 To Rows ' ispisujemo caption stupca
        CurrentY = .Height - (RowOffset / 2) - (TextHeight("I") / 2)
        CurrentX = i * RowSize - (TextWidth(RowCaption(i)) / 2)
        CX(i) = CurrentX
        Print RowCaption(i)
    Next i
    
    ' ispisujemo vrijednosti kolona
    StepSize = -(MaxValue / 5) ' izracunamo da prikazuje s odredjenim razmakom brojeve
    If StepSize > -0.6 Then StepSize = -1
    For i = MaxValue To MinValue Step StepSize
        ColValue = (i * -1) + MaxValue
        CurrentX = (ColOffset / 2) - (TextWidth(ColValue) / 2)
        CurrentY = i * ColSize
        Print ColValue
    Next i
    ' zadnju ispisemo jos jednom u slucaju da nam je step izracunao bez zadnje
    CurrentX = (ColOffset / 2) - (TextWidth(ColValue) / 2)
    CurrentY = 0
    Print MaxValue
    
    ' iscrtavamo linije
    .DrawStyle = vbSolid
    For ArrayIndex = LBound(LinesArray) To UBound(LinesArray)
        LineColor = ColorArray(ArrayIndex)
        LineDimensions = Split(LinesArray(ArrayIndex), ",")
            For i = LBound(LineDimensions) To UBound(LineDimensions) - 1
                LineDimensions(i) = Replace(LineDimensions(i), ".", ",")
                LineDimensions(i + 1) = Replace(LineDimensions(i + 1), ".", ",")
                
                On Local Error GoTo Erro_First
                FirstPoint = (.Height - RowOffset) - (CInt(LineDimensions(i)) * ColSize)
First:
                On Local Error GoTo Erro_Second
                SecondPoint = (.Height - RowOffset) - (CInt(LineDimensions(i + 1)) * ColSize)
Second:
                CurrentY = FirstPoint + 10
                CurrentX = CX(i + 1)
                Print CInt(LineDimensions(i))
                Line (ColOffset + (i * RowSize) + (RowSize / 2), FirstPoint)-(ColOffset + RowSize + (i * RowSize) + (RowSize / 2), SecondPoint), LineColor
            Next i
            Print CInt(LineDimensions(i))
    Next ArrayIndex
End With
Exit Sub

Erro_First:
FirstPoint = 0
Resume First

Erro_Second:
SecondPoint = 0
Resume Second
End Sub

Public Sub DrawLegend(LegendArray() As String, ColorArray() As Long, Complemento As String)
Dim MaxLength As Integer, StartTop As Integer, i As Integer, tmpPos As Integer

If mnuLegend.Checked Then
    If LegendX1 = 0 Then LegendX1 = 1000
    If LegendY1 = 0 Then LegendY1 = 1000
   
    CurrentY = LegendY1 + 100
    StartTop = CurrentY
    ' uzmemo najduzu rijec da znamo kolko velika ce legenda bit
    For i = LBound(LegendArray) To UBound(LegendArray)
        If MaxLength < TextWidth(LegendArray(i)) Then MaxLength = TextWidth(LegendArray(i))
        Print
    Next i
    Print
    If MaxLength < TextWidth(Complemento) Then MaxLength = TextWidth(Complemento)
        
    LegendX2 = LegendX1 + MaxLength + 300
    
    Line (LegendX1, LegendY1)-(LegendX2, CurrentY + 100), vbWhite, BF
    Line (LegendX1, LegendY1)-(LegendX2, CurrentY), vbGrayText, B
    
    CurrentY = StartTop
    ' ispisemo opis objekata legende i boje
    tmpPos = CurrentY
    For i = LBound(LegendArray) To UBound(LegendArray)
        tmpPos = CurrentY
        Line (LegendX1 + 200, tmpPos + 50)-(LegendX1 + 400, tmpPos + 125), ColorArray(i), BF
        CurrentX = LegendX1 + 600
        CurrentY = tmpPos
        Print LegendArray(i)
    Next i
    CurrentX = LegendX1 + 200
    Print Complemento
    LegendY2 = CurrentY + 100
End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Rows() As Integer
    Rows = m_Rows
End Property

Public Property Let Rows(ByVal New_Rows As Integer)
    m_Rows = New_Rows
    PropertyChanged "Rows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Cols() As Integer
    Cols = m_Cols
End Property

Public Property Let Cols(ByVal New_Cols As Integer)
    m_Cols = New_Cols
    PropertyChanged "Cols"
End Property

Private Sub mnuLegend_Click()
If mnuLegend.Checked Then
    mnuLegend.Checked = False
Else
    mnuLegend.Checked = True
End If
Form1.RefreshGraph
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Rows = m_def_Rows
    m_Cols = m_def_Cols
    m_MinValue = m_def_MinValue
    m_MaxValue = m_def_MaxValue
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
IsMovingLegend = False
If Button = vbLeftButton Then
    If mnuLegend.Checked Then
        If X >= LegendX1 And X <= LegendX2 Then
            If Y >= LegendY1 And Y <= LegendY2 Then
                IsMovingLegend = True
                tmpOffsetY = Y - LegendY1
                tmpOffsetX = X - LegendX1
                Shape1.Top = LegendY1
                Shape1.Left = LegendX1
                Shape1.Height = LegendY2 - LegendY1
                Shape1.Width = LegendX2 - LegendX1
                Shape1.Visible = True
            End If
        End If
    End If
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IsMovingLegend Then
    Shape1.Top = Y - tmpOffsetY
    Shape1.Left = X - tmpOffsetX
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IsMovingLegend Then
IsMovingLegend = False
    Shape1.Visible = False
    LegendY1 = Shape1.Top
    LegendX1 = Shape1.Left
    LegendY2 = Shape1.Top + Shape1.Height
    LegendX2 = Shape1.Left + Shape1.Width
Form1.RefreshGraph
End If
If Button = vbRightButton Then
    If IsMovingLegend = False Then
        Shape1.Visible = False
        PopupMenu mnuMain
    End If
End If
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Rows = PropBag.ReadProperty("Rows", m_def_Rows)
    m_Cols = PropBag.ReadProperty("Cols", m_def_Cols)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
End Sub

Private Sub UserControl_Resize()
If Ambient.UserMode Then
    Form1.RefreshGraph
End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Rows", m_Rows, m_def_Rows)
    Call PropBag.WriteProperty("Cols", m_Cols, m_def_Cols)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MinValue() As Integer
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As Integer)
    m_MinValue = New_MinValue
    PropertyChanged "MinValue"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MaxValue() As Integer
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Integer)
    m_MaxValue = New_MaxValue
    PropertyChanged "MaxValue"
End Property

