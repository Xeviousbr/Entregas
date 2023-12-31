VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfigImpr 
   Caption         =   "Configuração de impressoras"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4245
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4245
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   9
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3600
      Width           =   3855
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   9
      Left            =   3960
      TabIndex        =   22
      Top             =   3600
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   8
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   20
      Top             =   3240
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   7
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2880
      Width           =   3855
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   7
      Left            =   3960
      TabIndex        =   18
      Top             =   2880
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   6
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   16
      Top             =   2520
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   5
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   14
      Top             =   2160
      Width           =   195
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2340
      TabIndex        =   13
      Top             =   4020
      Width           =   1215
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   12
      Top             =   1800
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   4
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   10
      Top             =   1440
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   3
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   8
      Top             =   1080
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   2
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   1
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   435
      Left            =   555
      TabIndex        =   2
      Top             =   4020
      Width           =   1215
   End
   Begin VB.CheckBox ckImpress 
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   195
   End
   Begin VB.TextBox txImpress 
      Height          =   285
      Index           =   0
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Fita"
      Height          =   255
      Left            =   3900
      TabIndex        =   4
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Impressora"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfigImpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.1 Aumento da previsão de impressoras instaladas de 5 para 10
'4.9.8 Selecionar a impressora ao imprimir
'4.9.7 Tela para gravar se a impressora é tipo fita ou não

Option Explicit

Private bMenu    As Boolean
Private gOK      As Boolean
Private JaAtivou As Boolean
Private MaxImpr       As Integer
Private ImprEscolhida As Integer
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Sub ckImpress_Click(Index As Integer)
'4.9.8 Selecionar a impressora ao imprimir
Static dentro As Boolean
Dim Vlr       As Integer
Dim X         As Long

If JaAtivou Then
    X = DoEvents
    Sleep 100
    If dentro = False Then
        dentro = True
        
        '4.9.9 Ajustes na escolha da impressora
        Vlr = ckImpress(Index).Value
        
        ckImpress(ImprEscolhida).Value = 0
        ImprEscolhida = Index
        ckImpress(ImprEscolhida).Value = Vlr
    End If
    dentro = False
End If
End Sub

Private Sub Command1_Click()
Dim a As Integer

If bMenu Then
    For a = 0 To MaxImpr
        ImprFitaNr a, ckImpress(a).Value
    Next
    Unload Me
Else
    OK = True
    INI.ImprEscolhida = ImprEscolhida
    Me.Hide
End If
End Sub

Private Sub Command2_Click()
If bMenu Then
    Unload Me
Else
    OK = False
    Me.Hide
End If
End Sub

Private Sub Form_Activate()
Dim lngImpr As Long
Dim Buffer  As String

If JaAtivou = False Then
    Buffer = Space(8192)
    lngImpr = GetProfileString("PrinterPorts", vbNullString, "", Buffer, Len(Buffer))
    SelecionaImpressora Buffer
    If bMenu = False Then
        ImprEscolhida = INI.ImprEscolhida
        ckImpress(ImprEscolhida).Value = 1
        Label2.Visible = False
    End If
    JaAtivou = True
End If
End Sub

Private Sub Form_Load()
InicForm Me
End Sub

Private Sub SelecionaImpressora(ByVal Buffer As String)
Dim i    As Integer
Dim intI As Integer
Dim strS As String

Do
    intI = InStr(Buffer, Chr(0))
    If intI > 0 Then
        strS = Left(Buffer, intI - 1)
        If Len(Trim(strS)) Then
            AnotaImpr i, strS
        End If
        Buffer = Mid(Buffer, intI + 1)
    Else
        If Len(Trim(Buffer)) Then
            AnotaImpr i, Buffer
        End If
        Buffer = ""
    End If
    i = i + 1
Loop While intI > 0
End Sub

Private Sub AnotaImpr(Nr As Integer, Nome As String)
txImpress(Nr).Text = Nome
If bMenu Then
    ckImpress(Nr).Value = INI.ImprFitaNr(Nr)
Else
    ckImpress(Nr).Value = 0
End If
MaxImpr = Nr
End Sub

Private Sub ImprFitaNr(Posic As Integer, Opc As Integer)
Dim nrfita As String
Dim sOpc   As String

nrfita = "Fita" & Trim(STR(Posic))
sOpc = Trim(STR(Opc))
WritePrivateProfileString "Impressao", nrfita, sOpc, AplicDirNat & "\Orcarro.ini"
End Sub

'4.9.8 Selecionar a impressora ao imprimir
Public Sub Menu()
bMenu = True
Command1.Caption = "Salvar"
Command2.Caption = "Fechar"
End Sub

'4.9.8 Selecionar a impressora ao imprimir
Public Property Get OK() As Boolean
OK = gOK
End Property

'4.9.8 Selecionar a impressora ao imprimir
Public Property Let OK(ByVal vNewValue As Boolean)
gOK = vNewValue
End Property

Public Property Get ImprFita() As Integer
ImprFita = INI.ImprFitaNr(ImprEscolhida)
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
JaAtivou = False
End Sub
