VERSION 5.00
Begin VB.Form Senha 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   Caption         =   "Senha"
   ClientHeight    =   900
   ClientLeft      =   3780
   ClientTop       =   2985
   ClientWidth     =   1920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3.75
   ScaleMode       =   4  'Character
   ScaleWidth      =   16
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   360
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Digite a senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2.7.4 Identificação do Mecânico que esta operando

Option Explicit
Dim Texto As String

Private Sub Form_Load()
InicForm Me
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
Texto = ""
senhaOK = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
SemPlim KeyAscii
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Senha.Hide
Else
   If KeyCode = 13 Then
      If Consulta("Select Count(*) From Mecanicos Where Senha = '" & Text1.Text & "'") > 0 Then
          CodiMecanicoLogadoQueFezLoginAgora = Text1.Text
          Senha.Hide
      End If
   End If
End If
End Sub

