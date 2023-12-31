VERSION 5.00
Begin VB.Form frmManutencao 
   Caption         =   "Manutenção de Tarefas"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5145
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Procer ajuste"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1965
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Selecionar"
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Top             =   900
      Width           =   3555
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   540
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecionar"
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   180
      Width           =   3555
   End
   Begin VB.Label lbDescr 
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mecânico Destino:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Placa:"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mecânico Original:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "frmManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2.9.2 Manutenção para as tarefas

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

