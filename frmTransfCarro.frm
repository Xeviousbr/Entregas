VERSION 5.00
Begin VB.Form frmTransfCarro 
   Caption         =   "Transferência de carro no OrCarro"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6990
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btOK 
      Caption         =   "Salvar"
      Height          =   435
      Left            =   2940
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton btNovo 
      Caption         =   "Novo Cliente"
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      Top             =   900
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   900
      Width           =   4875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   900
      Width           =   855
   End
   Begin VB.Label lbCarro 
      AutoSize        =   -1  'True
      Caption         =   "Carro : RENAULT CLIO SEDAN 1.0 16V ANO 2001 XXX-MMMM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   6870
   End
   Begin VB.Label lbCliOrig 
      AutoSize        =   -1  'True
      Caption         =   "Cliente Original : ALEXANDRE SALIN [MERCADO DA FAMILIA]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmTransfCarro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
