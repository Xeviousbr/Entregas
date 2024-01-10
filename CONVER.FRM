VERSION 4.00
Begin VB.Form Conver 
   Caption         =   "Conversor"
   ClientHeight    =   1860
   ClientLeft      =   1560
   ClientTop       =   2280
   ClientWidth     =   5700
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   2280
   Icon            =   "CONVER.frx":0000
   Left            =   1500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   5700
   Top             =   1920
   Width           =   5820
   Begin VB.CommandButton BtCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton BtOK 
      Caption         =   "Converte"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin GaugeLib.Gauge Gauge 
      Height          =   270
      Left            =   270
      TabIndex        =   3
      Top             =   1485
      Visible         =   0   'False
      Width           =   5160
      _version        =   65536
      _extentx        =   9102
      _extenty        =   476
      _stockprops     =   73
      forecolor       =   -2147483646
      backcolor       =   16777215
      innertop        =   0
      innerleft       =   0
      innerright      =   0
      innerbottom     =   0
      max             =   10
      needlewidth     =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Converte o formato do banco de dados"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   13.5
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4920
   End
End
Attribute VB_Name = "Conver"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Private Sub BtCancel_Click()
End
End Sub

Private Sub BtOK_Click()
Dim db As Database
Dim tbPeca, tbItem, tbOrc As Recordset
Dim TbDef As TableDef
Dim NovoField, NovoField2 As Field
Dim x As Integer

BtOK.Enabled = False
Gauge.Visible = True

'Criar uma cópia do banco de dados antigo
Conver.Caption = "Criando uma reserva do banco de dados"
Gauge.Value = Gauge.Value + 1
x = DoEvents
FileCopy "Trojan.mdb", "Copia.mdb"
FileCopy "Trojan.ldb", "Copia.ldb"

'Acrescentar o ítem valor para ítem_orc
Conver.Caption = "Abrindo o banco de dados"
Gauge.Value = Gauge.Value + 1
x = DoEvents
Set db = OpenDatabase(CurDir$ + "\Trojan.mdb")

Conver.Caption = "Criando novo campo na tabela de itens"
Gauge.Value = Gauge.Value + 1
x = DoEvents
Set TbDef = db.TableDefs("Itens_Orc")
' Criar novo Field object.
Set NovoField = TbDef.CreateField("Valor", dbCurrency)
TbDef.Fields.Append NovoField

Conver.Caption = "Abrindo a tabela de Ítens"
Gauge.Value = Gauge.Value + 1
x = DoEvents
Set tbItem = db.OpenRecordset("Itens_Orc")

Conver.Caption = "Colocando valores na tabela de Itens"
Gauge.Value = Gauge.Value + 1
x = DoEvents
If tbItem.RecordCount > 0 Then
    Conver.Caption = "Abrindo a tabela de Peças"
    Set tbPeca = db.OpenRecordset("Pecas")
    tbPeca.Index = "PrimaryKey"
    
    tbItem.MoveFirst
    Do While tbItem.EOF = False
        tbPeca.Seek "=", tbItem("Item")
        tbItem.Edit
        If tbPeca.NoMatch = False Then
            tbItem("Valor") = 0
        Else
            tbItem("Valor") = tbPeca("Valor")
        End If
        tbItem.Update
        tbItem.MoveNext
    Loop
    tbPeca.Close
End If
tbItem.Close

If Command$ <> "ATC" Then
    'Apagar a tabela de peças
    Conver.Caption = "Apagando a tabela de peças"
    Gauge.Value = Gauge.Value + 1
    x = DoEvents
    db.TableDefs.Delete "Pecas"
End If

'Adicionar os campos Kms e Kilom
Conver.Caption = "Adicionando os campos Kms e Kilom"
Gauge.Value = Gauge.Value + 1
x = DoEvents
Set TbDef = db.TableDefs("Orcamento")
Set NovoField = TbDef.CreateField("Kms", dbLong)
Set NovoField2 = TbDef.CreateField("Kilom", dbLong)
TbDef.Fields.Append NovoField
TbDef.Fields.Append NovoField2

Conver.Caption = "Preenchento os campos Kms e Kilom"
Gauge.Value = Gauge.Value + 1
x = DoEvents
Set tbOrc = db.OpenRecordset("Orcamento")
tbOrc.MoveFirst
If tbOrc.RecordCount > 0 Then
    Do While tbOrc.EOF = False
        tbOrc.Edit
        tbOrc("Kms") = 0
        tbOrc("Kilom") = 0
        tbOrc.Update
        tbOrc.MoveNext
    Loop
End If
tbOrc.Close

Conver.Caption = "Fechando a tabela"
Gauge.Value = Gauge.Value + 1
x = DoEvents
db.Close

Gauge.Value = Gauge.Value + 1
x = DoEvents
Gauge.Visible = False
Conver.Caption = "Conversão concluída com exito"
MsgBox "Conversão concluída com exito"
End
End Sub





