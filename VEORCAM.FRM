VERSION 4.00
Begin VB.Form VeOrcam 
   Caption         =   "Selecione o orçamento desejado"
   ClientHeight    =   2100
   ClientLeft      =   1590
   ClientTop       =   2280
   ClientWidth     =   5235
   Height          =   2520
   Left            =   1530
   LinkTopic       =   "Form2"
   ScaleHeight     =   2100
   ScaleWidth      =   5235
   Top             =   1920
   Width           =   5355
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\PROG\TROJAN\TROJAN.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VeSeTemOrcamento"
      Top             =   1560
      Width           =   1215
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "VEORCAM.frx":0000
      DataField       =   "Cliente"
      DataSource      =   "Data1"
      Height          =   1035
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      _version        =   65536
      _extentx        =   7011
      _extenty        =   1826
      _stockprops     =   77
      forecolor       =   -2147483640
      backcolor       =   -2147483643
      listfield       =   "Data"
      boundcolumn     =   "Cliente"
   End
End
Attribute VB_Name = "VeOrcam"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

