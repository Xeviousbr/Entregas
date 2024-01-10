VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.Form frmOrc 
   Caption         =   "Orçamento"
   ClientHeight    =   7710
   ClientLeft      =   675
   ClientTop       =   405
   ClientWidth     =   8445
   Icon            =   "ORC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7710
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txNT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   2460
      TabIndex        =   36
      Top             =   4860
      Width           =   615
   End
   Begin VB.TextBox txNT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   2460
      TabIndex        =   32
      Top             =   4620
      Width           =   615
   End
   Begin VB.TextBox txNT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   2460
      TabIndex        =   28
      Top             =   4380
      Width           =   615
   End
   Begin VB.TextBox txNT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   2460
      TabIndex        =   24
      Top             =   4140
      Width           =   615
   End
   Begin VB.TextBox txNT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   2460
      TabIndex        =   20
      Top             =   3900
      Width           =   615
   End
   Begin VB.TextBox txNT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2460
      TabIndex        =   16
      Top             =   3660
      Width           =   615
   End
   Begin VB.TextBox txNT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2460
      TabIndex        =   12
      Top             =   3420
      Width           =   615
   End
   Begin VB.CommandButton Botao 
      Caption         =   "&Visualizar Orçamento"
      Height          =   315
      Index           =   4
      Left            =   3360
      TabIndex        =   91
      Top             =   7320
      Width           =   1635
   End
   Begin VB.CommandButton Botao 
      Caption         =   "&Imprimir Observação"
      Height          =   315
      Index           =   3
      Left            =   5220
      TabIndex        =   89
      Top             =   7320
      Width           =   1635
   End
   Begin VB.TextBox txNome 
      Height          =   285
      Index           =   6
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   88
      Top             =   4860
      Width           =   1245
   End
   Begin VB.TextBox txNome 
      Height          =   285
      Index           =   5
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   87
      Top             =   4620
      Width           =   1245
   End
   Begin VB.TextBox txNome 
      Height          =   285
      Index           =   4
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   4380
      Width           =   1245
   End
   Begin VB.TextBox txNome 
      Height          =   285
      Index           =   3
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   85
      Top             =   4140
      Width           =   1245
   End
   Begin VB.TextBox txNome 
      Height          =   285
      Index           =   2
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   84
      Top             =   3900
      Width           =   1245
   End
   Begin VB.TextBox txNome 
      Height          =   285
      Index           =   1
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   83
      Top             =   3660
      Width           =   1245
   End
   Begin VB.TextBox txNome 
      Height          =   285
      Index           =   0
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   82
      Top             =   3420
      Width           =   1245
   End
   Begin VB.TextBox txData 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   4860
      Width           =   855
   End
   Begin VB.TextBox txData 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   4620
      Width           =   855
   End
   Begin VB.TextBox txData 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   4380
      Width           =   855
   End
   Begin VB.TextBox txData 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   4140
      Width           =   855
   End
   Begin VB.TextBox txData 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   3900
      Width           =   855
   End
   Begin VB.TextBox txData 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   3660
      Width           =   855
   End
   Begin VB.TextBox txData 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   3420
      Width           =   855
   End
   Begin VB.Timer TimerItens 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7920
      Top             =   3600
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   11
      Left            =   2940
      MaxLength       =   30
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5160
      Width           =   375
   End
   Begin VB.Data dataConsertos 
      Caption         =   "Consertos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select concerto From tpConcertos Order By tipo"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin TrueDBGrid50.TDBDropDown tdbConsertos 
      Bindings        =   "ORC.frx":030A
      Height          =   1035
      Left            =   180
      OleObjectBlob   =   "ORC.frx":0326
      TabIndex        =   66
      Top             =   5220
      Width           =   1815
   End
   Begin VB.Data dataMecs 
      Caption         =   "Mecanicos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\vmware-host\Shared Folders\SourceTree\OrCarro\OrCarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Left            =   2820
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select Nome From Mecanicos Order by Nome"
      Top             =   4140
      Visible         =   0   'False
      Width           =   2415
   End
   Begin TrueDBGrid50.TDBDropDown tdbMecanicos 
      Bindings        =   "ORC.frx":17A0
      Height          =   1815
      Left            =   3540
      OleObjectBlob   =   "ORC.frx":17B7
      TabIndex        =   65
      Top             =   5040
      Width           =   2355
   End
   Begin VB.Data DataT 
      Caption         =   "DataT"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\vmware-host\Shared Folders\SourceTree\OrCarro\OrCarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TarefasTemp"
      Top             =   2460
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame4 
      Caption         =   "Consertos"
      Height          =   1935
      Left            =   120
      TabIndex        =   62
      Top             =   1320
      Width           =   8175
      Begin VB.PictureBox Picture1 
         Height          =   75
         Left            =   840
         ScaleHeight     =   75
         ScaleWidth      =   15
         TabIndex        =   64
         Top             =   540
         Width           =   15
      End
      Begin TrueDBGrid50.TDBGrid GridT 
         Bindings        =   "ORC.frx":2C25
         Height          =   1515
         Left            =   180
         OleObjectBlob   =   "ORC.frx":2C39
         TabIndex        =   63
         Top             =   300
         Width           =   7815
      End
   End
   Begin VB.TextBox Text 
      Height          =   855
      Index           =   13
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   61
      Top             =   6360
      Visible         =   0   'False
      Width           =   8235
   End
   Begin VB.OptionButton opObs 
      Caption         =   "Histórico"
      Height          =   255
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6060
      Width           =   795
   End
   Begin VB.TextBox Text 
      Height          =   855
      Index           =   12
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   59
      Top             =   6360
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.OptionButton opObs 
      Caption         =   "Observação do Mecânico"
      Height          =   255
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   6060
      Width           =   2115
   End
   Begin VB.OptionButton opObs 
      Caption         =   "Observação"
      Height          =   255
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6060
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1260
      Top             =   2940
   End
   Begin VB.OptionButton btPagamento 
      Caption         =   "Não Pago"
      Height          =   375
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5940
      Width           =   1275
   End
   Begin VB.TextBox Text 
      Height          =   855
      Index           =   8
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   41
      Top             =   6360
      Width           =   8175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Situação do carro"
      Height          =   975
      Left            =   5640
      TabIndex        =   53
      Top             =   420
      Width           =   1815
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   10
         Left            =   720
         MaxLength       =   30
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   9
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Kms:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Com&bustível"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.TextBox txQuant 
      Height          =   285
      Index           =   6
      Left            =   6780
      TabIndex        =   38
      Top             =   4860
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   3075
      MaxLength       =   41
      TabIndex        =   37
      Top             =   4860
      Width           =   3720
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   7260
      TabIndex        =   39
      Top             =   4860
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   7260
      TabIndex        =   35
      Top             =   4620
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   7260
      TabIndex        =   31
      Top             =   4380
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   7260
      TabIndex        =   27
      Top             =   4140
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   7260
      TabIndex        =   23
      Top             =   3900
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   7260
      TabIndex        =   19
      Top             =   3660
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   7260
      TabIndex        =   15
      Top             =   3420
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3075
      MaxLength       =   41
      TabIndex        =   33
      Top             =   4620
      Width           =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3075
      MaxLength       =   41
      TabIndex        =   29
      Top             =   4380
      Width           =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3075
      MaxLength       =   41
      TabIndex        =   25
      Top             =   4140
      Width           =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3075
      MaxLength       =   41
      TabIndex        =   21
      Top             =   3900
      Width           =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3075
      MaxLength       =   41
      TabIndex        =   17
      Top             =   3660
      Width           =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3075
      MaxLength       =   41
      TabIndex        =   13
      Top             =   3420
      Width           =   3720
   End
   Begin VB.CommandButton Botao 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Index           =   2
      Left            =   7080
      TabIndex        =   44
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Botao 
      Caption         =   "&Imprimir Orçamento"
      Height          =   315
      Index           =   1
      Left            =   1500
      TabIndex        =   43
      Top             =   7320
      Width           =   1635
   End
   Begin VB.CommandButton Botao 
      Caption         =   "&Gravar"
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   42
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txQuant 
      Height          =   285
      Index           =   5
      Left            =   6780
      TabIndex        =   34
      Top             =   4620
      Width           =   495
   End
   Begin VB.TextBox txQuant 
      Height          =   285
      Index           =   4
      Left            =   6780
      TabIndex        =   30
      Top             =   4380
      Width           =   495
   End
   Begin VB.TextBox txQuant 
      Height          =   285
      Index           =   3
      Left            =   6780
      TabIndex        =   26
      Top             =   4140
      Width           =   495
   End
   Begin VB.TextBox txQuant 
      Height          =   285
      Index           =   2
      Left            =   6780
      TabIndex        =   22
      Top             =   3900
      Width           =   495
   End
   Begin VB.TextBox txQuant 
      Height          =   285
      Index           =   1
      Left            =   6780
      TabIndex        =   18
      Top             =   3660
      Width           =   495
   End
   Begin VB.TextBox txQuant 
      Height          =   285
      Index           =   0
      Left            =   6780
      TabIndex        =   14
      Top             =   3420
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   1920
      Left            =   120
      TabIndex        =   48
      Top             =   3240
      Width           =   255
   End
   Begin MSMask.MaskEdBox MsData 
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Format          =   "DD/MM/YYYY"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   1395
      Left            =   840
      TabIndex        =   46
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   1
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cbCliente 
         Height          =   315
         ItemData        =   "ORC.frx":60B5
         Left            =   960
         List            =   "ORC.frx":60B7
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   2
         Left            =   3720
         MaxLength       =   30
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cbCarro 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text 
         DataField       =   "Ender"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "&Nome:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "&Carro:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NT"
      Height          =   195
      Index           =   3
      Left            =   2700
      TabIndex        =   92
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label lbCapTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   7320
      TabIndex        =   81
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbCapTotal 
      AutoSize        =   -1  'True
      Caption         =   "Falta"
      Height          =   195
      Index           =   3
      Left            =   6720
      TabIndex        =   80
      Top             =   5700
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lbCapTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   7320
      TabIndex        =   79
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbCapTotal 
      AutoSize        =   -1  'True
      Caption         =   "Pago"
      Height          =   195
      Index           =   1
      Left            =   6720
      TabIndex        =   78
      Top             =   5460
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbCapTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Index           =   0
      Left            =   6720
      TabIndex        =   77
      Top             =   5220
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lbVendedor 
      Height          =   255
      Left            =   6660
      TabIndex        =   69
      Top             =   6000
      Width           =   1755
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "meses"
      Height          =   195
      Index           =   1
      Left            =   3420
      TabIndex        =   68
      Top             =   5220
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Garantia: "
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   67
      Top             =   5220
      Width           =   690
   End
   Begin VB.Label lbGarantia 
      AutoSize        =   -1  'True
      Caption         =   "Garantia Até: "
      Height          =   195
      Left            =   1140
      TabIndex        =   56
      Top             =   5820
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Pagamento:"
      Height          =   195
      Index           =   13
      Left            =   4380
      TabIndex        =   54
      Top             =   6060
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "&Data"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lbTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7320
      TabIndex        =   52
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Index           =   11
      Left            =   7680
      TabIndex        =   51
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quant."
      Height          =   195
      Index           =   10
      Left            =   6840
      TabIndex        =   50
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Peça"
      Height          =   195
      Index           =   9
      Left            =   3120
      TabIndex        =   49
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lbAtraso 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "Cliente em Atraso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2662
      TabIndex        =   90
      Top             =   5400
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Menu MnuDesfaz 
      Caption         =   "Desfazer Pagamento"
      Visible         =   0   'False
   End
   Begin VB.Menu MnuItens 
      Caption         =   "Itens do Orçamento"
      Visible         =   0   'False
   End
   Begin VB.Menu MnuImprGraf 
      Caption         =   "Impressão Gráfica"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.9 Conserto da situação que não dava pra colocar o mecânico nas tarefas
'5.0.4 Correções de bugs
'5.0.3 Retirada da previsão para orçamento duplicado
'5.0.2 Utilizar a escolha da impressão somente se a impressora for definida como USB
'5.0.1 Escolha da impressoda na impressão do Orçamento, pela tela
'5.0.0 Ajuste na impressão em fita
'5.0.0 Ajuste na impressão em folha inteira
'5.0.0 Conserto da visualização dos valores dos itens quando tem mais que a quantidade visivel
'5.0.0 Apagar o restante da linha ao deletar a ultima linha dos itens
'5.0.0 Aumento de itens do orçamento de 200 para 300
'4.9.3 Ajuste na impressão USB
'4.9.1 Ajuste da Impressão em Fita via Usb
'4.8.9 Ajustes para base zerada
'4.7.4 Correção da visualização dos preços errada só na primeira operação
'4.7.3 Ajuste da retirada da visualização de valores para o modo mecânico
'4.7.2 Retirar a visualização de valores para o modo mecânico
'4.6.8 Impressão da observação do cliente na impressão do orçamento
'4.6.7 Bloquear a edição dos itens do orçamento para cliente em atraso, agora também para o perfil de escritório
'4.5.8 Mostrar a gravação automática
'4.5.6 Não permitir novo orçamento para cliente em atraso em modo balcão
'4.5.5 Em modo Escritório pode editar cliente em atraso
'4.5.4 Cliente em atraso
'4.3.9 Previsão do erro EsseCont = 1
'4.3.9 Ajuste no número do orçamento na impressão em folha inteira
'4.3.2 Permitir concluir as tarefas no pagamento
'4.2.7 Inscrição Estatual
'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
'4.2.3 Mais informação na impressão da observação
'4.2.2 Ajuste na Impressão da observação
'4.2.1 Impressão da observação
'4.1.2 Conferência adicional para ver se não há itens duplicados
'4.1.1 Conserto da alteração dos itens
'4.1.0 Conserto da impressão do orçamento, depois da versão 4.0.5
'4.1.0 Correção da deleção de peças quando tem mais de 7 itens
'4.0.9 Correção da deleção de peças[2]
'4.0.7 Ajuste na reimplantação da deleção dos itens do orçamento
'4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
'4.0.5 Formatação de ano com dois caracteres na visualização da data do item no orçamento
'4.0.5 Impedir duplicações de itens no novo orçamento
'4.0.5 Impedir gravação automática sem a identificação do balconista primeiro
'4.0.4 Impedir de pedir novamente o operador depois de ter pedido antes
'4.0.4 Previsão para orçamento duplicado ao gravar

Option Explicit

'Arquivos
'Private TbCli As Recordset
'Private tbCAR As Recordset
Private Orca  As Recordset

'Auxiliares
Private TemItens     As Boolean         'Indica que num orcamento lido tem itens para serem mostrados
Private FezActivate  As Boolean
Private Backspace    As Boolean
Private FAlterado    As Boolean
Private MudandoCombo As Boolean
Private OrcDoCLi     As Boolean
Private lcImprEmBra  As Boolean
Private Gravando     As Boolean
Private Mudando      As Boolean
Private AQuemPagou   As Integer
Private pIndex       As Integer
Private ItensDig     As Integer
Private l_NrOrcamento As Integer
Private UltTot       As Currency
Private TotalGrid    As Currency
Private OrcaMecanica As Double
Private GPlaca       As String * 8
Private strcbCliente As String
Private strcbCarro   As String
Private strcbItem    As String
Private PAGO         As Date

'itens do orcamento
Private Linha             As Integer
Private AltScrol          As Boolean

'5.0.0 Aumento de itens do orçamento de 200 para 300
Private Const MaxItensOrcs = 300
'Private Const MaxItensOrcs = 200

'2.5.3 Aumento dos itens de orçamento de 100 para 200
Private ITO(MaxItensOrcs) As ItOrc            '1.6 Aumento do quantidade nos orçamentos

'Margens da impressão
Private MargemH As Integer
Private MargemV As Integer
Private posY As Integer

'Objeto da impressão
Private Objeto As Object

Private GravaAutom As Boolean
Private PriTotal As Currency
Private Falta As Currency
Private Balconista   As Integer
Private ItemAlterado As Boolean

'4.0.5 Impedir duplicações de itens no novo orçamento
Private OrcNovo As Boolean

'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
Private Data As Date

'4.7.2 Retirar a visualização de valores para o modo mecânico
Private tx2Vi        As Boolean

'5.0.4 Correções de bugs
Private FazerPagar  As Boolean

Private Sub ImprAlinhada(Texto As String)
Dim Tamanho As Single
Dim Diferenca As Single

Const Default = 18.521
Const TamR = 9
Const Conversao = 34.0154419307813

Tamanho = TextWidth(Texto) / Conversao
Diferenca = Default - Tamanho

Objeto.CurrentX = MargemH + 155
Objeto.CurrentY = posY
Objeto.Print "R$:"

Objeto.CurrentY = posY
Objeto.CurrentX = MargemH + 155 + TamR + Diferenca
Objeto.Print Texto

End Sub

Sub Imprimir()
Dim a      As Integer
Dim MaObra As Currency
Dim Cont%
Dim Aux1   As String
Dim Aux2   As String

Const TmObs = 10
Const TmTitObs = 14
Const TmNormal = 14
Const TmTit = 19

frmOrc.Caption = "Impressão"

'Preview
If InStr(UCase(LinhaDeComando), "PREVIEW") > 0 Then
   Set Objeto = Preview
   Objeto.Show
Else
   Set Objeto = Printer
End If

MargemV = 0
MargemH = 15

Objeto.ScaleMode = 6   'Milimetros
Objeto.Font = "Arial" ' Narrow"
Objeto.CurrentY = MargemV + 0
Objeto.CurrentX = MargemH + 75
Objeto.FontSize = TmTit
Objeto.FontBold = True
On Error GoTo PauDaImpressora
Objeto.Print "ORÇAMENTO"
On Error GoTo 0

Objeto.FontSize = TmNormal
Objeto.FontBold = False

'4.3.9 Ajuste no número do orçamento na impressão em folha inteira
Objeto.CurrentX = MargemH + 130

Objeto.Print "Numero da Nota: " & NrOrcamento

Objeto.CurrentY = MargemV + 8
Objeto.CurrentX = MargemH + 25
Objeto.Print "Porto Alegre, " + Format$(CDate(MsData.Text), "dd") + " de " + Format$(CDate(MsData.Text), "mmmm") + " de " + Format$(CDate(MsData.Text), "yyyy")

Objeto.CurrentY = MargemV + 19
Objeto.CurrentX = MargemH + 0
Objeto.Print "Proprietário do Veículo: " + cbCliente.Text

Objeto.CurrentX = MargemH + 0
Objeto.Print "Endereço: " + Text(0).Text

Objeto.CurrentX = MargemH + 0
Objeto.Print "Telefone: " + clsCLi.Telefone

Objeto.CurrentX = MargemH + 0
posY = Objeto.CurrentY
Objeto.Print "Veículo Marca: " + cbCarro.Text + " " + Text(1).Text

Objeto.CurrentX = MargemH + 0
Objeto.Print "Placa " + Text(2).Text

If GridT.ApproxCount > 0 Then
    GridT.MoveFirst
    For a = 0 To (GridT.ApproxCount - 1)
        Objeto.CurrentX = MargemH + 0
        posY = Objeto.CurrentY
        Objeto.Print "Mão de Obra " & GridT.Columns(1).Value & ": "
        ImprAlinhada Format(GridT.Columns(2).Value, "#,###,##0.00")
        
        '5.0.0 Ajuste na impressão em folha inteira
        MaObra = MaObra + IIf((GridT.Columns(2).Value = ""), 0, GridT.Columns(2).Value)
        'MaObra = MaObra + GridT.Columns(2).Value = ""
        
        GridT.MoveNext
    Next
End If
Objeto.Print

Objeto.CurrentX = MargemH + 0
posY = Objeto.CurrentY
Objeto.Print "Soma da Mão de Obra:"
ImprAlinhada Format$(MaObra, "###,###.00")
Objeto.CurrentX = MargemH + 0
posY = Objeto.CurrentY
Objeto.Print "Combustível: " + Text(9).Text
Objeto.CurrentX = MargemH + 66
Objeto.CurrentY = posY
Objeto.Print "Kms: " + Text(10).Text
Itens:
Objeto.Print

Objeto.CurrentX = MargemH + 0
posY = Objeto.CurrentY
Objeto.Print "Peças"
Objeto.CurrentX = MargemH + 135
Objeto.CurrentY = posY
Objeto.Print "Quant."
Objeto.CurrentX = MargemH + 155
Objeto.CurrentY = posY
Objeto.Print "Valor"

Dim EsseCont As Integer

For Cont% = 0 To ItensDig
    
    '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
    EsseCont = NrIt(Cont%)
    
    '4.3.9 Previsão do erro EsseCont = 1
    If EsseCont > -1 Then
    
        If ITO(EsseCont).Quant > 0 Then
            Objeto.CurrentX = MargemH + 0
            posY = Objeto.CurrentY
            Objeto.FontSize = 15
            Objeto.Print ITO(EsseCont).Peca
            Objeto.CurrentX = MargemH + 140
            Objeto.CurrentY = posY
            Objeto.FontSize = TmNormal
            Objeto.Print ITO(EsseCont).Quant
            ImprAlinhada Format$(ITO(EsseCont).Valor, "###,###.00")
    '    If ITO(Cont%).Quant > 0 Then
    '        Objeto.CurrentX = MargemH + 0
    '        posY = Objeto.CurrentY
    '        Objeto.FontSize = 15
    '        Objeto.Print ITO(Cont%).Peca
    '        Objeto.CurrentX = MargemH + 140
    '        Objeto.CurrentY = posY
    '        Objeto.FontSize = TmNormal
    '        Objeto.Print ITO(Cont%).Quant
    '        ImprAlinhada Format$(ITO(Cont%).Valor, "###,###.00")
        
        End If
    End If
Next
Objeto.Print
posY = Objeto.CurrentY
Objeto.CurrentX = MargemH
Objeto.Print "TOTAL"

'3.7.3 Log grava o conteudo do total do orçamento
ImprAlinhada lbTotal.Caption
'ImprAlinhada Label2.Caption

Objeto.FontSize = TmTitObs
Objeto.Print
Objeto.CurrentX = MargemH
Objeto.Print "OBSERVAÇÃO"
Objeto.FontSize = TmObs
a = InStr(Text(8), Chr(13))
If a > 0 Then
   Aux1 = Text(8)
   Do While Len(Aux1) > 0
      a = InStr(Aux1, Chr(13))
      If a > 0 Then
         Aux2 = Left(Aux1, InStr(Aux1, Chr(13)) - 1)
         Aux1 = Mid(Aux1, Len(Aux2) + 3)
      Else
         Aux2 = Aux1
         Aux1 = ""
      End If
      Objeto.CurrentX = MargemH
      Objeto.Print Aux2
   Loop
Else
   Objeto.CurrentX = MargemH
   Objeto.Print Text(8).Text
End If

'4.6.8 Impressão da observação do cliente na impressão do orçamento
If IsNull(Orca!Observacao) = False Then
    Objeto.FontSize = TmTitObs
    Objeto.Print
    Objeto.CurrentX = MargemH
    Objeto.Print "OBSERVAÇÃO DO CLIENTE"
    a = InStr(Orca!Observacao, Chr(13))
    If a > 0 Then
       Aux1 = Orca!Observacao
       Do While Len(Aux1) > 0
          a = InStr(Aux1, Chr(13))
          If a > 0 Then
             Aux2 = Left(Aux1, InStr(Aux1, Chr(13)) - 1)
             Aux1 = Mid(Aux1, Len(Aux2) + 3)
          Else
             Aux2 = Aux1
             Aux1 = ""
          End If
          Objeto.CurrentX = MargemH
          Objeto.Print Aux2
       Loop
    Else
       Objeto.CurrentX = MargemH
       Objeto.Print Text(8).Text
    End If
End If

frmOrc.Caption = "Orçamento"
Botao(2).Caption = "Fechar"
If InStr(UCase(LinhaDeComando), "PREVIEW") = 0 Then
    Objeto.EndDoc
    Botao(2).SetFocus
End If
Exit Sub

PauDaImpressora:
msgboxL "Verifique a situação e tente novamente", vbExclamation, "Erro na impressora"
End Sub

Private Function VeOpcaoImprEBranco(C As Control) As Boolean
If lcImprEmBra Then
    VeOpcaoImprEBranco = True
Else
    VeOpcaoImprEBranco = Valo(C)
End If
End Function

Private Sub Total(Index As Integer)
Dim Cont    As Integer
Dim lcTotal As Currency

If Index > -1 Then
    If ItensDig < (Linha + Index) Then ItensDig = Linha + Index
End If
lcTotal = TotalGrid

Dim EsseCont As Integer

'For Cont = 0 To ItensDig
Do While Cont < (ItensDig + 1)
        
    '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
    'EsseCont = NrIt(Cont)
    Do While Not (ITO(Cont).Existente = False Or ITO(Cont).Deletar = False)
        Cont = Cont + 1
    Loop
    If Cont > -1 Then
        lcTotal = lcTotal + ITO(Cont).Valor
        'lcTotal = lcTotal + ITO(Cont).Valor
    Else
        Exit Do
    End If
    Cont = Cont + 1
    
Loop
UltTot = lcTotal

'3.7.3 Log grava o conteudo do total do orçamento
SetaLbTotal lcTotal, 0
'MostraValor Label2, lcTotal

If PAGO > 0 Then
    If PriTotal > 0 Then
        Falta = 0
        If UltTot <> PriTotal Then
            btPagamento.Caption = "Pagamento Parcial"
            btPagamento.Enabled = True
            MostraPagtoParc PriTotal, (UltTot - SN(Orca("VlrPago"), vbCurrency))
        End If
    End If
End If
End Sub

Private Sub MostraPagtoParc(Pato As Currency, lcFalta As Currency)
Dim a       As Integer
Dim Mostrar As Boolean

'3.9.2 Atualizar a situação de pagamento ao pagar um orçamento com pagto parcial
Mostrar = (lcFalta > 0)
For a = 0 To 4
    lbCapTotal(a).Visible = Mostrar
Next

'3.9.2 Atualizar a situação de pagamento ao pagar um orçamento com pagto parcial
MostraValor lbCapTotal(2), Pato
If Mostrar Then

    '5.0.4 Correções de bugs
    MostraValor lbCapTotal(4), Abs(PriTotal - Pato)
    'MostraValor lbCapTotal(4), (PriTotal - Pato)

    MostraValor lbCapTotal(2), Pato
End If
Falta = lcFalta
End Sub

'3.7.3 Log grava o conteudo do total do orçamento
Private Sub SetaLbTotal(Valor As Currency, Lugar As Integer)
Dim Aux As String

Select Case Lugar
    Case 0
        Aux = "Sub Total"
    Case 1
        Aux = "MostraOrcamento"
    Case 2
        Aux = "ExclusaoItens"
End Select
Loga "lbTotal = " & Valor & " apartir de " & Aux
MostraValor lbTotal, Valor
End Sub

Private Sub GravaItens()
Dim DeuCerto As Boolean
Dim Cont%, VezErro%
Dim ProxItem&
Dim SQL As String

584     SQL = "DELETE DISTINCTROW * "
586     SQL = SQL & "From Itens_Orc "
588     SQL = SQL & "WHERE Orçamento=" & NrOrcamento

'4.0.5 Impedir duplicações de itens no novo orçamento
If OrcNovo Then
    ExecSql SQL
Else

    Dim Aux As String
    For Cont% = 0 To ItensDig
            
        '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
        If ITO(Cont%).Existente Then
        
            If ITO(Cont%).Alterado Then
            
                If ITO(Cont%).ID > 0 Then
                    Aux = Aux & Trim(STR(ITO(Cont%).ID)) & ","
                End If
                
            Else
            
                '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
                If ITO(Cont%).Deletar Then
                    SQL = "Delete From Itens_Orc WHERE Orçamento=" & NrOrcamento
                    SQL = SQL & " and Cont = " & ITO(Cont%).ID
                    ExecSql SQL
                    ITO(Cont%).Existente = False
                End If
                
            End If
        
        End If
        
    Next
    If Len(Aux) > 1 Then
        Aux = Left(Aux, Len(Aux) - 1)
        SQL = SQL & " and Cont in (" & Aux & ")"
590     ExecSql SQL
    End If
End If
'590     ExecSql SQL

600     On Error GoTo NaoConhecoEsteErro
610     For Cont% = 0 To ItensDig
620         If ITO(Cont%).Peca > "" Then

                If ITO(Cont%).Alterado = True Then

                    DeuCerto = False
                    VezErro% = 0
                    Do While DeuCerto = False
                                    
630                         ProxItem& = Consulta("Select Max(Cont) From Itens_Orc") + 1
                        SQL = "Insert Into Itens_Orc Values (" & ProxItem& & ", "
                        SQL = SQL & NrOrcamento & ", '"
                        SQL = SQL & ITO(Cont%).Peca & "', "
                        SQL = SQL & ITO(Cont%).Quant & ", "
640                         SQL = SQL & VlrSql(STR(ITO(Cont%).Valor)) & ", "
645                         SQL = SQL & DTSqld(ITO(Cont%).Data)
    
                        SQL = SQL & ")"
650                         ExecSql SQL
                        SQL = "Select Count(*) From Itens_Orc Where Cont = " & ProxItem&
                        SQL = SQL & " and Orçamento = " & NrOrcamento
                        SQL = SQL & " and Item = '" & ITO(Cont%).Peca & "'"
660                         If Consulta(SQL) > 0 Then
                            DeuCerto = True
                        Else
                            VezErro% = VezErro% + 1
670                             Sleep (VezErro% * 100)
                        End If
                        If VezErro% > 9 Then
680                             msgboxL "Atenção houve um erro desconhecido ao gravar os itens do orçamento", vbCritical, "Contacte o programador"
                            End
                        End If
                        If ITO(Cont%).ID Then
                            SQL = "Update AuditItemOrc Set Item_Orc = " & ProxItem&
                            SQL = SQL & " Where Item_Orc = " & ITO(Cont%).ID
                            ExecSql SQL
                        End If
                        SQL = "Insert Into AuditItemOrc (Item_Orc, Balconista, Item, Quant, Valor, DtItOrc) Values ("
                        SQL = SQL & ProxItem& & ", "
                        SQL = SQL & Balconista & ", "
                        SQL = SQL & FA(ITO(Cont%).Peca) & ", "
                        SQL = SQL & ITO(Cont%).Quant & ", "
                        SQL = SQL & VlrSql(STR(ITO(Cont%).Valor)) & ", "
                        SQL = SQL & DTSqld(Now) & ")"
                        ExecSql SQL

                    Loop
                End If
700         End If
710     Next

'4.1.2 Conferência adicional para ver se não há itens duplicados
If OrcNovo = False Then
    VerificaDuplicados
End If

Exit Sub

NaoConhecoEsteErro:
Loga "Erro: " & Error(Err) & " na linha " & Erl, lERR
970 msgboxL "Erro nº" & STR(Err) & " - " & Error$(Err) & ", guarde este orçamento para mostrar ao programador" & Chr$(10) & "Tentarei continuar", vbExclamation, "OrCarro - Erro"
980 Resume Next
End Sub

'4.1.2 Conferência adicional para ver se não há itens duplicados
Private Sub VerificaDuplicados()
Dim OK        As Boolean
Dim SQL       As String
Dim rsConfere As Recordset
Dim rsEncontr As Recordset

Do While OK = False
    SQL = "SELECT Count(*) AS Expr1, Item, DtItOrc "
    SQL = SQL & "from Itens_Orc "
    SQL = SQL & "Where Orçamento = " & NrOrcamento
    SQL = SQL & " GROUP BY Orçamento, Item, DtItOrc "
    SQL = SQL & "Having  Count(*) > 1 "
    AbreTB rsConfere, SQL, dbOpenDynaset
    If rsConfere.EOF Then
        OK = True
    Else
        Loga "Encontrado Item Duplicado " & FA(rsConfere!Item), lDBG
        SQL = "SELECT Cont, Valor "
        SQL = SQL & "from Itens_Orc "
        SQL = SQL & "Where Item = " & FA(rsConfere!Item)
        SQL = SQL & " And DtItOrc = " & DTSqld(rsConfere!DtItOrc) '#2/24/2015 8:55:20 AM#) "
        SQL = SQL & " And Orçamento = " & NrOrcamento
        SQL = SQL & " ORDER BY Cont "
        AbreTB rsEncontr, SQL, dbOpenDynaset
        Loga "Apagado item com valor " & STR(rsEncontr!Valor), lDBG
        ExecSql "Delete From Itens_Orc Where Cont = " & rsEncontr!Cont
        rsEncontr.MoveNext
        Loga "Mantido o com valor " & STR(rsEncontr!Valor), lDBG
        rsConfere.Close
        rsEncontr.Close
    End If
Loop
End Sub

Private Sub Gravar(Automatico As Boolean)
Dim Adicao          As Boolean
Dim TemProdTarefas  As Boolean
Dim SQL             As String
Dim Orc             As Recordset
Dim TBItens         As Recordset
Dim rsOrcAnt        As Recordset
Dim Orcam           As Long
Dim X As Long
Dim ProxItem&

On Local Error GoTo ErroParteDeCima
Gravando = True
Loga "frmOrc.Gravar", lDBG
Loga "Cliente: " & cbCliente.Text, lDBG
Loga "Carro: " & Text(2).Text, lDBG
Loga "QuantItens: " & ItensDig, lDBG

'4.0.4 Previsão para orçamento duplicado ao gravar
Inicio_Gravar:

110 If NrOrcamento > 0 Then
265     AbreTB Orc, "Select * From Orcamento Where Orcamento = " & NrOrcamento, dbOpenDynaset

X = DoEvents()
Sleep 100

330         Orc.Edit
340         Orc("Data") = DateValue(MsData.Text)
360     Orc("Cliente") = cbCliente.Text
430     If Text(8).Text = "" Then
440         Orc("Obs") = " "
450     Else
460         Orc("Obs") = Text(8).Text
470     End If

'3.7.3 Log grava o conteudo do total do orçamento
480     Orc("Total") = Val(VlrSql(lbTotal.Caption))
'480     Orc("Total") = Val(VlrSql(Label2.Caption))

490     Orc("Kms") = Text(9).Text
500     Orc("Kilom") = Text(10).Text
    Orc("Garantia") = Val(Text(11).Text)
    If PAGO > 0 And lbGarantia.Visible = False Then
        Orc("Pagamento") = PAGO
    End If
    If Text(12).Text = "" Then
        Orc("ObsMec") = " "
    Else
        Orc("ObsMec") = Text(12).Text
    End If
520     Orc.Update
    If Text(13).Text > "" Then
        SQL = "Update Carros Set Historico = '" & Text(13).Text & "' "
    Else
        SQL = "Update Carros Set Historico = null "
    End If
    SQL = SQL & " Where Placa = '" & Orc("Carro") & "'"
    ExecSql SQL
    
530     On Local Error GoTo OutroErro
540     Orc.Close
    Orcam = NrOrcamento
    
    Adicao = False

Else
    Adicao = True
    'Orçamento Novo
    Orcam = Consulta("Select Max(Orcamento) From Orcamento") + 1
    
    '4.8.9 Ajustes para base zerada
    If Orcam > 1 Then
        
        AbreTB rsOrcAnt, "Select Cliente, Carro, Kilom From Orcamento Where Orcamento = " & (Orcam - 1)
        
        '5.0.3 Retirada da previsão para orçamento duplicado
        '4.0.4 Previsão para orçamento duplicado ao gravar
'        Loga "Orcamento Anterior " & rsOrcAnt("Cliente") & " " & rsOrcAnt("Carro") & " " & rsOrcAnt("Kilom")
'        If rsOrcAnt("Cliente") = Trim(cbCliente.Text) And rsOrcAnt("Carro") = Trim(Text(2).Text) And rsOrcAnt("Kilom") = Text(10).Text Then
'            NrOrcamento = Orcam - 1
'            Loga "Detectado Orcamento duplicado"
'            GoTo Inicio_Gravar
'        End If
    
    End If
    
    SQL = "Insert Into Orcamento (Orcamento , Data, Cliente, Carro, Obs, Kms, Kilom, Total, Garantia, Pagamento, ObsMec, Vend, VlrPago) Values ( " & Orcam & ", "
    'SQL = "Insert Into Orcamento Values ( " & Orcam & ", "
    
    '4.2.3 Mais informação na impressão da observação
    SQL = SQL & DTSqls(Format(Now, "DD/MM/YYYY HH:MM")) & ", '"
    'SQL = SQL & DTSqls(MsData.Text) & ", '"
    
    SQL = SQL & Trim(cbCliente.Text) & "', '"
    SQL = SQL & Text(2).Text & "', "            'Carro
        
    '3.0.8 Aceitar ponto na kilometragem
    SQL = SQL & "'" & Text(8).Text & " ', '"     'Obs
    SQL = SQL & Text(9).Text & "' , '"           'Kms
    SQL = SQL & Text(10).Text & "' , "          'Kilom
    'SQL = SQL & "'" & Text(8).Text & " ', "     'Obs
'    SQL = SQL & Valo(Text(9)) & " , "           'Kms
'    SQL = SQL & Valo(Text(10)) & " , "          'Kilom
    
    '3.7.3 Log grava o conteudo do total do orçamento
    SQL = SQL & VlrSql(lbTotal.Caption) & " , "  'Total
    'SQL = SQL & VlrSql(Label2.Caption) & " , "  'Total
    
    SQL = SQL & Valo(Text(11)) & " , "          'Garantia
    
    'Pagamento      Data/Hora
    If PAGO > 0 Then
        SQL = SQL & DTSqls(Format(PAGO, "dd/mm/yyyy"))
    Else
        SQL = SQL & "0"
    End If
    SQL = SQL & ",'" & Text(12).Text & "'"
    
    '3.0.1 Nova forma da gravação do operador que recebeu
    If AQuemPagou > 0 Then
    
        SQL = SQL & ",0 "
    Else
        '3.0.1 Nova forma da gravação do operador que recebeu
        SQL = SQL & "," & AQuemPagou
        'SQL = SQL & "," & SN(Consulta("Select Codi From Mecanicos Where Nome = '" & lbVendedor.Caption & "'"), vbInteger)
    End If
    
    SQL = SQL & ",0 "
    SQL = SQL & ")"
    
    ExecSql SQL
    If Text(13).Text > "" Then
        SQL = "Update Carros Set Historico = '" & Text(13).Text & "' "
    Else
        SQL = "Update Carros Set Historico = null "
    End If
    SQL = SQL & " Where Placa = '" & Text(2).Text & "'"
    ExecSql SQL
    NrOrcamento = Orcam
End If
550 If ItensDig > 0 Then
If ItemAlterado Then
    GravaItens
End If
730     On Error GoTo OutroErro
End If

Dim ResultTarfas As Integer
ResultTarfas = GravaTarefas(Orcam)

If ResultTarfas > 0 Then
    FezActivate = True
    If ResultTarfas > 1 Then
        If Automatico Then
            TemProdTarefas = True
        Else
            MsgBox "Há lançamentos de taréfas sem ter sido finalizada a gravação", vbCritical, "Gravação"
        End If
    End If
    If NrOrcamento > 0 Then
        If Automatico = False Then
            frmOrc.SetFocus
        End If
    End If
    If TemProdTarefas = True Then
        Frame4.BackColor = vbRed
    Else
        Frame4.BackColor = vbBlack
    End If
    Exit Sub
End If
750     Botao(0).Enabled = False
760     Botao(1).Enabled = True
770     Botao(2).Caption = "Fechar"

If Automatico = False Then

780     Botao(1).SetFocus
790     If NrOrcamento > 0 Then
810         frmOrc.SetFocus
820     End If
End If
Loga "Gravação concluída", lDBG

    Gravando = False

830 Exit Sub

ErroParteDeCima:
Loga "Erro: " & Error(Err) & " na linha " & Erl, lERR
msgboxL "Erro de tipo " & Error(Err) & " na linha " & Erl
Resume Next

OutroErro:
Loga "Erro: " & Error(Err) & " na linha " & Erl, lERR
990 msgboxL "Erro de tipo " & Error(Err) & " na linha " & Erl
End
End Sub

Private Sub Scroll()
Dim Cont%
Dim EsseCont As Integer

For Cont% = 0 To 6

    '2.5.3 Aumento dos itens de orçamento de 100 para 200
    If (Linha + Cont%) > MaxItensOrcs Then
        msgboxL "A quantidade de itens do orçamento esta acima do previsto", vbCritical, "Impossível continuar, contacte o progrador"
        Exit Sub
        
    Else
    
        '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
        EsseCont = NrIt(Cont%)
        If EsseCont > -1 Then
        
            Text1(Cont%).Text = ITO(EsseCont).Peca
            txQuant(Cont%).Text = STR$(ITO(EsseCont).Quant)
                        
            '4.7.2 Retirar a visualização de valores para o modo mecânico
            If tx2Vi Then
                MostraValor Text2(Cont%), ITO(EsseCont).Valor
            End If
            
            txData(Cont%).Text = Format(ITO(EsseCont).Data, "DD/MM/YY")
            txNome(Cont%).Text = ITO(EsseCont).Nome
        Else
            Text1(Cont%).Text = ""
            txQuant(Cont%).Text = ""
            Text2(Cont%).Text = ""
            txData(Cont%).Text = ""
            txNome(Cont%).Text = ""
        End If
    End If
Next
AltScrol = True
VScroll1.Value = Linha
AltScrol = False
End Sub

Private Sub Botao_Click(Index As Integer)
Dim Result As Boolean

Select Case Index
    Case Is = 0
    
        '2.0.3 Ajuste no funcionamento das permissões
        If Permissao = False Then
            If NrOrcamento > 100 Then
                msgboxL "Em modo de demonstração só é permitido 100 orçamentos"
                Botao(0).Enabled = False
                Exit Sub
            Else
                If Consulta("Select Count(*) From Orcamento") > 100 Then
                    msgboxL "Em modo de demonstração só é permitido 100 orçamentos"
                    Botao(0).Enabled = False
                    Exit Sub
                End If
            End If
        End If
        Dim PedirSenha As Boolean
        Dim Func       As Boolean
        Func = Consulta("Select Funcionario From Clientes Where NrCli = " & clsCLi.NrCli)
        PedirSenha = Func
        If PedirSenha = False Then
            PedirSenha = (Text1(0).Text > "")
        End If
        If PedirSenha Then
        
            '4.0.4 Impedir de pedir novamente o operador depois de ter pedido antes
            If Balconista = 0 Then
            
                Dim sNome As String
                Balconista = OperadorQueGravou(sNome, Func)
                If Balconista = 0 Then Exit Sub
            End If
        End If
        
        '4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
        Data = Now
        
        Gravar False
    Case Is = 1
    
        '5.0.2 Utilizar a escolha da impressão somente se a impressora for definida como USB
        If INI.TpImpress = 2 Then
        
            '5.0.1 Escolha da impressoda na impressão do Orçamento, pela tela
            Load frmConfigImpr
            frmConfigImpr.Caption = "Escolha a impressora"
            frmConfigImpr.Label2.Caption = ""
            frmConfigImpr.Show 1
            If frmConfigImpr.OK Then
                If frmConfigImpr.ImprFita Then
                    ImprimeFita
                Else
                    Imprimir
                End If
            End If
            Unload frmConfigImpr
            
        Else
            If INI.TpImpress = 0 Then
                Imprimir
            Else
                ImprimeFita
            End If
        End If
        
        Botao(2).Caption = "Fechar"
    Case Is = 2
        Unload Me
        
    '4.2.1 Impressão da observação
    Case Is = 3
        Botao(3).Enabled = False
                
        '4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
        AcionaImprObs
        '4.2.3 Mais informação na impressão da observação
        '4.2.2 Ajuste na Impressão da observação
        'ImpressaoObs cbCarro.Text, Text(1).Text, Text(2).Text, Text(12).Text
        
        Botao(3).Enabled = True
    
End Select
End Sub

'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
Private Sub AcionaImprObs()
Dim NrVend     As Integer
Dim NmVendedor As String

NmVendedor = Consulta("SELECT Nome from Mecanicos WHERE codi = " & Balconista)
ImpressaoObs cbCarro.Text, Text(1).Text, Text(2).Text, Text(12).Text, cbCliente.Text, Data, NmVendedor
End Sub

Private Sub btPagamento_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Resp       As String
Dim SQL        As String
Dim FazerPagar As Boolean

FazerPagar = False
If INI.UtComissoes = 0 Then
    PAGO = Now()
    AQuemPagou = 0
Else
    If TarefasEstaoConcluidas() = False Then
        
        '4.3.2 Permitir concluir as tarefas no pagamento
        If MsgBox("Deseja concluí-las agora?", vbExclamation + vbYesNo + vbDefaultButton2, "Existem tarefas não concluídas") = vbNo Then
            Exit Sub
        Else
            FazerPagar = True
        End If
'        MsgBox "Não será possível receber o pagamento pois existe tarefa não concluída", vbExclamation, "Impossível receber"
'        Exit Sub

    End If
    
    Dim VlrPago As Currency
    ReceberPagamento AQuemPagou, VlrPago
    If AQuemPagou > 0 Then
        PAGO = Now()
        GravaItens
        MostraPagtoParc VlrPago, (UltTot - VlrPago)
    End If
End If
If PAGO > 0 Then
    PAGO = Now
    If VlrPago = Falta Then
        btPagamento.Caption = "Pago"
    Else
        btPagamento.Caption = "Pagamento Parcial"
    End If
    If l_NrOrcamento > 0 Then
        SQL = "Update Orcamento Set Pagamento = " & DTSqls(Format(Now, "DD/MM/YYYY"))
        SQL = SQL & ", Vend = " & AQuemPagou
        SQL = SQL & ", Total = " & VlrSql(Trim(STR(UltTot)))
        SQL = SQL & ", VlrPago = " & VlrSql(Trim(STR(VlrPago)))
        SQL = SQL & " WHERE Orcamento.Orcamento=" & l_NrOrcamento
        ExecSql SQL
        
        '4.5.5 Em modo Escritório pode editar cliente em atraso
        If clsCLi.EmAtraso Then
            ReConfereAtraso
        End If
        
    End If
    pIndex = 0
    Alterado = True
    Botao(0).Enabled = False
    Botao(2).Caption = "Fechar"
    
    '4.3.2 Permitir concluir as tarefas no pagamento
    If FazerPagar Then
        ExecSql "Update Tarefas Set Situacao = 3 Where Orc = " & NrOrcamento
        DataT.Enabled = False
        CarregaTarefas NrOrcamento
        DataT.Enabled = True
        DataT.Refresh
    End If
    
End If
End Sub


'4.5.5 Em modo Escritório pode editar cliente em atraso
Private Sub ReConfereAtraso()
Dim EmAtraso As Boolean
Dim SQL      As String
Dim rsAux    As Recordset

SQL = "SELECT Cliente, Sum(Total) AS SomaDeTotal, Sum(VlrPago) AS SomaDeVlrPago "
SQL = SQL & "from Orcamento "
SQL = SQL & "GROUP BY Orcamento.Cliente "
SQL = SQL & "Having Cliente = " & FA(cbCliente.Text)
AbreTB rsAux, SQL, dbOpenSnapshot

EmAtraso = (rsAux!SomaDeTotal > rsAux!SomaDeVlrPago)
If EmAtraso = False Then
    ExecSql "Update Clientes Set EmAtraso = 0 Where Nome = " & FA(cbCliente.Text)
    MostraEmAtraso False
End If
End Sub

Private Sub ReceberPagamento(ByRef QuemRecebeu As Integer, ByRef VlrPago As Currency)
Dim sValor As String

'5.0.4 Correções de bugs
If UltTot > 0 Then
    Gravar True
End If

Load FormaPagamento
FormaPagamento.Orcamento = NrOrcamento

If Falta > 0 Then
    FormaPagamento.Valor = Falta
    sValor = lbCapTotal(4)
    
Else
    FormaPagamento.Valor = UltTot
    sValor = lbTotal.Caption
End If
FormaPagamento.lbTotal.Caption = "Valor Total: " & sValor
FormaPagamento.txEntrada.Text = sValor
FormaPagamento.Show 1

QuemRecebeu = FormaPagamento.QuemPagou
VlrPago = FormaPagamento.VlrPago
Unload FormaPagamento
End Sub

Private Sub btPagamento_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
btPagamento.Value = (PAGO > 0)
End Sub

Private Sub cbCarro_Click()
If MudandoCombo = True Then Exit Sub
MudandoCombo = False
MostraCarro
Loga "Saiu do cbCarro_Click", lDBG
End Sub

Private Sub MostraCarro(Optional Placa As String = "", Optional NrCli As Integer = 0)
Dim Cor    As String
Dim sPlaca As String
Dim Carro  As String

Static UltMostrado As String
If Placa <> UltMostrado Or Text(1).Text = "" Then
    UltMostrado = Placa
    If NrCli = 0 Then
        If NrOrcamento > 0 Then sPlaca = Placa
        NrCli = NrOrcamento
    Else
        sPlaca = Placa
    End If
        
    clsCLi.CorPlaca cbCarro.ListIndex, Cor, sPlaca, Carro
    If Cor > "" And Carro > "" Then
        Text(1).Text = Cor
        Text(2).Text = sPlaca
        
        If NrCli Then
            MudandoCombo = True
            cbCarro.Text = Carro
            MudandoCombo = False
        End If
    Else
        If GCliente = "" Then
            msgboxL "O cliente não consta mais como proprietário do carro", vbInformation, "Atenção"
        Else
            If msgboxL("Deseja retornar o carro para este cliente", vbExclamation + vbDefaultButton2 + vbYesNo, "O cliente não consta mais como proprietário do carro") = vbYes Then
                ReajustaCarro
            Else
                Me.Enabled = False
            End If
        End If
    End If
End If
Loga "MostraCarro", lDBG
End Sub

Private Sub ReajustaCarro()
Dim SQL   As String

SQL = "SELECT count(*) "
SQL = SQL & "from Orcamento "
SQL = SQL & "Where Cliente = '" & GCliente & "'"
SQL = SQL & "and Carro not in ( "
SQL = SQL & "SELECT Carros.Placa "
SQL = SQL & "from Carros "
SQL = SQL & "INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli "
SQL = SQL & "WHERE Clientes.Nome = '" & GCliente & "')"
If Consulta(SQL) = 0 Then
    msgboxL "Não foi possível retornar o carro ao cliente"
Else
    SQL = "Update Orcamento "
    SQL = SQL & "Set Carro = '" & GGPlaca & "' "
    SQL = SQL & "Where Orcamento in ("
    SQL = SQL & "SELECT Orcamento "
    SQL = SQL & "from Orcamento "
    SQL = SQL & "Where Cliente = '" & GCliente & "'"
    SQL = SQL & "and Carro not in ("
    SQL = SQL & "SELECT Carros.Placa "
    SQL = SQL & "from Carros "
    SQL = SQL & "INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli "
    SQL = SQL & "WHERE Clientes.Nome = '" & GCliente & "'))"
    ExecSql SQL
    MostraCarro GGPlaca
End If
End Sub

Private Sub cbCarro_KeyPress(KeyAscii As Integer)
TrataCombo strcbCarro, cbCarro, KeyAscii
End Sub

Private Sub cbCarro_LostFocus()
'strcbCarro = ""
If MudandoCombo = True Then Exit Sub
MudandoCombo = False
MostraCarro
End Sub

Private Sub cbCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    strcbCliente = ""
Else
    TrataCombo strcbCliente, cbCliente, KeyAscii
End If
End Sub

Private Sub cbCliente_LostFocus()
strcbCliente = ""
If cbCliente.Text > "" Then
    MostraCarrosDoCLi
End If
End Sub

Private Sub MostraCarrosDoCLi(Optional Placa As String = "", Optional NrCli As Integer = 0, Optional Endereco As String = "")
If clsCLi.SetaClienteAtual(Trim(cbCliente.Text)) = 0 Then
    msgboxL "Informe o nome do cliente novamente", vbExclamation, "Cliente nao idenficado"
    cbCliente.ListIndex = -1
    cbCliente.SetFocus
    Exit Sub
End If
If Endereco > "" Then
    Text(0).Text = Endereco
Else
    Text(0).Text = clsCLi.Endereco
End If
clsCLi.CarregaCarros cbCarro
MudandoCombo = True
If cbCarro.ListCount = 0 Then
    cbCarro.ListIndex = -1
Else
    cbCarro.ListIndex = 0
End If
MudandoCombo = False
If Placa > "" Then
    MostraCarro Placa, NrCli
End If

'4.5.6 Não permitir novo orçamento para cliente em atraso em modo balcão
lbAtraso.Visible = clsCLi.EmAtraso

End Sub

Private Sub MOSTRAORCAMENTO()
Dim Hab      As Boolean
Dim Desab    As Boolean
Dim NrCli    As Integer
Dim Cont     As Integer
Dim SQL      As String
Dim Controle As Control
Dim rsAudit  As Recordset

Const Mult = 1.02

OrcNovo = False
NrCli = clsCLi.NrDoNome(RTrim(Orca("Cliente")))
MsData.Text = Format$(Orca("Data"), "dd/mm/yyyy")
cbCliente.Text = RTrim(Orca("Cliente"))
MostraCarrosDoCLi Orca("Carro"), NrCli
Text(8).Text = Orca("Obs")
If Not IsNull(Orca("Kms")) Then
    Text(9).Text = Orca("Kms")
End If
If Not IsNull(Orca("Kms")) Then
    Text(10).Text = Orca("Kilom")
    
    '5.0.4 Correções de bugs
    Text(10).Tag = Text(10).Text
    
End If
If Not IsNull(Orca("Total")) Then

    PriTotal = Orca("Total")

    '3.7.3 Log grava o conteudo do total do orçamento
    SetaLbTotal Orca("Total"), 1
    'MostraValor Label2, Orca("Total")
            
    '3.7.3 Log grava o conteudo do total do orçamento
    lbTotal.Tag = lbTotal.Caption
    TotalGrid = Consulta("SELECT Sum(Vlr), Orc from Tarefas GROUP BY Orc HAVING Orc = " & NrOrcamento)

End If
If Orca("Garantia") > 0 Then
    Text(11).Text = STR(Orca("Garantia"))
Else
    Text(11).Text = ""
End If
If TemItens Then
    ItensDig = 0
    If IsNull(Orca("Item")) = False Then
        Do While Orca.EOF = False
            ITO(ItensDig).Peca = Orca("Item")
            ITO(ItensDig).Quant = Orca("Quant")
            ITO(ItensDig).Valor = Orca("Valor")
            ITO(ItensDig).Data = Orca("DtItOrc")
            ITO(ItensDig).ID = Orca("Cont")
            ITO(ItensDig).Alterado = False
            SQL = "SELECT top 1 Nome "
            SQL = SQL & "from AuditItemOrc "
            SQL = SQL & "INNER JOIN Mecanicos ON AuditItemOrc.Balconista = Mecanicos.codi "
            SQL = SQL & "Where Item_Orc = " & Orca("Cont")
            SQL = SQL & " Order By DtItOrc desc"
            ITO(ItensDig).Nome = Consulta(SQL)
            ITO(ItensDig).Existente = True
            ITO(ItensDig).Deletar = False
            Orca.MoveNext
            ItensDig = ItensDig + 1
        Loop
        Orca.MoveFirst
    End If
    
    '4.7.4 Correção da visualização dos preços errada só na primeira operação
    'Scroll
    
    ItemAlterado = False
End If

'3.0.1 Nova forma da gravação do operador que recebeu
AQuemPagou = 0

PAGO = SN(Orca("Pagamento"), vbDate)

If PAGO > 0 Then
    
    If Orca("VlrPago") = Orca("Total") Then
        btPagamento.Caption = "Pago"
        btPagamento.Value = True
        btPagamento.Enabled = False
    Else
        btPagamento.Caption = "Pagamento Parcial"
        btPagamento.Enabled = True
        
        '5.0.4 Correções de bugs
        'Deve ser valor que foi pago, somatório do campo Valor em Parcelas
        Dim ValorQueFoiPago As Currency
        SQL = "SELECT Sum(Parcelas.Valor) AS SomaDeValor, Parcelas.Orc "
        SQL = SQL & "from Parcelas "
        SQL = SQL & "GROUP BY Parcelas.Orc "
        SQL = SQL & "HAVING Parcelas.Orc  = " & Orca("Orcamento")
        ValorQueFoiPago = Consulta(SQL)
        MostraPagtoParc Orca("VlrPago"), (Orca("Total") - Orca("VlrPago"))
        
    End If
    
    lbGarantia.Caption = DizGarantia(PAGO, Orca("Garantia"))
    lbGarantia.Visible = True
    
    '3.0.1 Nova forma da gravação do operador que recebeu
    AQuemPagou = SN(Orca("Vend"), vbInteger)

    If AQuemPagou > 0 Then
        lbVendedor.Caption = Consulta("Select Nome From Mecanicos Where Codi = " & AQuemPagou & " and Ativo = True ")
    End If
    MnuDesfaz.Visible = True

    frmOrc.Height = frmOrc.Height * Mult
        
    Hab = False
    
    '5.0.4 Correções de bugs
    If INI.OrcGrav = 0 Then
    
        '3.9.0 Se foi pago a mais de 30 dias não permitir editar com exessao do modo escritório
        If (Now - PAGO) > 30 Then
            Desab = True
        Else
            Desab = (Consulta("Select Count(*) From Parcelas Where Orc = " & NrOrcamento & " and NrParc > 0") > 0)
        End If
        
        If Desab = True Then
            GridT.AllowAddNew = False
            GridT.AllowDelete = False
            GridT.AllowUpdate = False
            For Cont = 0 To 6
                Text1(Cont).Locked = True
                Text2(Cont).Locked = True
                txQuant(Cont).Locked = True
            Next
            Text(8).Locked = True
            Text(12).Locked = True
        End If
    
    End If
                
Else
    MnuDesfaz.Visible = False
    btPagamento.Enabled = True
    PAGO = 0
    lbGarantia.Visible = False
    Hab = True
    VeTotalGrid
    
    '3.7.3 Log grava o conteudo do total do orçamento
    If lbTotal.Tag <> lbTotal.Caption Then
        ExecSql "update Orcamento Set Total = " & VlrSql(Trim(STR(UltTot))) & " Where Orcamento = " & NrOrcamento
    End If
        
    MnuItens.Visible = True
    
    If MnuItens.Visible Then
        frmOrc.Height = frmOrc.Height * Mult
    End If
   
'    If MnuItens.Visible = False Then
'        frmOrc.Height = Nr
'    Else
'        frmOrc.Height = Nr + 300
'    End If
End If
Frame1.Enabled = False

If INI.ModoOperacao = tpMecanico Then
    Hab = False
    Botao(0).Enabled = False
    Botao(1).Enabled = False
    Label1(11).Visible = False
    GridT.Columns(2).Visible = False
    btPagamento.Enabled = False
    
    '3.7.3 Log grava o conteudo do total do orçamento
    lbTotal.Visible = False
    'Label2.Visible = False
    
    Text(0).Locked = True
    Text(1).Locked = True
    Text(2).Locked = True
    Text(8).Locked = True
    Text(9).Locked = True
    Text(10).Locked = True
    Text(11).Locked = True
    Text(13).Locked = True
    Frame1.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    MnuDesfaz.Visible = False
    
    '4.7.3 Ajuste da retirada da visualização de valores para o modo mecânico
    tx2Vi = False
    'tx2Vi = True
    
Else

    '4.7.3 Ajuste da retirada da visualização de valores para o modo mecânico
    tx2Vi = True
    'tx2Vi = False
    
End If

'4.7.4 Correção da visualização dos preços errada só na primeira operação
Scroll

Frame3.Enabled = Hab

Text(12).Text = SN(Orca("ObsMec"), vbString)
Text(13).Text = SN(Consulta("Select Historico From Carros Where Placa = '" & Orca("Carro") & "'"), vbString)
Mudando = True
opObs(0).Value = True
opObs(1).Value = False
Text(8).Visible = True
Text(12).Visible = False
Mudando = False
If INI.ModoOperacao = tpBalcao Then
    If PAGO > 0 Then
        Text(8).Enabled = False
    End If
ElseIf INI.ModoOperacao = tpEscritorio Then
    If INI.TpImpress = 1 Then
        MnuImprGraf.Visible = True
    End If
End If

'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
Balconista = SN(Orca("Vend"), vbInteger)
Data = Orca("Data")

'4.5.4 Cliente em atraso
If clsCLi.EmAtraso Then
    MostraEmAtraso True
End If

VScroll1.Enabled = True
DataT.Enabled = True
DataT.Refresh
FezActivate = True

'5.0.4 Correções de bugs
Caption = "Orçamento " & Orca("Orcamento")
End Sub
'4.5.4 Cliente em atraso
Private Sub MostraEmAtraso(Optional Atraso As Boolean = True)
Dim a As Integer

lbAtraso.Visible = Atraso

'4.5.5 Em modo Escritório pode editar cliente em atraso
If INI.ModoOperacao = tpBalcao Then
    Frame4.Enabled = Not (Atraso)
    For a = 0 To 6
        Text1(a).Enabled = Not (Atraso)
        txQuant(a).Enabled = Not (Atraso)
        Text2(a).Enabled = Not (Atraso)
    Next
End If
End Sub

Public Sub NovoOrcDoCli(Cliente As String, Placa As String, NrCli As Integer, Endereco As String, Carro As String)
OrcDoCLi = True
cbCliente.Text = Cliente
cbCliente.Enabled = False
MostraCarrosDoCLi Placa, NrCli, Endereco
MsData.Text = Format$(Now, "dd/mm/yyyy")
strcbCarro = Carro

'5.0.0 Conserto da visualização dos valores dos itens quando tem mais que a quantidade visivel
tx2Vi = True
End Sub

Private Sub VeTotalGrid()
Dim a     As Integer
Dim Valor As Double

GridT.Row = 0
GridT.Col = 2
TotalGrid = 0
On Local Error GoTo Err_VeTotalGrid
Do While GridT.EOF = False
    TotalGrid = TotalGrid + Valo(GridT, GridT.Text)
    If (GridT.Row + 1) > (GridT.ApproxCount - 1) Then
        Exit Do
    End If
    GridT.Row = GridT.Row + 1
Loop

Sai_VeTotalGrid:
On Local Error GoTo 0
Total -1
Exit Sub

Err_VeTotalGrid:
Loga "Erro de tipo: " & Error(Err) & " em VeTotalGrid ", lERR
Resume Sai_VeTotalGrid
End Sub

Private Sub Form_Activate()
Dim Controle   As Control
Dim MaxTarefas As Long
Dim MaxOrc     As Long
Dim X          As Long
Dim SQL        As String

If FezActivate Then Exit Sub

Balconista = 0
strcbCliente = ""
clsCLi.CarCliCB cbCliente
If NrOrcamento = 0 Then
    MsData.Text = Format$(Now, "dd/mm/yyyy")
    Text(11).Text = STR(INI.Garantia)
    If OrcDoCLi = False Then
        Text(0).Text = ""
        Text(1).Text = ""
        Text(2).Text = ""
        MudandoCombo = True
        cbCarro.ListIndex = -1
        MudandoCombo = False
    End If
Else
    If FezActivate = False Then
        MOSTRAORCAMENTO
        
        'Verificar se o endereço apareceu
        If Trim(Text(0).Text = "") Then
            Text(0).Text = Consulta("SELECT Clientes.Ender from Carros INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli Where Carros.Placa = '" & Text(2).Text & "'")
        End If
        ExecSql "Delete From ItensConcertoTemp Where IDPC = " & INI.PC
    End If
End If
FezActivate = True
Me.Top = 0
End Sub

Private Function DizGarantia(DtPagamento As Date, MesesGarantia As Integer) As String
Dim Aux$
Dim DtGarantia As Date

DtGarantia = DtPagamento + (30 * MesesGarantia)
Aux$ = "Garantia até " & Format(DtGarantia, "DD/MM/YYYY")
If DtGarantia < Now Then
    Aux$ = Aux$ & " - (Garantia Expirada)"
End If
DizGarantia = Aux$
End Function

Private Sub Form_Load()
Dim rsMec As Recordset
Dim a     As Integer

'4.0.5 Impedir duplicações de itens no novo orçamento
OrcNovo = True

Falta = 0
PriTotal = 0
DataT.DatabaseName = Base
dataConsertos.DatabaseName = Base
dataMecs.DatabaseName = Base
DataT.RecordSource = "Select * From TarefasTemp Where IDPC = " & INI.PC
OrcDoCLi = False
NrOrcamento = 0
Erase ITO()
FezActivate = False
ItensDig = 0
FAlterado = False
PAGO = 0
AQuemPagou = 0
TotalGrid = 0
InicForm Me
If Orcas.Mensal > "" Then
   SetaCliente
End If
lcImprEmBra = INI.ImprEmBrano

If INI.ModoOperacao <> tpMecanico Then

    GravaAutom = INI.GravaAutom
    
    If clsCLi.NrCli > 0 Then
        If Consulta("Select Funcionario From Clientes Where NrCli = " & clsCLi.NrCli) = True Then
            GravaAutom = False
        End If
    End If
    
End If

Loga "Orçamento Carregado"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Linha = 0
MomChamOrc = -1
AQuemPagou = 0
End Sub

Private Sub GridT_AfterColUpdate(ByVal ColIndex As Integer)
If Mudando = False Then
    VeTotalGrid
End If
End Sub

Private Sub GridT_AfterInsert()
ExecSql "Update TarefasTemp Set IDPC = " & INI.PC & " Where IDPC is null "
End Sub

Private Sub GridT_ComboSelect(ByVal ColIndex As Integer)
ExecSql "Update TarefasTemp Set Situacao = 'Em Andamento' Where Situacao = '" & SemMec & "' and Nome > '' and IDPC = " & INI.PC
End Sub

Private Sub MnuDesfaz_Click()
Dim Result As Boolean

frmSenha.Show 1
Result = frmSenha.Resultado
Unload frmSenha
If Result Then
    ExecSql "Update Orcamento Set Pagamento = null, Vend=null, VlrPago=0 Where Orcamento = " & l_NrOrcamento
    ExecSql "Delete From Parcelas Where Orc = " & l_NrOrcamento
    msgboxL "Pagamento desfeito", vbInformation
    Unload Me
End If
End Sub

Private Sub MnuImprGraf_Click()
Imprimir
End Sub

Private Sub MnuItens_Click()
'3.1.3
Load frmImprOpOrc
frmImprOpOrc.NrOrc = NrOrcamento
'LinhaDeComando = cbCliente.Text
frmImprOpOrc.Show
End Sub

Private Sub MsData_GotFocus()
Seleciona
End Sub

Private Sub MsData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub MsData_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Screen.ActiveControl.SelStart = 0
    KeyAscii = 0
End If
End Sub

Private Sub opObs_Click(Index As Integer)
If Mudando = False Then
    Text(8).Visible = (Index = 0)
    Text(12).Visible = (Index = 1)
    Text(13).Visible = (Index = 2)
    opObs(0).Value = IIf(Index = 0, 1, 0)
    opObs(1).Value = IIf(Index = 1, 1, 0)
    opObs(2).Value = IIf(Index = 2, 1, 0)
End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
If Index <> 8 Then Seleciona
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Not (Index = 8 Or Index = 12 Or Index = 13) Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Backspace = False

If KeyAscii = 13 And Not (Index = 8 Or Index = 12 Or Index = 13) Then
    KeyAscii = 0
ElseIf Index = 8 And KeyAscii = 8 Then
    Backspace = True
Else

    If INI.ModoOperacao = tpMecanico And Index <> 12 Then
        KeyAscii = 0
    Else

        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        pIndex = Index
        Alterado = True
    End If
        
End If
End Sub

Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Aux      As String

Const Lim = 78

If (Index = 8 Or Index = 12) And Backspace = False Then
    Aux = Text(Index).Text
    If Len(Aux) > Lim Then
        Aux = Right(Aux, Lim)
        If InStr(Aux, Chr(13)) = 0 Then
            Beep
            SendKeys "{ENTER}"
        End If
    End If
End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
Dim MecanicaAux As Double
Dim EletricaAux As Double

If Index < 8 Then
    If Index = 3 Or Index = 6 Then
        If Text(Index) <> "0" Then
            MostraValor Text(Index)
        End If
        If Trim(Text(3).Text) = "" Then
            MecanicaAux = -1
        Else
            MecanicaAux = Valo(Text(3))
        End If
        If Trim(Text(6).Text) = "" Then
            EletricaAux = -1
        Else
            EletricaAux = Valo(Text(3))
        End If
    Else
        MostraValor Text(Index)
    End If
End If
Total 0
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Seleciona
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Text1(Index).Text = "" Then
        If Text(8).Visible Then
            Text(8).SetFocus
        Else
            Text(12).SetFocus
        End If
    Else
        SendKeys "{Tab}"
    End If
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
Else
    
    If INI.ModoOperacao = tpMecanico Then
        KeyAscii = 0
    Else
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        pIndex = Index
        Alterado = True
    End If
End If
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Shift = 2 Then
    ExclusaoItens Index
Else
    If GravaAutom Then
        If TimerItens.Enabled Then
            TimerItens.Enabled = False
            TimerItens.Enabled = True
        End If
    End If
    
End If
End Sub

Private Function ExclusaoItens(Index As Integer)
Dim Sair   As Boolean
Dim Cont   As Integer
Dim iLinha As Integer

iLinha = Index

'4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
Dim EsseCont As Integer
EsseCont = NrIt(Index)
ITO(EsseCont).Deletar = True
ItemAlterado = True
UltTot = UltTot - ITO(EsseCont).Valor

Cont = EsseCont
'Cont = NrIt(Index)

'Cont% = Linha + Index
'UltTot = UltTot - ITO(Cont).Valor

Do While Sair = False
    '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
    'UltTot = UltTot - ITO(Cont).Valor
    'ITO(Cont) = ITO(Cont + 1)
    'UltTot = UltTot + ITO(Cont).Valor
    
    If iLinha < 7 Then
    
        '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
        EsseCont = NrIt(EsseCont)
        If EsseCont > -1 Then
            If ITO(EsseCont).Peca > "" Then
                Text1(iLinha).Text = ITO(EsseCont).Peca
                MostraValor Text2(iLinha), ITO(EsseCont).Valor
                txNome(iLinha).Text = ITO(EsseCont).Nome
                txQuant(iLinha).Text = STR$(ITO(EsseCont).Quant)
                txData(iLinha).Text = Format(ITO(EsseCont).Data, "DD/MM/YY")
                Debug.Print iLinha & " " & ITO(EsseCont).Peca
                If ITO(EsseCont).Existente = False Then
                    Sair = True
                End If
            Else
            
                '5.0.0 Apagar o restante da linha ao deletar a ultima linha dos itens
                Text1(iLinha).Text = ""
                Text2(iLinha).Text = ""
                txNome(iLinha).Text = ""
                txQuant(iLinha).Text = ""
                txData(iLinha).Text = ""
            
                Sair = True
            End If
        Else
            Text1(iLinha).Text = ""
            Text2(iLinha).Text = ""
            txNome(iLinha).Text = ""
            txQuant(iLinha).Text = ""
            txData(iLinha).Text = ""
            Sair = True
        End If
        '4.1.0 Correção da deleção de peças quando tem mais de 7 itens
        EsseCont = EsseCont + 1
        iLinha = iLinha + 1
    Else
        Sair = True
    End If
Loop

'4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
If iLinha < 7 Then
    txData(iLinha).Text = ""
    txQuant(iLinha).Text = ""
End If

'3.7.3 Log grava o conteudo do total do orçamento
SetaLbTotal UltTot, 2
'MostraValor Label2, UltTot
End Function

Private Sub Text1_LostFocus(Index As Integer)
Dim EsseCont As Integer

EsseCont = NrIt(Index, True)
If Trim(ITO(EsseCont).Peca) <> Trim(Text1(Index).Text) Then
    ITO(EsseCont).Peca = Text1(Index).Text
    ITO(EsseCont).Alterado = True
    ItemAlterado = True
End If
If ItensDig < (EsseCont) Then ItensDig = EsseCont
If ITO(EsseCont).Data = 0 Then
    ITO(EsseCont).Data = Now
End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
Seleciona
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Vlr As Double

If KeyCode = 13 Then
    If Index = 6 Then
        If Text2(Index).Text = "" Then
            SendKeys "{Tab}"
        Else
            
            Vlr = Valo(Text2(Index))
            
            '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
            Dim EsseCont As Integer
            EsseCont = NrIt(Index)
            If ITO(EsseCont).Valor <> Vlr Then
                ITO(EsseCont).Valor = Vlr
                ITO(EsseCont).Alterado = True
                ItemAlterado = True
                Total Index
            End If
'            If ITO(Linha + Index).Valor <> Vlr Then
'                ITO(Linha + Index).Valor = Vlr
'                ITO(Linha + Index).Alterado = True
'                ItemAlterado = True
'                Total Index
'            End If
'            'ITO(Linha + Index).Valor = Valo(Text2(Index))
            
            VScroll1.Max = Index + Linha
            VScroll1.Enabled = True
            
            Dim Valor As Double
            VeValor Text2(Index).Text, Valor, Text2(Index), 1
            ITO(NrIt(Index)).Valor = Valor
            Total Index
            
            Linha = Linha + 1
            
            Scroll
            Text1(Index).SetFocus
            
        End If
    Else
        SendKeys "{Tab}"
    End If
End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
Else
    pIndex = Index
    Alterado = True
    
    '4.1.1 Conserto da alteração dos itens
    Dim Cont%
    Cont% = NrIt(Index)
    ItemAlterado = True
    ITO(Cont%).Alterado = True
    
End If
End Sub

Private Sub Text2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If GravaAutom Then
    If KeyCode <> 13 Then
        MarcaItemAlterado
    End If
End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
If Index < 6 Then
    Dim Valor As Double
    
    VeValor Text2(Index).Text, Valor, Text2(Index), 1
    
    '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
    ITO(NrIt(Index)).Valor = Valor
    'ITO(Linha + Index).Valor = Valor
    
    Total Index
End If
End Sub

Private Sub Timer_Timer()
Timer.Enabled = False
Loga "cbCarro.Text[2] = " & cbCarro.Text
If cbCarro.Text <> strcbCarro Then
    MudandoCombo = True
    TextCombo strcbCarro, cbCarro, 1
    MudandoCombo = False
End If
If MomChamOrc > -1 Then
    Loga "Tempo de carregamento do Orçamento: " & Right(Format(Now - MomChamOrc, "hh:mm:ss.mm"), 7)
End If
End Sub

Private Sub TimerItens_Timer()
Loga "TimerItens", lDBG
TimerItens.Enabled = False

If NrOrcamento = 0 Then
    
    '4.0.5 Impedir gravação automática sem a identificação do balconista primeiro
    If Balconista > 0 Then
            
        If UltTot > 0 Then
            Gravar True
        End If
    End If
    
Else
    Botao(0).Enabled = False
    Dim GravarAgora As Boolean
    If TemItens Then
        If Balconista Then
            GravarAgora = True
        End If
    Else
        GravarAgora = True
    End If
    If GravarAgora Then
    
        '4.5.8 Mostrar a gravação automática
        Caption = "Orçamento - gravado automaticamente as " & Format(Now, "HH:MM:SS")
        
        GravaItens
    End If
    'GravaItens
    
    Botao(0).Enabled = True
End If
End Sub

Private Sub MarcaItemAlterado()
If ItensDig > 0 Then
    If TimerItens.Enabled = False Then
        TimerItens.Enabled = True
        Loga "MarcaItemAlterado - Inicializado", lDBG
    Else
        TimerItens.Enabled = False
        TimerItens.Enabled = True
        Loga "MarcaItemAlterado - REInicializado", lDBG
    End If
    Botao(0).Enabled = True
End If
End Sub

Private Sub txNome_DblClick(Index As Integer)
If txNome(Index).Text > "" Then
    Load AuditEdicItem
    
    '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
    AuditEdicItem.NrItem = ITO(NrIt(Index)).ID
    'AuditEdicItem.NrItem = ITO(Linha + Index).ID
    
    AuditEdicItem.Show
End If
End Sub

Private Sub txQuant_GotFocus(Index As Integer)
Seleciona
End Sub

Private Sub txQuant_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txQuant_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
Else

    If INI.ModoOperacao = tpMecanico And Index <> 12 Then
        KeyAscii = 0
    Else

        pIndex = Index
        Alterado = True
    End If
End If
End Sub

Private Sub txQuant_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If GravaAutom Then
If TimerItens.Enabled Then
        TimerItens.Enabled = False
        TimerItens.Enabled = True
    End If
End If
End Sub

Private Sub txQuant_LostFocus(Index As Integer)
Dim Quant As Single
Quant = Val(txQuant(Index).Text)

Dim EsseCont As Integer
EsseCont = NrIt(Index)
If ITO(EsseCont).Quant <> Quant Then
    ITO(EsseCont).Quant = Quant
    ITO(EsseCont).Alterado = True
    ItemAlterado = True
End If
'If ITO(Linha + Index).Quant <> Quant Then
'    ITO(Linha + Index).Quant = Quant
'    ITO(Linha + Index).Alterado = True
'    ItemAlterado = True
'End If

'ITO(Linha + Index).Quant = Val(txQuant(Index).Text)
End Sub

Private Sub SetaCliente()
If Not (OrcasRecordset.EOF Or OrcasRecordset.BOF) Then
    clsCLi.SetaClienteAtual OrcasRecordset("Cliente")
    clsCLi.NrCli = OrcasRecordset("NrCli")
    If IsNull(clsCLi.Endereco) = False Then
        Text(0).Text = clsCLi.Endereco
    End If
    If cbCarro.ListCount = 0 Then
       clsCLi.CarregaCarros cbCarro
    End If
    If cbCarro.ListCount = 0 Then
        cbCarro.ListIndex = -1
    Else
        cbCarro.ListIndex = 0
    End If
End If
Mudando = True
If Text(8).Visible = False And Text(12).Visible - False Then
    Text(8).Visible = True
    Text(12).Visible = False
End If
opObs(0).Value = Text(8).Visible
opObs(1).Value = Text(12).Visible
Mudando = False
Loga "Saiu do SetaCliente", lDBG
End Sub

Private Sub VScroll1_Change()
If AltScrol = False Then
   Linha = VScroll1.Value
   Scroll
End If
End Sub

Public Property Get NrOrcamento() As Integer
NrOrcamento = l_NrOrcamento
End Property

Public Property Let NrOrcamento(vNewValue As Integer)
Dim SQL    As String
Dim Selec  As String
Dim Result As Boolean

l_NrOrcamento = vNewValue

If l_NrOrcamento > 0 Then
    If Gravando = False Then
        If vNewValue > 0 Then
            SQL = "Select DISTINCTROW Orcamento.Obs, "
            SQL = SQL & "Orcamento.Carro, Orcamento.Data, Orcamento.Orcamento, Orcamento.Cliente,"
            SQL = SQL & "Orcamento.Kms, Orcamento.Kilom, Orcamento.Total, Orcamento.Garantia, Orcamento.Pagamento, "
            SQL = SQL & "Itens_Orc.Item, Itens_Orc.Quant, Itens_Orc.Valor "
            SQL = SQL & ",Orcamento.ObsMec "
            SQL = SQL & ",Orcamento.Vend "
            SQL = SQL & ",Orcamento.VlrPago "
            SQL = SQL & ",Itens_Orc.DtItOrc "
            SQL = SQL & ",Itens_Orc.Cont "
            
            '4.6.8 Impressão da observação do cliente na impressão do orçamento
            SQL = SQL & ",Clientes.Observacao "
            SQL = SQL & "from Clientes "
            SQL = SQL & "INNER JOIN (Orcamento LEFT JOIN Itens_Orc ON Orcamento.Orcamento = Itens_Orc.Orçamento) ON Clientes.Nome = Orcamento.Cliente "
            SQL = SQL & "Where Orcamento.Orcamento = " & l_NrOrcamento & " "
            SQL = SQL & "ORDER BY Itens_Orc.Item "
'            SQL = SQL & "FROM Orcamento "
'            SQL = SQL & "LEFT JOIN Itens_Orc ON Orcamento.Orcamento = Itens_Orc.Orçamento "
'            SQL = SQL & "WHERE Orcamento.Orcamento=" & l_NrOrcamento
'            SQL = SQL & " Order By Itens_Orc.Item "
            
            AbreTB Orca, SQL, dbOpenDynaset
            clsCLi.NrCli = clsCLi.NrDoNome(Orca("Cliente"))
            If Orca.EOF Then
                Orca.Close
                Selec = "SELECT DISTINCTROW Obs, Guarnição, Eletricidade, Chapeação, Pintura, Mecanica, Carro, Data, Orcamento, Cliente, Kms, Kilom, Total  "
                Selec = Selec & "From Orcamento "
                Selec = Selec & "WHERE ((Orcamento=" & l_NrOrcamento & ")); "
                AbreTB Orca, SQL, dbOpenDynaset
                TemItens = False
            Else
                TemItens = True
            End If
            Placa = Trim(Orca("Carro"))
            Text(13).Text = SN(Consulta("Select Historico From Carros Where Placa = '" & Placa & "'"), vbString)
            strcbCarro = Consulta("SELECT Modelo from Carros WHERE Carros.Placa = '" & Trim(Orca("Carro")) & "'")
            SQL = "SELECT Tarefas.Mec, Tarefas.concerto, Mecanicos.Nome, Tarefas.Situacao "
            SQL = SQL & "FROM Tarefas INNER JOIN Mecanicos ON Tarefas.Mec = Mecanicos.codi "
            SQL = SQL & "Where Tarefas.Orc = " & NrOrcamento
            SQL = SQL & " ORDER BY Tarefas.concerto "
            Mudando = False
            frmOrc.Timer.Enabled = True
        End If
    End If
End If
End Property

Private Property Get Placa()
Placa = GPlaca
End Property

Private Property Let Placa(vNewValue)
GPlaca = vNewValue
End Property

Public Sub ImprimeFita()
Dim TamPeca  As Integer
Dim a        As Integer
Dim Tam      As Integer
Dim Linha    As String
Dim Aux      As String
Dim Aux1     As String
Dim Aux2     As String
Dim sObs     As String
Dim MaObra   As Currency
Dim Cont%
Dim TbConfig As Recordset

'4.2.1 Impressão da observação
'Const TamFita = 55

'4.2.7 Inscrição Estatual
Dim CgcCpf   As String
Dim InscrEst As String
Dim rsCli    As Recordset

Dim sCEP     As String

If lbTotal.Tag <> lbTotal.Caption Then
    Gravar False
End If

ImprBuferizada_Inicializa
LPRINT Chr$(15)
LPRINT CENTRAL("ORCAMENTO", TamFita / 2)
LPRINT CENTRAL(SemAcento(INI.Empresa), TamFita / 2)
LPRINT CENTRAL(SemAcento(INI.Endereco), TamFita / 2)
LPRINT CENTRAL(INI.Fones, TamFita / 2)
LPRINT " "
LPRINT "Porto Alegre, " & Format$(CDate(MsData.Text), "dd") & " de " & Format$(CDate(MsData.Text), "mmmm") + " de " + Format$(CDate(MsData.Text), "yyyy")

'4.9.1 Ajuste da Impressão em Fita via Usb
LPRINT UmaOuDuasLinhas("Proprietario do Veiculo: ", SemAcento(cbCliente.Text))
LPRINT UmaOuDuasLinhas("Endereco: ", SemAcento(Text(0).Text))
'LPRINT "Proprietario do Veiculo: " & SemAcento(cbCliente.Text)
'LPRINT "Endereco: " & SemAcento(Text(0).Text)

'4.2.8 CEP
If sCEP > "" Then
    LPRINT "CEP: " & sCEP
End If

AbreTB rsCli, "Select CgcCpf, InscrEst From Clientes Where NrCli = " & clsCLi.NrCli
CgcCpf = SN(rsCli!CgcCpf)
InscrEst = SN(rsCli!InscrEst)
rsCli.Close

If InscrEst > "" Then

    '4.9.1 Ajuste da Impressão em Fita via Usb
    LPRINT UmaOuDuasLinhas("Inscrição Estadual: ", InscrEst)
    'LPRINT "Inscrição Estadual: " & InscrEst
    
End If

If CgcCpf > "" Then
    LPRINT "CNPj/CPF: " & CgcCpf
End If

'4.9.1 Ajuste da Impressão em Fita via Usb
LPRINT UmaOuDuasLinhas("Veiculo Marca: ", SemAcento(cbCarro.Text & " " & Text(1).Text))
'LPRINT "Veiculo Marca: " & SemAcento(cbCarro.Text & " " & Text(1).Text)

LPRINT "Placa " & Text(2).Text
If GridT.VisibleRows > 0 Then
    GridT.Row = 0
End If
TotalGrid = 0

'5.0.0 Ajuste na impressão em fita
If TamFita = 0 Then TamFita = 55

If GridT.ApproxCount > 0 Then
    GridT.MoveFirst
    For a = 0 To (GridT.ApproxCount - 1)
        Aux = " " & Format(GridT.Columns(2).Value, "#,###,##0.00")
        Aux1 = "Mao de Obra " & GridT.Columns(1).Value
        
        '4.9.1 Ajuste da Impressão em Fita via Usb
        LPRINT Aux1 & Space(TamFita - Len(Aux) - Len(Aux1)) & Aux
        'LPRINT Aux1 & Space(55 - Len(Aux) - Len(Aux1)) & Aux
        
        GridT.MoveNext
    Next
End If
LPRINTST "Garantia: ", Trim(Text(11).Text), "Meses"
If PAGO > 0 Then
    LPRINT "PAGO" & IIf(lbVendedor.Caption > "", " [" & lbVendedor.Caption & "]", "")
    LPRINT DizGarantia(PAGO, Val(Text(11).Text))
End If
If TotalGrid > 0 Then LPRINT "Soma da Mao de Obra: " & Format$(TotalGrid, "###,###.00")
LPRINT "Combustivel: " & Text(9).Text
LPRINT "Kms: " & Text(10).Text

'4.9.1 Ajuste da Impressão em Fita via Usb
LPRINT "Pecas" & Space(TamFita - 18) & "Quant.  Valor"
'LPRINT "Pecas" & Space(37) & "Quant.  Valor"

'----5----1---15---20---25---30---35---40---45---50---55
'Pecas                                      QTD  Valor"
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx 999 8,888.00
'TOTAL..........................................8,888.00

'4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
Dim EsseCont As Integer

For Cont% = 0 To ItensDig

    EsseCont = NrIt(Cont%)
    
    '4.1.0 Conserto da impressão do orçamento, depois da versão 4.0.5
    If EsseCont > -1 Then
    
        Aux = SemAcento(Trim(ITO(EsseCont).Peca))
        
        Tam = Len(Aux)
        
        If Tam > 0 Then
        
            '4.9.1 Ajuste da Impressão em Fita via Usb
            If Tam > (TamFita - 14) Then
            
                '4.9.3 Ajuste na impressão USB
                LPRINT Aux
                Linha = Space(18)
                'LPRINT Left(Aux, (TamFita - 13))
                'Aux = ComplStr(Right(Aux, Tam - TamFita - 13), TamFita - 12 - Tam, " ", 0)
                
            Else
                Linha = Aux & Space(TamFita - 12 - Tam)
            End If
    '        If Tam > 41 Then
    '            LPRINT Left(Aux, 42)
    '            Aux = ComplStr(Right(Aux, Tam - 42), 43 - Tam, " ", 0)
    '        Else
    '            Linha = Aux & Space(43 - Tam)
    '        End If
            
            Linha = Linha & ComplStr(Trim(STR(ITO(EsseCont).Quant)), 4, " ", 1)
            LPRINT Linha & ComplStr(Format$(ITO(EsseCont).Valor, "###,###.00"), 8, " ", 2)
        '    Aux = SemAcento(Trim(ITO(Cont%).Peca))
        '    Tam = Len(Aux)
        '    If Tam > 41 Then
        '        LPRINT Left(Aux, 42)
        '        Aux = ComplStr(Right(Aux, Tam - 42), 43 - Tam, " ", 0)
        '    Else
        '        Linha = Aux & Space(43 - Tam)
        '    End If
        '    Linha = Linha & ComplStr(Trim(STR(ITO(Cont%).Quant)), 4, " ", 1)
        '    LPRINT Linha & ComplStr(Format$(ITO(Cont%).Valor, "###,###.00"), 8, " ", 2)
        End If
    End If

Next
LPRINT " "

'4.9.1 Ajuste da Impressão em Fita via Usb
LPRINT "TOTAL" & Space(TamFita - 11) & lbTotal.Caption
'LPRINT "TOTAL" & Space(TamFita - 13) & lbTotal.Caption
'3.7.3 Log grava o conteudo do total do orçamento
'LPRINT "TOTAL" & Space(42) & lbTotal.Caption

If Consulta("Select Count(*) From IC_Orc Where Orc = " & NrOrcamento) > 0 Then
    LPRINT " "
    ImprimeItensConcerto
End If
If Text(8).Text > "" Then
    LPRINT "OBSERVACAO"
    a = InStr(Text(8).Text, Chr(13))
    If a > 0 Then
       Aux1 = Text(8)
       Do While Len(Aux1) > 0
          a = InStr(Aux1, Chr(13))
          If a > 0 Then
             Aux2 = Left(Aux1, InStr(Aux1, Chr(13)) - 1)
             TamPeca = Len(Aux2)
             If TamPeca > TamFita Then
                Do While TamPeca > TamFita
                    Aux2 = Left(Aux1, TamFita)
                    a = InStr(Aux1, Chr(13))
                    If a > 0 And a < TamFita Then
                        Aux2 = Left(Aux1, a - 1)
                        Aux1 = Mid(Aux1, Len(Aux2) + 3)
                    Else
                        Aux1 = Mid(Aux1, TamFita + 1)
                    End If
                    TamPeca = Len(Aux1)
                    LPRINT Aux2
                Loop
                Aux2 = Aux1
             Else
                Aux1 = Mid(Aux1, Len(Aux2) + 3)
             End If
          Else
             Aux2 = Aux1
             Aux1 = ""
          End If
          Aux2 = Replace(Aux2, Chr(10), "")
          LPRINT Aux2
       Loop
    Else
       LPRINT Text(8).Text
    End If
End If
sObs = Consulta("Select Observacao From Clientes Where NrCli = " & clsCLi.NrCli)
If sObs > "" Then
    LPRINT "Observação: " & sObs
End If
'If INI.LinhasApos > 0 Then
'    For a = 0 To INI.LinhasApos
'        LPRINT " "
'    Next
'End If

ImprBuferizada_Finaliza
End Sub

'4.9.1 Ajuste da Impressão em Fita via Usb
Private Function UmaOuDuasLinhas(Parte1 As String, Parte2 As String) As String
If TamFita = 55 Then
    UmaOuDuasLinhas = Parte1 & Parte2
Else
    LPRINT Parte1
    UmaOuDuasLinhas = Parte2
End If
End Function

Private Sub ImprimeItensConcerto()
Dim a            As Integer
Dim SQL          As String
Dim rsIC_Orc(3)  As Recordset
Dim rsConfMod(3) As Recordset
Dim rsTitulos    As Recordset

Const Margem = "     "

LPRINT "ITENS A SEREM CONCERTADOS"
AbreTB rsTitulos, "Select TitModelo1, TitModelo2, TitModelo3, TitModelo4 From Config "
For a = 0 To 3
    SQL = "Select Conteudo, Linha From ConfigModelo Where Conteudo > '' and Coluna = " & (a + 1)
    SQL = SQL & " Order By Linha "
    AbreTB rsConfMod(a), SQL
    SQL = "Select Lin From IC_Orc Where Orc = " & NrOrcamento
    SQL = SQL & " and Col = " & (a + 1)
    SQL = SQL & " ORDER BY Lin "
    AbreTB rsIC_Orc(a), SQL
    If rsIC_Orc(a).EOF = False Then
        LPRINT rsTitulos.Fields(a).Value
        Do While rsConfMod(a).EOF = False
            If rsConfMod(a)!Linha = rsIC_Orc(a)!lin Then
                LPRINT Margem & rsConfMod(a)!Conteudo
                rsIC_Orc(a).MoveNext
                If rsIC_Orc(a).EOF Then
                    Exit Do
                End If
            End If
            rsConfMod(a).MoveNext
        Loop
    End If
Next
LPRINT " "
End Sub

Public Property Get Alterado() As Boolean
Alterado = FAlterado
End Property

Public Property Let Alterado(ByVal vNewValue As Boolean)
If FAlterado = False And vNewValue = True Then
    Botao(0).Visible = True
    Botao(0).Enabled = True
End If
FAlterado = vNewValue
End Property

Private Function TarefasEstaoConcluidas() As Boolean
Dim a As Integer

If GridT.ApproxCount = 0 Then
    TarefasEstaoConcluidas = True
Else
    GridT.MoveFirst
    For a = 0 To (GridT.ApproxCount - 1)
        If GridT.Columns(7).Value = "" Then
            TarefasEstaoConcluidas = False
            Exit Function
        End If
        GridT.MoveNext
    Next
End If
TarefasEstaoConcluidas = True
End Function

'4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
Private Function NrIt(Cont As Integer, Optional Adic As Boolean) As Integer
Dim Sair As Boolean

NrIt = Linha + Cont
Do While Sair = False
    If ITO(NrIt).Existente = True And ITO(NrIt).Deletar = False Then
        Sair = True
    Else
        If ITO(NrIt).Existente = False Then
        
            '4.0.7 Ajuste na reimplantação da deleção dos itens do orçamento
            If Adic Then
                ITO(NrIt).Existente = True
                Sair = True
                
            Else
                If ITO(NrIt).Peca = "" Then
                    Sair = True
                    NrIt = -1
                End If
            End If
        Else
            NrIt = NrIt + 1
        End If
    End If
Loop
End Function

Private Sub PagamentoRecebido(VlrPago As Currency)
'5.0.4 Correções de bugs
Dim SQL        As String

PAGO = Now

'5.0.4 Correções de bugs
If VlrPago = Falta Or VlrPago = PriTotal Then
'If VlrPago = Falta Then

    btPagamento.Caption = "Pago"
Else
    btPagamento.Caption = "Pagamento Parcial"
End If
If l_NrOrcamento > 0 Then
    SQL = "Update Orcamento Set Pagamento = " & DTSqls(Format(Now, "DD/MM/YYYY"))
    SQL = SQL & ", Vend = " & AQuemPagou
    
    '5.0.4 Correções de bugs
    If UltTot > 0 Then
        SQL = SQL & ", Total = " & VlrSql(Trim(STR(UltTot)))
    End If
    
    SQL = SQL & ", VlrPago = " & VlrSql(Trim(STR(VlrPago)))
    SQL = SQL & " WHERE Orcamento.Orcamento=" & l_NrOrcamento
    ExecSql SQL
    
    '4.5.5 Em modo Escritório pode editar cliente em atraso
    If clsCLi.EmAtraso Then
        ReConfereAtraso
    End If
    
End If
pIndex = 0
Alterado = True
Botao(0).Enabled = False
Botao(2).Caption = "Fechar"

'4.3.2 Permitir concluir as tarefas no pagamento
If FazerPagar Then
    ExecSql "Update Tarefas Set Situacao = 3 Where Orc = " & NrOrcamento
    DataT.Enabled = False
    CarregaTarefas NrOrcamento
    DataT.Enabled = True
    DataT.Refresh
End If
End Sub

