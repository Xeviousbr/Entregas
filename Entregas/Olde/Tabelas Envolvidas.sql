CREATE TABLE [CliAlfa]
(   [NrCli] COUNTER,
   [Nome] TEXT(40),
   [Ender] TEXT(40),
   [Telefone] TEXT(30)

)

CREATE TABLE [CliDupl] 
(   [Expr1] INTEGER,
   [Nome] TEXT(40) 

)           

CREATE TABLE [Clientes] 
(   [NrCli] COUNTER,
   [Nome] TEXT(40),
   [Ender] TEXT(40),
   [Telefone] TEXT(30),
   [Observacao] MEMO,
   [Vend] INTEGER,
   [funcionario] YESNO,
   [email] TEXT(27),
   [CgcCpf] TEXT(20),
   [InscrEst] TEXT(10),
   [Operadora] INTEGER,
   [CEP] TEXT(8),
   [EmAtraso] YESNO,
   [foto] TEXT(20),
   [Apelido] TEXT(16) 

)

CREATE TABLE [Config] 
(   [Empresa] TEXT(24),
   [Endereco] TEXT(41),
   [Fones] TEXT(20),
   [Cor] INTEGER,
   [Imagem] TEXT(141),
   [TpImpress] INTEGER,
   [LinhasApos] INTEGER,
   [Versao] TEXT(5),
   [Garantia] INTEGER,
   [TitModelo1] TEXT(20),
   [TitModelo2] TEXT(20),
   [TitModelo3] TEXT(20),         
   [TitModelo4] TEXT(20),
   [CGC] TEXT(20),
   [VlrGatComiss] FLOAT,
   [UtComissoes] WORD,
   [Orc1] INTEGER,
   [LogEmRede] WORD,
   [QtdCarrComiss] INTEGER 

)

CREATE TABLE [Entregas] 
(   [ID] COUNTER,
   [idCliente] INTEGER,
   [idForma] INTEGER,
   [idBoy] INTEGER,
   [Obs] MEMO,
   [Valor] MEMO,
   [VlNota] MEMO,
   [Data] DATETIME,
   [Pago] MEMO 
                             
)

CREATE TABLE [IC_Orc] 
(   [ID] COUNTER,
   [Orc] INTEGER,
   [Col] INTEGER,
   [Lin] INTEGER

)

CREATE TABLE [Itens_Orc] 
(   [Cont] COUNTER,
   [Orçamento] SMALLINT,
   [Item] TEXT(50),
   [Quant] FLOAT,
   [Valor] MEMO,
   [DtItOrc] DATETIME,
   [NT] INTEGER 

)

CREATE TABLE [Orcamento] 
(   [Orcamento] COUNTER,
   [Data] DATETIME,
   [Cliente] TEXT(40),
   [Carro] TEXT(10),
   [Obs] MEMO,
   [Kms] TEXT(10),
   [Kilom] TEXT(10),
   [Total] MEMO,
   [Garantia] INTEGER,
   [Pagamento] DATETIME,
   [ObsMec] MEMO,
   [Vend] INTEGER,
   [VlrPago] MEMO 

)

CREATE TABLE [Parcelas] 
(   [idParc] COUNTER,
   [Orc] INTEGER,
   [Cli] INTEGER,
   [NrParc] INTEGER,
   [Data] DATETIME,
   [Valor] MEMO,
   [Pagto] DATETIME,
   [BalcFez] INTEGER,
   [BalcRec] INTEGER,
   [Obs] MEMO 

)

CREATE TABLE [ParcelasPagtoForn] 
(   [idParcelasPagtoForn] COUNTER,
   [idPagtoForn] INTEGER,
   [Nr] SMALLINT,
   [Valor] MEMO,
   [Data] DATETIME,
   [Pago] YESNO 

)

CREATE TABLE [ParcelasTemp] 
(   [Auto] COUNTER,
   [IDPC] INTEGER,
   [Data] DATETIME,
   [Valor] MEMO,
   [Orig] INTEGER,
   [Obs] TEXT(50) 

)

CREATE TABLE [tpSituacao] 
(   [tipo] COUNTER,
   [situacao] TEXT(15) 

)

CREATE TABLE [Vales] 
(   [ID] COUNTER,
   [IdOperador] SMALLINT,
   [Data] DATETIME,
   [Valor] MEMO,
   [Pago] DATETIME,
   [Tipo] SMALLINT,
   [obs] MEMO,
   [NomeAvulso] TEXT(20),
   [Periodo] TEXT(50),
   [txValor] TEXT(10) 

)

