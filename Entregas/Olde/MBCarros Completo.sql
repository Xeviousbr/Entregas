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

CREATE TABLE [ConfigModelo] 
(   [Coluna] INTEGER,
   [Linha] INTEGER,
   [Conteudo] TEXT(20),
   [Valor] MEMO 

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

CREATE TABLE [Ferramentas] 
(   [Codigo] TEXT(10),
   [Descricao] TEXT(20),
   [Marca] TEXT(10),
   [Func] TEXT(30),
   [Data] DATETIME 

)                      

CREATE TABLE [FerrMec] 
(   [ID] COUNTER,
   [idMec] INTEGER,
   [codigo] TEXT(10),
   [Data] DATETIME 

)

CREATE TABLE [Fornecedores] 
(   [codi] COUNTER,
   [Nome] TEXT(30),
   [Ender] TEXT(30),
   [Telefone] TEXT(10),
   [Fax] TEXT(10),
   [Vendedor] TEXT(30),
   [UltForn] DATETIME,
   [RazaoDoc] TEXT(30),
   [CGC] TEXT(14),
   [Bai] TEXT(20),
   [Cid] TEXT(30),
   [CEP] TEXT(10),
   [TemEntrPed] SMALLINT,
   [idBanco] INTEGER,
   [Age] TEXT(10),
   [Conta] TEXT(10),
   [email] TEXT(40) 

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

CREATE TABLE [ItensConcertoTemp] 
(   [ID] COUNTER,
   [Coluna] SMALLINT,
   [Linha] SMALLINT,
   [Valor] SMALLINT,
   [Conteudo] TEXT(20),
   [IDPC] INTEGER 

)                            

CREATE TABLE [Mecanicos] 
(   [codi] INTEGER,
   [Nome] TEXT(30),
   [Ende] TEXT(30),
   [Telefone] TEXT(20),
   [PercComiss] FLOAT,
   [Senha] TEXT(8),
   [Oper] SMALLINT,
   [Recebe] WORD,
   [rg] TEXT(10),
   [Ativo] YESNO,
   [TpRec] INTEGER 

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

CREATE TABLE [PagtoForn] 
(   [idPagtoForn] COUNTER,
   [idForn] INTEGER,
   [Valor] MEMO,
   [Data] DATETIME,
   [DOC] TEXT(20),
   [Obs] MEMO,
   [QtdParc] INTEGER 

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

CREATE TABLE [PCs] 
(   [IDPC] COUNTER,
   [Nome] TEXT(20) 

)

CREATE TABLE [Tarefas] 
(   [id] COUNTER,
   [Orc] INTEGER,
   [Mec] SMALLINT,
   [Vlr] MEMO,
   [concerto] SMALLINT,
   [Situacao] SMALLINT,
   [Pago] DATETIME,
   [DtConclusao] DATETIME,
   [DtAssumiu] DATETIME 

)

CREATE TABLE [TarefasTemp] 
(   [ID] COUNTER,
   [concerto] TEXT(15),
   [Vlr] MEMO,
   [Situacao] TEXT(15),
   [Nome] TEXT(30),
   [Pago] DATETIME,
   [IDPC] INTEGER,
   [DtConclusao] DATETIME,
   [dtassumiu] DATETIME 

)

CREATE TABLE [tpConcertos] 
(   [tipo] SMALLINT,
   [concerto] TEXT(15),
   [Mec] SMALLINT

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

