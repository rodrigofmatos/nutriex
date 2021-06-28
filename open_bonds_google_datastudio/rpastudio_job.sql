----------------------------------------------------------
--// TEMA: ROBO TITULOS EM ABERTO - RPASTUDIO
--// -----------------------------------------------------
--// CRIADO POR: RODRIGO MATOS (CIENTISTA DADOS)
--// VERSAO: 3.0.0		
--// DATA: 18.MAR.2020
--// SOLIC.: BRANDAO/CARLA SILVA (CREDITO E COBRANCA)
--// -----------------------------------------------------

IF OBJECT_ID('[mbs].[dbo].[CONTAS_A_RECEBER]', 'U') IS NOT NULL begin DROP TABLE [mbs].[dbo].[CONTAS_A_RECEBER] end
IF OBJECT_ID('tempdb..##empresas', 'U') IS NOT NULL begin DROP TABLE ##empresas end

CREATE TABLE [mbs].[dbo].[CONTAS_A_RECEBER](
	[SEGMENTO] [nvarchar](50) NULL,
	[SEMAFORO] [nvarchar](50) NULL,
	[BASE] [nvarchar](50) NULL,
	[EMPRESA] [nvarchar](50) NULL,
	[Documento] [nvarchar](50) NULL,
	[DocumentoSAP] [int] NULL,
	[FormaCobranca] [varchar](20) NULL,
	[ValorTitulo] [decimal](18,2) NULL,
	[Emissao] [date] NULL,
	[Vencimento] [date] NULL,
	[Atraso] [int] NULL,
	[Categoria] [nvarchar](100) NULL,
	[Vendedor] [nvarchar](100) NULL,
	[Supervisor] [nvarchar](100) NULL,
	[Grupo] [nvarchar](50) NULL,
	[CodigoCliente] [nvarchar](15) NULL,
	[RazaoSocial] [nvarchar](100) NULL,
	[Municipio] [nvarchar](100) NULL,
	[Estado] [nvarchar](2) NULL,
	[Fone] [nvarchar](100) NULL,
	[Version] [nvarchar](100) NULL
) ON [PRIMARY]

DECLARE @loop int = (SELECT COUNT(baseid) FROM mbs..bases);
DECLARE @i int = 0;
DECLARE @empresa nvarchar(MAX);
DECLARE @pathocrd nvarchar(MAX);
DECLARE @pathINV6 nvarchar(MAX);
DECLARE @pathOINV nvarchar(MAX);
DECLARE @pathOSLP nvarchar(MAX);
DECLARE @pathOCRG nvarchar(MAX);
DECLARE @pathOCNT nvarchar(MAX);
DECLARE @pathFT_OSEG nvarchar(MAX);
DECLARE @pathFT_OHRQ nvarchar(MAX);
DECLARE @pathFT_ORCA nvarchar(MAX);


while @i < @loop
	begin
		set @i = @i + 1; --inicio do laço
		set @empresa = (select empresa from mbs..bases where baseid = @i)
		set @pathocrd = CONCAT(N'CREATE SYNONYM tpathocrd FOR ',@empresa,N'.DBO.ocrd');
		set @pathINV6 = CONCAT(N'CREATE SYNONYM tpathINV6 FOR ',@empresa,N'.DBO.INV6');
		set @pathOINV = CONCAT(N'CREATE SYNONYM tpathOINV FOR ',@empresa,N'.DBO.OINV');
		set @pathFT_OSEG = CONCAT(N'CREATE SYNONYM tpathFT_OSEG FOR ',@empresa,N'.DBO.[@FT_OSEG]');
		set @pathFT_OHRQ = CONCAT(N'CREATE SYNONYM tpathFT_OHRQ FOR ',@empresa,N'.DBO.[@FT_OHRQ]');
		set @pathFT_ORCA = CONCAT(N'CREATE SYNONYM tpathFT_ORCA FOR ',@empresa,N'.DBO.[@FT_ORCA]');
		set @pathOSLP = CONCAT(N'CREATE SYNONYM tpathOSLP FOR ',@empresa,N'.DBO.OSLP');
		set @pathOCRG = CONCAT(N'CREATE SYNONYM tpathOCRG FOR ',@empresa,N'.DBO.OCRG');
		set @pathOCNT = CONCAT(N'CREATE SYNONYM tpathOCNT FOR ',@empresa,N'.DBO.OCNT');

		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathocrd') > 0 BEGIN DROP SYNONYM tpathocrd; EXEC(@pathocrd); END ELSE BEGIN exec(@pathocrd); END
		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathINV6') > 0 BEGIN DROP SYNONYM tpathINV6; EXEC(@pathINV6); END ELSE BEGIN exec(@pathINV6); END
		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathOINV') > 0 BEGIN DROP SYNONYM tpathOINV; EXEC(@pathOINV); END ELSE BEGIN exec(@pathOINV); END
		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathFT_OSEG') > 0 BEGIN DROP SYNONYM tpathFT_OSEG; EXEC(@pathFT_OSEG); END ELSE BEGIN exec(@pathFT_OSEG); END
		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathFT_OHRQ') > 0 BEGIN DROP SYNONYM tpathFT_OHRQ; EXEC(@pathFT_OHRQ); END ELSE BEGIN exec(@pathFT_OHRQ); END
		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathFT_ORCA') > 0 BEGIN DROP SYNONYM tpathFT_ORCA; EXEC(@pathFT_ORCA); END ELSE BEGIN exec(@pathFT_ORCA); END
		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathOSLP') > 0 BEGIN DROP SYNONYM tpathOSLP; EXEC(@pathOSLP); END ELSE BEGIN exec(@pathOSLP); END
		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathOCRG') > 0 BEGIN DROP SYNONYM tpathOCRG; EXEC(@pathOCRG); END ELSE BEGIN exec(@pathOCRG); END
		IF (select COUNT(NAME) from sys.synonyms WHERE NAME = 'tpathOCNT') > 0 BEGIN DROP SYNONYM tpathOCNT; EXEC(@pathOCNT); END ELSE BEGIN exec(@pathOCNT); END


		INSERT INTO [mbs].[dbo].[CONTAS_A_RECEBER]
			Select 
				  'SEGMENTO' AS SEGMENTO
				 , 'OKAY' AS SEMAFORO
				 , upper(@empresa) as BASE
				 , replace(upper(@empresa),'SBO','') AS EMPRESA
				 ,  (Convert (NvarChar (10),v.Serial) + '-' + Convert (VarChar(2), i.InstlmntID)) As Documento 
				 , c.DocEntry As DocumentoSAP
				 , FormaCobranca=(Case When IsNull(v.Indicator, '-1')='CH' Then 'CHEQUE DEVOLVIDO' Else 'CARTEIRA' End)
				 , ((i.InsTotal -i.PaidToDate)) As ValorTitulo
				 , CAST(v.TaxDate AS DATE) As Emissao
				 , CAST(i.DueDate AS DATE) As Vencimento
				 , DateDiff (D, i.DueDate , GetDate()) As Atraso
				 , UPPER(COALESCE((Select s.[Name] From tpathFT_OSEG s where s.Code = c.U_Segmento),'### INDEFINIDA ###')) As Categoria
				 , UPPER((Select p.Slpname From tpathOSLP p Where p.SlpCode = c.SlpCode)) As Vendedor
				 , IsNull ((Select h.Name From tpathFT_OHRQ h Where h.Code = (Select r.U_Hierarquia From tpathFT_ORCA r Where r.Code = c.SlpCode)), 'NAO DEFINIDO') As Supervisor
				 , UPPER((Select g.Groupname From tpathOCRG g Where g.GroupCode = c.GroupCode)) As Grupo
				 , c.CardCode As CodigoCliente
				 , Upper (c.CardName) As RazaoSocial
				 , UPPER(IsNull((Select IsNull(t.Name,'') From tpathOCNT t Where t.AbsId = c.County),'')) As Municipio
				 , c.State1 As Estado
				 , COALESCE(c.Phone1,'') As Fone
				 , 'Fonte: SAP | Departamento de Inteligencia | ' + Format(CURRENT_TIMESTAMP,'dd.MMM.yyyy HH:mm') + ' @rpastudio' AS [VERSION]
			From
				tpathocrd c inner Join 
				tpathOINV v On c.CardCode = v.CardCode inner Join
				tpathINV6 i On v.DocEntry = i.DocEntry  
			Where
				i.Status = 'O' AND
				(i.InsTotal - i.PaidToDate) > 1 AND
				DateDiff (D, i.DueDate , GetDate()) < 100 AND
				DateDiff (D, i.DueDate , GetDate()) > 5
	end;	  
	

UPDATE 
	[Mbs].[dbo].[CONTAS_A_RECEBER] 
	SET 
		SEMAFORO = 
			CASE
				WHEN ATRASO < 5 THEN '9OKAY (-5D)'
				WHEN ATRASO >= 5 AND ATRASO <= 15 THEN '1ALERTA (+5D)'
				WHEN ATRASO >= 16 AND ATRASO <= 40 THEN '2PROT. (+15D)'
				WHEN ATRASO >= 41 AND ATRASO <= 60 THEN '3NEGOC. (+40D)'
				WHEN ATRASO >= 61 AND ATRASO <= 90 THEN '4CRITIC. (+60D)'
				ELSE '5PERDA (+90D)'
			END,

		EMPRESA = 
				CASE
					WHEN EMPRESA = 'A7' THEN 'A7'
					WHEN EMPRESA = 'ChuaChua' THEN 'CHUA CHUA'
					WHEN EMPRESA = 'DL' THEN 'DL'
					WHEN EMPRESA = 'EmpLREZENDE' THEN 'LREZENDE'
					WHEN EMPRESA = 'EmpreendimentosRZ' THEN 'RZ'
					WHEN EMPRESA = 'EmpreendimentosSM' THEN 'SM'
					WHEN EMPRESA = 'Equilibrium' THEN 'EQUILIBRIUM'
					WHEN EMPRESA = 'GDSMARCAS' THEN 'BELA CARIOCA'
					WHEN EMPRESA = 'GdsMarcasEU' THEN 'GDS EUROPA'
					WHEN EMPRESA = 'Infinix1' THEN 'INFINIX 1'
					WHEN EMPRESA = 'Infinix2' THEN 'INFINIX 2'
					WHEN EMPRESA = 'InnovaBR' THEN 'INNOVA BR'
					WHEN EMPRESA = 'InnovaLab' THEN 'INNOVA LAB'
					WHEN EMPRESA = 'InnovaPA' THEN 'INNOVA PA'
					WHEN EMPRESA = 'LCA' THEN 'LCA'
					WHEN EMPRESA = 'ML' THEN 'ML'
					WHEN EMPRESA = 'MW' THEN 'MW'
					WHEN EMPRESA = 'NutGynMatriz' THEN 'NPH'
					WHEN EMPRESA = 'NutriexInd' THEN 'NUTRIEX COSMETICOS'
					WHEN EMPRESA = 'Oralls' THEN 'ORALLS'
					WHEN EMPRESA = 'RZLaboratorios' THEN 'RZ PORTUGAL'
					WHEN EMPRESA = 'Vidafarma' THEN 'VDM'
					WHEN EMPRESA = 'VZ' THEN 'VZ'
					WHEN EMPRESA = 'Z9' THEN 'Z9'
					ELSE EMPRESA
				END,

		SEGMENTO = 
				CASE
					WHEN BASE IN ('SBOINNOVABR', 'SBOINNOVAPA', 'SBOVIDAFARMA')	THEN 'RENNOVA'
					WHEN BASE IN ('SBOA7', 'SBODL', 'SBOEQUILIBRIUM', 'SBOINNOVALAB', 'SBOLICITAACAO') THEN 'LICITACAO'
					WHEN BASE IN ('SBOML', 'SBOMW', 'SBONUTGYNMATRIZ', 'SBONUTRIEXIND','SBOLCA','SboChuaChua','SBOGDSMARCAS') THEN 'NUTRIEX'					
					WHEN BASE IN ('SBOEmpLREZENDE', 'SBOEmpreendimentosRZ', 'SBOEmpreendimentosSM', 'SBOInfinix1','SBOInfinix2','SBOVZ') THEN 'HOLDING'					
					WHEN BASE IN ('SBOGdsMarcasEU', 'SBOOralls', 'SBORZLaboratorios', 'SBOZ9') THEN 'DIST_LUCROS'					
					ELSE 'GERAL'
				END,
		
		GRUPO = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(GRUPO,'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' '),

		RAZAOSOCIAL = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(RAZAOSOCIAL,'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' '),

		CATEGORIA = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(CATEGORIA,'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' '),

		MUNICIPIO = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(MUNICIPIO,'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' ')
		

--SELECT 
--	 [SEGMENTO]
--	,[SEMAFORO]
--    ,[BASE]
--    ,[EMPRESA]
--    ,[Documento]
--    ,[DocumentoSAP]
--    ,[FormaCobranca]
--    ,[Emissao]
--    ,[Vencimento]
--    ,[Atraso]
--    ,[Categoria]
--    ,[Vendedor]
--    ,[Supervisor]
--    ,[Grupo]
--    ,[CodigoCliente]
--    ,[RazaoSocial]
--    ,[Municipio]
--    ,[Estado]
--    ,[Fone]
--    ,[ValorTitulo]
--	,[Version]
--FROM 
--	[Mbs].[dbo].[CONTAS_A_RECEBER]
--WHERE
--	GRUPO NOT LIKE '%ENT%BLICO%' AND
--	GRUPO NOT LIKE '%ENT%ESTADUAL%' AND
--	GRUPO NOT LIKE '%LIGADA%' AND
--	GRUPO NOT LIKE '%EXTERIOR%' AND 
--	SUPERVISOR NOT LIKE '%COLIGADA%' AND
--	SEGMENTO NOT LIKE 'HOLDING' AND
--	SEGMENTO NOT LIKE 'DIST_LUCROS' AND
--	ESTADO NOT LIKE 'EX'