----------------------------------------------------------
--// TEMA: ROBO ORDEM PRODUCAO DIARIA - RPA005
--// -----------------------------------------------------
--// CRIADO POR: RODRIGO MATOS (CIENTISTA DADOS)
--// VERSAO: 3.0.0		
--// DATA: 04.MAIO.2020
--// SOLIC.: CRISTINA GUIMARAES (ANALISTA CUSTOS CONTROLADORIA)
--// -----------------------------------------------------

DROP TABLE MBS.[dbo].[RPA005]

CREATE TABLE MBS.[dbo].[RPA005](
	[OP] [nvarchar](20) NULL,
	[DATAOP] [date] NULL,
	[ITEMCOD] [nvarchar](20) NULL,
	[ITEMNOME] [nvarchar](100) NULL,
	[QTDEPAI] [float] NULL,
	[CUSTOMEDIOPAI] [float] NULL,
	[CODCOMP] [nvarchar](20) NULL,
	[DESCCOMP] [nvarchar](100) NULL,
	[CUSTOCOMP] [float] NULL,
	[QTDECOMP] [float] NULL,
	[CUSTOREAL_OP] [float] NULL,
	[VLRTOTCOMP] [float] NULL,
	[VLRTOTOP] [float] NULL
) ON [PRIMARY]

INSERT INTO MBS..RPA005

SELECT Distinct 
	CONCAT('NN_',FORMAT(T1.DOCNUM,'00000')) OP,
	CAST(T3.[DocDate] AS DATE) DATAOP, 
	T0.[ItemCode] ITEMCOD, 
	UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(T0.DSCRIPTION,'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' ')) ITEMNOME, 
	CAST(T0.[Quantity] AS FLOAT) QTDEPAI, 
	CAST(T0.Price AS FLOAT)  CUSTOMEDIOPAI,
	T4.ItemCode AS CODCOMP,
	UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(T4.[Dscription],'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' ')) DESCCOMP, 
	CAST(T4.Price AS FLOAT)  CUSTOCOMP,  
	CAST(T4.[Quantity] AS FLOAT)  QTDECOMP, 
	CAST((T4.lineTotal/T0.[Quantity]) AS FLOAT) CUSTOREAL_OP,
	CAST(T4.lineTotal AS FLOAT)  VLRTOTCOMP,
	CAST((select sum(I.lineTotal) from sbomw..IGE1 I where I.BaseRef = T4.BASEREF) AS FLOAT) VLRTOTOP
FROM 
	sbomw..IGN1 T0  INNER JOIN
	sbomw..OWOR T1 on T1.DocNum=T0.BaseRef 		   INNER JOIN 
	sbomw..OIGN T3 ON T0.DocEntry = T3.DocEntry    INNER JOIN 
	sbomw..IGE1 T4 ON T1.DOCNUM=T4.BASEREF 		   INNER JOIN 
	sbomw..OIGE T5 ON T5.DOCENTRY=T4.DOCENTRY 	   INNER JOIN 
	sbomw..WOR1 T6 ON T6.DOCENTRY=T1.DOCENTRY 	    
WHERE 
	format(t3.docdate,'yyyyMM') >= 201909


INSERT INTO MBS..RPA005

SELECT Distinct 
	CONCAT('IC_',FORMAT(T1.DOCNUM,'00000')) OP,
	CAST(T3.[DocDate] AS DATE) DATAOP, 
	T0.[ItemCode] ITEMCOD, 
	UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(T0.DSCRIPTION,'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' ')) ITEMNOME, 
	CAST(T0.[Quantity] AS FLOAT) QTDEPAI, 
	CAST(T0.Price AS FLOAT)  CUSTOMEDIOPAI,
	T4.ItemCode AS CODCOMP,
	UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(T4.[Dscription],'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' ')) DESCCOMP, 
	CAST(T4.Price AS FLOAT)  CUSTOCOMP,  
	CAST(T4.[Quantity] AS FLOAT)  QTDECOMP, 
	CAST((T4.lineTotal/T0.[Quantity]) AS FLOAT) CUSTOREAL_OP,
	CAST(T4.lineTotal AS FLOAT)  VLRTOTCOMP,
	CAST((select sum(I.lineTotal) from sbonutriexind..IGE1 I where I.BaseRef = T4.BASEREF) AS FLOAT) VLRTOTOP
FROM 
	sbonutriexind..IGN1 T0  INNER JOIN
	sbonutriexind..OWOR T1 on T1.DocNum=T0.BaseRef 		   INNER JOIN 
	sbonutriexind..OIGN T3 ON T0.DocEntry = T3.DocEntry    INNER JOIN 
	sbonutriexind..IGE1 T4 ON T1.DOCNUM=T4.BASEREF 		   INNER JOIN 
	sbonutriexind..OIGE T5 ON T5.DOCENTRY=T4.DOCENTRY 	   INNER JOIN 
	sbonutriexind..WOR1 T6 ON T6.DOCENTRY=T1.DOCENTRY 	    
WHERE 
	format(t3.docdate,'yyyyMM') >= 201909


INSERT INTO MBS..RPA005

SELECT Distinct 
	CONCAT('NH_',FORMAT(T1.DOCNUM,'00000')) OP,
	CAST(T3.[DocDate] AS DATE) DATAOP, 
	T0.[ItemCode] ITEMCOD, 
	UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(T0.DSCRIPTION,'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' ')) ITEMNOME, 
	CAST(T0.[Quantity] AS FLOAT) QTDEPAI, 
	CAST(T0.Price AS FLOAT)  CUSTOMEDIOPAI,
	T4.ItemCode AS CODCOMP,
	UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(T4.[Dscription],'É','E'),'Ê','E'),'Á','A'),'Â','A'),'Ã','A'),'À','A'),'Ç','C'),'Ú','U'),'Í','I'),'Ó','O'),'Õ','O'),'Ô','O'),',',' ')) DESCCOMP, 
	CAST(T4.Price AS FLOAT)  CUSTOCOMP,  
	CAST(T4.[Quantity] AS FLOAT)  QTDECOMP, 
	CAST((T4.lineTotal/T0.[Quantity]) AS FLOAT) CUSTOREAL_OP,
	CAST(T4.lineTotal AS FLOAT)  VLRTOTCOMP,
	CAST((select sum(I.lineTotal) from sbonutgynmatriz..IGE1 I where I.BaseRef = T4.BASEREF) AS FLOAT) VLRTOTOP
FROM 
	sbonutgynmatriz..IGN1 T0  INNER JOIN
	sbonutgynmatriz..OWOR T1 on T1.DocNum=T0.BaseRef 		   INNER JOIN 
	sbonutgynmatriz..OIGN T3 ON T0.DocEntry = T3.DocEntry    INNER JOIN 
	sbonutgynmatriz..IGE1 T4 ON T1.DOCNUM=T4.BASEREF 		   INNER JOIN 
	sbonutgynmatriz..OIGE T5 ON T5.DOCENTRY=T4.DOCENTRY 	   INNER JOIN 
	sbonutgynmatriz..WOR1 T6 ON T6.DOCENTRY=T1.DOCENTRY 	    
WHERE 
	format(t3.docdate,'yyyyMM') >= 201909
	