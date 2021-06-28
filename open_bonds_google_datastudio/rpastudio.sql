SELECT 
	 [SEGMENTO]
	,[SEMAFORO]
    ,[BASE]
    ,[EMPRESA]
    ,[Documento]
    ,[DocumentoSAP]
    ,[FormaCobranca]
    ,[Emissao]
    ,[Vencimento]
    ,[Atraso]
    ,[Categoria]
    ,[Vendedor]
    ,[Supervisor]
    ,[Grupo]
    ,[CodigoCliente]
    ,[RazaoSocial]
    ,[Municipio]
    ,[Estado]
    ,[Fone]
    ,[ValorTitulo]
	,[Version]
FROM 
	[Mbs].[dbo].[CONTAS_A_RECEBER]
WHERE
	GRUPO NOT LIKE '%ENT%BLICO%' AND
	GRUPO NOT LIKE '%ENT%ESTADUAL%' AND
	GRUPO NOT LIKE '%LIGADA%' AND
	GRUPO NOT LIKE '%EXTERIOR%' AND 
	SUPERVISOR NOT LIKE '%COLIGADA%' AND
	SEGMENTO NOT LIKE 'HOLDING' AND
	SEGMENTO NOT LIKE 'DIST_LUCROS' AND
	ESTADO NOT LIKE 'EX'