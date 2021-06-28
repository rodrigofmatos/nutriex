sqlcmd -S "SRV2-SAP" -d DBReport -U sa -P @SapN22RwDt@ -i "f:\bot\rpa001\sql_setnocount.sql","f:\bot\rpa001\rpa001.sql" -o "f:\bot\rpa001\rpa001.csv"  -s";" -W -f 65001 
python f:\bot\rpa001\rpa001.py
