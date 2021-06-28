sqlcmd -S "SRV2-SAP" -d DBReport -U sa -P @SapN22RwDt@ -i "f:\bot\rpa005\sql_setnocount.sql","f:\bot\rpa005\rpa005.sql" -o "f:\bot\rpa005\rpa005.csv"  -s"," -W -f 65001 
python f:\bot\rpa005\rpa005.py
