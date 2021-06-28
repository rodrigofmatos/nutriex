sqlcmd -S "SRV2-SAP" -d DBReport -U sa -P @SapN22RwDt@ -i "f:\bot\rpaStudio\sql_setnocount.sql","f:\bot\rpaStudio\rpastudio.sql" -o "f:\bot\rpaStudio\rpastudio.csv"  -s"," -W -f 65001 
python f:\bot\rpaStudio\rpastudio.py
