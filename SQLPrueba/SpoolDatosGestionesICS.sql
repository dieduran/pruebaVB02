SET HEADING  OFF
SET FEEDBACK OFF
SET ECHO OFF
SET PAGES     0
SET COLSEP   ";"
SET VERIFY OFF
SET LINESIZE 45
 
--ALTER SESSION SET NLS_DATE_FORMAT = "DD/MM/YYYY HH24:MI:SS";
--ALTER SESSION SET NLS_NUMERIC_CHARACTERS = ". ";   

--10.0.0.68
--SPOOL \\tn01\ftp$\cr\tsft\TelefonosActualizacion\ClientesTelefonos.csv
SPOOL C:\temp\ClientesGestiones.csv

SELECT contacto||';'||to_char(fecha,'dd/mm/yyyy hh24:mi:ss')||';'||action_id||';'||response_id||';'||contact_id||';'||replace(replace(observaciones,';',''),'''','')
from (
  with matriz as (select distinct customer_id, action_id, response_id, contact_id
                    from custom.matriz_tipos_contactos
                   where customer_id = 'NARANJA' and type_result = 'CONTACTO' and ef_nef = 'EFECTIVO')
  SELECT contacto, tipodoc, nrodoc, h.history_date fecha , h.action_id, h.response_id, h.contact_id,
         replace(replace(h.observations,chr(10),''),';',',')observaciones,
         row_number() Over (partition by h.customer_id, h.identity_code order by history_date desc) fila
  FROM custom.history_custom h, matriz m,
       (select distinct sc.customer_id, sc.identity_code, contacto, tipodoc, nrodoc
          from custom.tmp_contacto_tsft tc, custom.strategy_custom sc
        where trim(tc.tipodoc)||tc.nrodoc = sc.identity_code
          and sc.customer_id = 'NARANJA' and copy_date >= add_months(sysdate,-1)
          and exists (select customer_id, strategy from custom.grupo_estrategias_ics g
                        where sc.customer_id = g.customer_id and sc.strategy_id = g.strategy
                          and group_strategy = 'PRELEGAL')) asig
  WHERE h.customer_id = m.customer_id AND h.user_id != 'SYSTEM'
    AND m.customer_id = h.customer_id AND m.action_id = h.action_id
    AND m.response_id = h.response_id AND m.contact_id = h.contact_id
    and h.customer_id = asig.customer_id and h.identity_code = asig.identity_code
    AND h.history_date >= add_months(sysdate,-1)
)
where fila = 1

SPOOL OFF

EXIT;
