SET HEADING  OFF
SET FEEDBACK OFF
SET ECHO OFF
SET PAGES     0
SET COLSEP   ";"
SET VERIFY OFF
SET LINESIZE 45
 
ALTER SESSION SET NLS_DATE_FORMAT = "DD/MM/YYYY HH24:MI:SS";
ALTER SESSION SET NLS_NUMERIC_CHARACTERS = ". ";   

--10.0.0.68
--SPOOL \\tn01\ftp$\cr\tsft\TelefonosActualizacion\ClientesTelefonos.csv
SPOOL C:\temp\ClientesTelefonos.csv

select to_char(contacto)||';'||trim(tipodoc)||';'||to_char(nrodoc)||';'||trim(tipodom)||';'||trim(teledisc)||';'||trim(telefono) from custom.def_abo_tel_ics;
SPOOL OFF

--SPOOL \\tn01\ftp$\cr\tsft\TelefonosActualizacion\ClientesEmail.csv
SPOOL C:\temp\ClientesEmail.csv
select * from (
  select contacto ||';' || 
         custom.fnc_split(custom.fnc_devuelve_emails('NARANJA',trim(tipodoc)||nrodoc)||'|','|',1) dir
  from custom.tmp_contacto_tsft tt
union
  select contacto ||';' || 
         custom.fnc_split(custom.fnc_devuelve_emails('NARANJA',trim(tipodoc)||nrodoc)||'|','|',2)
  from custom.tmp_contacto_tsft tt
union
  select contacto ||';' || 
         custom.fnc_split(custom.fnc_devuelve_emails('NARANJA',trim(tipodoc)||nrodoc)||'|','|',3)
  from custom.tmp_contacto_tsft tt
union
  select contacto ||';' || 
         custom.fnc_split(custom.fnc_devuelve_emails('NARANJA',trim(tipodoc)||nrodoc)||'|','|',4)
  from custom.tmp_contacto_tsft tt
union
  selectcontacto ||';' || 
         custom.fnc_split(custom.fnc_devuelve_emails('NARANJA',trim(tipodoc)||nrodoc)||'|','|',5)
  from custom.tmp_contacto_tsft tt
)
where dir is not null
order by 1;
 
SPOOL OFF

EXIT;
