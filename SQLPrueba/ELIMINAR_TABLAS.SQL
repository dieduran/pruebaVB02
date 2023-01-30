SET ECHO OFF
SET VERIFY OFF
SET TERM OFF
SET SERVEROUTPUT ON

BEGIN
delete from Cr.DEF_ABO_TEL_ICS;
delete from Cr.Def_abo_email;
delete from Cr.Def_Abo_Gestion;
END;
/
exit