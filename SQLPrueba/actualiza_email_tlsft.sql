

delete from cr.email e where exists (select 1 from cr.def_abo_email d where d.nrodoc = e.contacto);


insert into cr.email (contacto,email,valido)
  select nrodoc,email,'S' from cr.def_abo_email; 

COMMIT;

EXIT;
 

