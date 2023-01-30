

delete from cr.cpc_ics_contacto e where exists (select 1 from cr.Def_Abo_Gestion d where d.contacto = e.contacto);


insert into cr.cpc_ics_contacto (contacto,fecha, action_id, response_id, contact_id, observaciones)
  select contacto,fecha, action_id, response_id, contact_id, observaciones from cr.Def_Abo_Gestion; 

COMMIT;

EXIT;
 

