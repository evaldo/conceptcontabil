use fluxocaixa;

create login escolateste with password = '1234';
create user escolateste with default_schema = dbo;

GRANT SELECT, UPDATE, DELETE, INSERT ON T_CLSSF_PLANO_CONTA TO  escolateste; 
GRANT SELECT, UPDATE, DELETE, INSERT ON T_CNRIO_IMPRT_ARQV TO  escolateste; 
GRANT SELECT, UPDATE, DELETE, INSERT ON T_CNRO_EXPRT_ARQV TO  escolateste; 
GRANT SELECT ON T_DMSAO_TEMPO TO  escolateste; 
GRANT SELECT, UPDATE, DELETE, INSERT ON T_LISTA_PLVR_EXCD TO  escolateste;

GRANT SELECT ON VW_FLUXO_CAIXA_POR_DS_MES TO  escolateste; 
GRANT SELECT ON VW_TRIM_SAIDA_CAIXA TO  escolateste; 
