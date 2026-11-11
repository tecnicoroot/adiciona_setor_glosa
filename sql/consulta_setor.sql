select g.nr_doc_convenio,
       obter_nome_setor(g.cd_setor_atendimento) setor
from (select distinct nr_doc_convenio, cd_setor_atendimento
        from material_atend_paciente
       where nr_doc_convenio in ({guias})) g