select g.nr_doc_convenio,
       g.cd_material_tiss,
       to_char(g.dt_prescricao, 'dd/mm/yyyy hh24:mi:ss') dt_prescricao,
       obter_nome_setor(g.cd_setor_atendimento) setor
from (select distinct nr_doc_convenio, cd_material_tiss, dt_prescricao, cd_setor_atendimento
        from material_atend_paciente
       where nr_doc_convenio in ({guias})) g