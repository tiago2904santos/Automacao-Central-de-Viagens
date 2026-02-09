from django.urls import path

from . import views


urlpatterns = [
    path("", views.dashboard_home, name="dashboard_home"),
    path("dashboard/", views.dashboard_home, name="dashboard_home_alt"),
    path("dashboard/data/", views.dashboard_data_api, name="dashboard_data_api"),
    path("oficios/novo/", views.formulario, name="formulario"),
    path("oficios/etapa-2/", views.oficio_step2, name="oficio_step2"),
    path("oficios/etapa-3/", views.oficio_step3, name="oficio_step3"),
    path("oficios/etapa-4/", views.oficio_step4, name="oficio_step4"),
    path(
        "oficios/<int:oficio_id>/rascunho/",
        views.oficio_draft_resume,
        name="oficio_draft_resume",
    ),
    path(
        "oficios/<int:oficio_id>/editar/etapa-1/",
        views.oficio_edit_step1,
        name="oficio_edit_step1",
    ),
    path(
        "oficios/<int:oficio_id>/editar/etapa-2/",
        views.oficio_edit_step2,
        name="oficio_edit_step2",
    ),
    path(
        "oficios/<int:oficio_id>/editar/etapa-3/",
        views.oficio_edit_step3,
        name="oficio_edit_step3",
    ),
    path(
        "oficios/<int:oficio_id>/editar/etapa-4/",
        views.oficio_edit_step4,
        name="oficio_edit_step4",
    ),
    path(
        "oficios/<int:oficio_id>/editar/salvar/",
        views.oficio_edit_save,
        name="oficio_edit_save",
    ),
    path(
        "oficios/<int:oficio_id>/editar/cancelar/",
        views.oficio_edit_cancel,
        name="oficio_edit_cancel",
    ),

    path("oficios/<int:oficio_id>/download-docx/",views.oficio_download_docx,name="oficio_download_docx"),
    path("oficios/<int:oficio_id>/download-pdf/", views.oficio_download_pdf, name="oficio_download_pdf"),
    path("api/validacoes/resultado/", views.validacao_resultado, name="validacao_resultado"),


    path("viajantes/novo/", views.viajante_cadastro, name="viajante_cadastro"),
    path("viajantes/<int:viajante_id>/editar/", views.viajante_editar, name="viajante_editar"),
    path("viajantes/", views.viajantes_lista, name="viajantes_lista"),
    path("veiculos/novo/", views.veiculo_cadastro, name="veiculo_cadastro"),
    path("veiculos/<int:veiculo_id>/editar/", views.veiculo_editar, name="veiculo_editar"),
    path("veiculos/", views.veiculos_lista, name="veiculos_lista"),
    path("oficios/", views.oficios_lista, name="oficios_lista"),
    path("oficios/<int:oficio_id>/excluir/", views.oficio_excluir, name="oficio_excluir"),
    path("configuracoes/oficio/", views.configuracoes_oficio, name="config_oficio"),
    path("api/viajantes/", views.viajantes_api, name="viajantes_api"),
    path("api/veiculo/", views.veiculo_api, name="veiculo_api"),
    path("api/cidades/", views.cidades_api, name="cidades_api"),
    path("api/ufs/", views.ufs_api, name="ufs_api"),
    path("api/cidades-busca/", views.cidades_busca_api, name="cidades_busca_api"),
    path("api/servidores/", views.servidores_api, name="servidores_api"),
    path("api/assinantes/", views.assinantes_api, name="assinantes_api"),
    path("api/motoristas/", views.motoristas_api, name="motoristas_api"),
    path("api/cep/<str:cep>/", views.cep_api, name="cep_api"),
    path("api/servidores/<int:servidor_id>/", views.servidor_detail_api, name="servidor_detail_api"),
    path("api/motoristas/<int:motorista_id>/", views.motorista_detail_api, name="motorista_detail_api"),
    path("api/veiculos/", views.veiculos_busca_api, name="veiculos_busca_api"),
    path("api/veiculos/<int:veiculo_id>/", views.veiculo_detail_api, name="veiculo_detail_api"),
    path("api/oficio-motorista-referencia/", views.oficio_motorista_referencia, name="oficio_motorista_referencia"),
    path("cargos/criar/", views.cargo_criar, name="cargo_criar"),
    path("modal/viajantes/novo/", views.modal_viajante_form, name="modal_viajante_form"),
    path("modal/veiculos/novo/", views.modal_veiculo_form, name="modal_veiculo_form"),
]
