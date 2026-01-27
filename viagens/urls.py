from django.urls import path

from . import views


urlpatterns = [
    path("", views.dashboard_home, name="dashboard_home"),
    path("dashboard/", views.dashboard_home, name="dashboard_home_alt"),
    path("dashboard/data/", views.dashboard_data_api, name="dashboard_data_api"),
    path("oficios/novo/", views.formulario, name="formulario"),
    path("oficios/etapa-2/", views.oficio_step2, name="oficio_step2"),
    path("oficios/etapa-3/", views.oficio_step3, name="oficio_step3"),
    path("documento/<int:oficio_id>/", views.documento_oficio, name="documento_oficio"),
    path("viajantes/novo/", views.viajante_cadastro, name="viajante_cadastro"),
    path("viajantes/", views.viajantes_lista, name="viajantes_lista"),
    path("veiculos/novo/", views.veiculo_cadastro, name="veiculo_cadastro"),
    path("veiculos/", views.veiculos_lista, name="veiculos_lista"),
    path("oficios/", views.oficios_lista, name="oficios_lista"),
    path("api/viajantes/", views.viajantes_api, name="viajantes_api"),
    path("api/veiculo/", views.veiculo_api, name="veiculo_api"),
    path("api/cidades/", views.cidades_api, name="cidades_api"),
    path("api/ufs/", views.ufs_api, name="ufs_api"),
    path("api/cidades-busca/", views.cidades_busca_api, name="cidades_busca_api"),
    path("api/servidores/", views.servidores_api, name="servidores_api"),
    path("api/motoristas/", views.motoristas_api, name="motoristas_api"),
    path("api/servidores/<int:servidor_id>/", views.servidor_detail_api, name="servidor_detail_api"),
    path("api/motoristas/<int:motorista_id>/", views.motorista_detail_api, name="motorista_detail_api"),
    path("api/veiculos/", views.veiculos_busca_api, name="veiculos_busca_api"),
    path("api/veiculos/<int:veiculo_id>/", views.veiculo_detail_api, name="veiculo_detail_api"),
    path("modal/viajantes/novo/", views.modal_viajante_form, name="modal_viajante_form"),
    path("modal/veiculos/novo/", views.modal_veiculo_form, name="modal_veiculo_form"),
]
