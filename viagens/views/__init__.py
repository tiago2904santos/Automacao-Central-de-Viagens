from . import _shared as _shared_module
from ._shared import *  # noqa: F401,F403
from ._shared import _get_cargo_choices, _normalizar_cargo_key
from .api import *  # noqa: F401,F403
from .cadastros import *  # noqa: F401,F403
from .configuracoes import *  # noqa: F401,F403
from .dashboard import *  # noqa: F401,F403
from .documentos import *  # noqa: F401,F403
from .oficios import *  # noqa: F401,F403
from .plano_trabalho import *  # noqa: F401,F403
from .roteiros import *  # noqa: F401,F403


def _sync_pdf_dependencies() -> None:
    _shared_module.build_oficio_docx_and_pdf_bytes = build_oficio_docx_and_pdf_bytes
    _shared_module.docx_bytes_to_pdf_bytes = docx_bytes_to_pdf_bytes


def termo_autorizacao_download_pdf(request, termo_id: int):
    _sync_pdf_dependencies()
    return _shared_module.termo_autorizacao_download_pdf(request, termo_id)


def oficio_download_termo_autorizacao_pdf(request, oficio_id: int):
    _sync_pdf_dependencies()
    return _shared_module.oficio_download_termo_autorizacao_pdf(request, oficio_id)


def plano_trabalho_download_pdf(request, oficio_id: int):
    _sync_pdf_dependencies()
    return _shared_module.plano_trabalho_download_pdf(request, oficio_id)


def ordem_servico_download_pdf(request, oficio_id: int):
    _sync_pdf_dependencies()
    return _shared_module.ordem_servico_download_pdf(request, oficio_id)


def oficio_download_plano_trabalho_pdf(request, oficio_id: int):
    _sync_pdf_dependencies()
    return _shared_module.oficio_download_plano_trabalho_pdf(request, oficio_id)


def oficio_download_ordem_servico_pdf(request, oficio_id: int):
    _sync_pdf_dependencies()
    return _shared_module.oficio_download_ordem_servico_pdf(request, oficio_id)


def oficio_download_pdf(request, oficio_id: int):
    _sync_pdf_dependencies()
    return _shared_module.oficio_download_pdf(request, oficio_id)


def api_cidades_por_estado(request):
    return cidades_api(request)
