from django.db.utils import OperationalError, ProgrammingError

from viagens.models import OficioConfig


def get_oficio_config() -> OficioConfig:
    try:
        config = OficioConfig.objects.order_by("id").first()
    except (OperationalError, ProgrammingError):
        return OficioConfig()
    if config:
        return config
    return OficioConfig()
