from django.apps import AppConfig


class ViagensConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "viagens"

    def ready(self):
        from . import signals  # noqa: F401
