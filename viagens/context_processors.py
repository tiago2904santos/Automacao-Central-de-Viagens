from django.conf import settings
from django.utils import timezone


def mask_settings(request):
    return {
        "PROTOCOLO_MASK": getattr(settings, "PROTOCOLO_MASK", None),
        "CURRENT_YEAR": str(timezone.localdate().year),
    }
