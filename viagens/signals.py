from django.db.models.signals import post_delete, post_save
from django.dispatch import receiver

from .models import Oficio, Trecho


@receiver(post_save, sender=Trecho)
@receiver(post_delete, sender=Trecho)
def atualizar_destino_oficio(sender, instance: Trecho, **kwargs) -> None:
    if not instance.oficio_id:
        return
    oficio = Oficio.objects.filter(id=instance.oficio_id).first()
    if not oficio:
        return
    oficio.save(update_fields=["destino"])
