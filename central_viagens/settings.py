import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = "django-insecure-change-me"
DEBUG = True
_allowed_hosts_env = os.getenv("DJANGO_ALLOWED_HOSTS", "")
ALLOWED_HOSTS: list[str] = [
    host.strip() for host in _allowed_hosts_env.split(",") if host.strip()
]

if DEBUG:
    for host in ("localhost", "127.0.0.1", "testserver"):
        if host not in ALLOWED_HOSTS:
            ALLOWED_HOSTS.append(host)
else:
    ALLOWED_HOSTS = [host for host in ALLOWED_HOSTS if host != "*"]

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "viagens",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "central_viagens.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
                "viagens.context_processors.mask_settings",
            ],
        },
    }
]

WSGI_APPLICATION = "central_viagens.wsgi.application"

DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": BASE_DIR / "db.sqlite3",
    }
}

AUTH_PASSWORD_VALIDATORS = [
    {
        "NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.MinimumLengthValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.CommonPasswordValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.NumericPasswordValidator",
    },
]

LANGUAGE_CODE = "pt-br"
TIME_ZONE = "America/Sao_Paulo"
USE_I18N = True
USE_TZ = True

STATIC_URL = "static/"
MEDIA_URL = "/media/"
MEDIA_ROOT = BASE_DIR / "media"

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

OFICIO_ORGAO_DESTINO_PADRAO = os.getenv("OFICIO_ORGAO_DESTINO_PADRAO", "DGA")
OFICIO_CHEFIA = {
    "nome_chefia": os.getenv("OFICIO_CHEFIA_NOME", ""),
    "cargo_chefia": os.getenv("OFICIO_CHEFIA_CARGO", ""),
}

PROTOCOLO_MASK = None
