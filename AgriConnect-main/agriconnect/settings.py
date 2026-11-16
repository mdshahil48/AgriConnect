import os
from pathlib import Path
from dotenv import load_dotenv
import dj_database_url

# Build paths inside the project like this: BASE_DIR / 'subdir'.
BASE_DIR = Path(__file__).resolve().parent.parent

# Load .env file if it exists
load_dotenv(os.path.join(BASE_DIR, '.env'))

# Quick-start development settings - unsuitable for production
DEBUG = os.getenv('DEBUG', 'False').lower() in ('true', '1', 't')

# SECRET_KEY: read from env var, fall back to dev key only when DEBUG is True
SECRET_KEY = os.getenv('SECRET_KEY')
if not SECRET_KEY:
    if DEBUG:
        SECRET_KEY = 'django-insecure-agriconnect-kisan-setu-sih-2026-secret-key'
    else:
        raise ValueError("The SECRET_KEY environment variable must be set in production when DEBUG=False.")

# Host and Domain configuration
raw_allowed_hosts = os.getenv('ALLOWED_HOSTS', 'localhost,127.0.0.1,testserver')
ALLOWED_HOSTS = [h.strip() for h in raw_allowed_hosts.split(',') if h.strip()]

# Automatic Render deployment hostname support
render_hostname = os.getenv('RENDER_EXTERNAL_HOSTNAME')
if render_hostname and render_hostname not in ALLOWED_HOSTS:
    ALLOWED_HOSTS.append(render_hostname)

# Always allow .onrender.com domain pattern
if '.onrender.com' not in ALLOWED_HOSTS:
    ALLOWED_HOSTS.append('.onrender.com')

if DEBUG:
    ALLOWED_HOSTS.extend(['*', 'testserver', 'localhost', '127.0.0.1'])
ALLOWED_HOSTS = list(dict.fromkeys(ALLOWED_HOSTS))

# CSRF Trusted Origins (required for Django 4.0+ over HTTPS on Render)
csrf_origins_env = os.getenv('CSRF_TRUSTED_ORIGINS', '')
CSRF_TRUSTED_ORIGINS = [origin.strip() for origin in csrf_origins_env.split(',') if origin.strip()]
if render_hostname:
    CSRF_TRUSTED_ORIGINS.append(f'https://{render_hostname}')
if 'https://*.onrender.com' not in CSRF_TRUSTED_ORIGINS:
    CSRF_TRUSTED_ORIGINS.append('https://*.onrender.com')
CSRF_TRUSTED_ORIGINS = list(dict.fromkeys(CSRF_TRUSTED_ORIGINS))

# Application definition
INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',

    # Third-party apps
    'rest_framework',

    # Local apps
    'apps.accounts',
    'apps.facilitators',
    'apps.farmers',
    'apps.products',
    'apps.marketplace',
    'apps.orders',
    'apps.market_intelligence',
    'apps.logistics',
    'apps.analytics',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'agriconnect.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [os.path.join(BASE_DIR, 'templates')],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
                'apps.orders.cart.cart_context',
            ],
        },
    },
]

WSGI_APPLICATION = 'agriconnect.wsgi.application'

# Database
# Default to SQLite for local development; switches to PostgreSQL (Neon) via DATABASE_URL
database_url = os.getenv("DATABASE_URL")
if database_url:
    is_neon_or_cloud = "neon.tech" in database_url or "sslmode=require" in database_url or not DEBUG
    DATABASES = {
        "default": dj_database_url.config(
            default=database_url,
            conn_max_age=600,
            conn_health_checks=True,
            ssl_require=is_neon_or_cloud,
        )
    }
else:
    DATABASES = {
        "default": {
            "ENGINE": "django.db.backends.sqlite3",
            "NAME": BASE_DIR / "db.sqlite3",
        }
    }

# Custom User Model
AUTH_USER_MODEL = 'accounts.CustomUser'

# Password validation
AUTH_PASSWORD_VALIDATORS = [
    {
        'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator',
        'OPTIONS': {'min_length': 6},
    },
    {
        'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator',
    },
]

# Internationalization
LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'Asia/Kolkata'
USE_I18N = True
USE_TZ = True

# Static files (CSS, JavaScript, Images)
STATIC_URL = '/static/'
STATICFILES_DIRS = [os.path.join(BASE_DIR, 'static')]
STATIC_ROOT = os.path.join(BASE_DIR, 'staticfiles')

# Media files (Produce photos, documents)
MEDIA_URL = '/media/'
MEDIA_ROOT = os.path.join(BASE_DIR, 'media')

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# WhiteNoise for production static file serving (Django 4.2+ STORAGES and legacy fallback)
STORAGES = {
    "default": {
        "BACKEND": "django.core.files.storage.FileSystemStorage",
    },
    "staticfiles": {
        "BACKEND": "whitenoise.storage.CompressedManifestStaticFilesStorage",
    },
}
STATICFILES_STORAGE = 'whitenoise.storage.CompressedManifestStaticFilesStorage'

# Authentication redirects
LOGIN_URL = 'login'
LOGIN_REDIRECT_URL = 'dashboard_router'
LOGOUT_REDIRECT_URL = 'home'

# Rest Framework Configuration
REST_FRAMEWORK = {
    'DEFAULT_PERMISSION_CLASSES': [
        'rest_framework.permissions.IsAuthenticatedOrReadOnly',
    ],
    'DEFAULT_PAGINATION_CLASS': 'rest_framework.pagination.PageNumberPagination',
    'PAGE_SIZE': 12,
}

# Open Government Data (data.gov.in) Agmarknet Mandi API Configuration
DATA_GOV_IN_API_KEY = os.getenv('DATA_GOV_IN_API_KEY', '579b464db66ec23bdd000001cdd3946e44ce4aad7209ff7b23ac571b')
DATA_GOV_IN_MANDI_RESOURCE_ID = os.getenv('DATA_GOV_IN_MANDI_RESOURCE_ID', '9ef84268-d588-465a-a308-a864a43d0070')
