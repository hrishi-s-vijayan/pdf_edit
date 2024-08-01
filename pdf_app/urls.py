from django.urls import path
from .views import generate_pdf

urlpatterns = [
    path('generate/', generate_pdf, name='generate_pdf')
]
