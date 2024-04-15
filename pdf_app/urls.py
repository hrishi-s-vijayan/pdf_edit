from django.urls import path
from .views import generate_pdf, modify_pdf

urlpatterns = [
    path('generate/', generate_pdf, name='generate_pdf'),
    path('modify/', modify_pdf, name='modify_pdf'),
]
