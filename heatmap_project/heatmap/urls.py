from django.urls import path
from . import views

urlpatterns = [
    path('',           views.home,           name='home'),
    path('procesar/',  views.procesar,        name='procesar'),
    path('descargar/', views.descargar_excel, name='descargar'),
]
