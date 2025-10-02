from django.urls import path
from . import views

urlpatterns = [
    path('',views.generar_secreto_bancario2,name='generar_secreto_bancario2')
]
