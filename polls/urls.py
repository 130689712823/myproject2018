from django.urls import path

from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('execl2word', views.execl2word, name='execl2word'),
]