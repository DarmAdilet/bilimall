from django.urls import path
from . import views

urlpatterns = [
    path('',views.index,name='index'),
    path('chat/', views.chat, name='chat'),  # chat.html-ге дұрыс маршрут
    path('chat1/', views.chat_with_gemini),
    path('slide/', views.slide1, name='slide'),
    path('generate-slide/', views.generate_slide, name='generate_slide'),
    path('biology/', views.biology, name='biology'),
    path('fizika/', views.fizika, name='fizika'),
    path('geography/', views.geography, name='geography'),
    path('informatika/', views.informatika, name='informatika'),
    path('matem/', views.matem, name='matem'),
    path('ximia/', views.ximia, name='ximia'),
]