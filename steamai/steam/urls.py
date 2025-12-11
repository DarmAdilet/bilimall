# urls.py - Толық маршруттар

from django.urls import path
from . import views

urlpatterns = [
    # Негізгі беттер
    path('', views.index, name='index'),
    path('chat/', views.chat, name='chat'),
    path('chat1/', views.chat_with_gemini, name='chat1'),
    path('slide/', views.slide1, name='slide'),
    path('exercise/', views.exercise, name='exercise'),
    path('chatbot-assistant/', views.chatbot_assistant, name='chatbot_assistant'),
    # Пәндер
    path('fizika/', views.fizika, name='fizika'),
    path('biology/', views.biology, name='biology'),
    path('geography/', views.geography, name='geography'),
    path('matem/', views.matem, name='matem'),
    path('ximia/', views.ximia, name='ximia'),
    path('informatika/', views.informatika, name='informatika'),

    # API endpoints
    path('api/chat/', views.chat_with_gemini, name='chat_api'),
    path('api/generate-slide/', views.generate_slide, name='generate_slide'),
    path('generate-slide/', views.generate_slide, name='generate_slide_alt'),  # Қосымша маршрут
    path('api/generate-exercise/', views.generate_exercise, name='generate_exercise'),

    # Ойын маршруттары
    path('game/', views.game_view, name='game'),
    path('api/generate-game-questions/', views.generate_game_questions_api, name='generate_game_questions'),

    # Word жүктеу
    path('api/download-exercise-word/', views.download_exercise_word, name='download_exercise_word'),
]