from django.urls import path
from web_fa.views import home

urlpatterns = [
    path('', home, name='home'),
]
