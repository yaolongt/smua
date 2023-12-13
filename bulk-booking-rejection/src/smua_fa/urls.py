from django.urls import path
from smua_fa.views import home 

urlpatterns = [
    path('', home, name='home')
]
