from django.urls import path
from .views import CreatorXLSXView

app_name = 'rdf'
urlpatterns = [
    path('creator',CreatorXLSXView.as_view(),name='creator')
    ]