from django.urls import path
from apps.rdf_app.views import CreatorXLSXView, GeneratorRIPSView

app_name = 'rdf'
urlpatterns = [
    path('creator',CreatorXLSXView.as_view(),name='creator'),
    path('generator',GeneratorRIPSView.as_view(),name='generator')
    ]