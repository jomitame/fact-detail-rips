from django.urls import path
from apps.rdf_app.views import CreatorXLSXView, GeneratorRIPSView

app_name = 'rdf'
urlpatterns = [
    path('detallados',CreatorXLSXView.as_view(),name='detallados'),
    path('rips',GeneratorRIPSView.as_view(),name='rips')
    ]