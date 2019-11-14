from datetime import datetime
from django.shortcuts import render

from django.shortcuts import render
from openpyxl import Workbook


from django.http import HttpResponse
from django.views.generic import TemplateView

# DetailMedi.objects.filter(fact__cod_fact='3456r3').aggregate(sum=Sum(F('cant')*F('medicine__price'), output_field=FloatField()))

from .models import Fact
from .forms import CreatorForm

class CreatorXLSXView(TemplateView):
    model = Fact
    template_name = 'rdf/creator.html'
    form_class = CreatorForm

    def get(self, request):
        return render(request, self.template_name, {'form':self.form_class})

    def form_invalid(self, form):
        return super(CreatorXLSXView, self).form_invalid(form)

    def post(self, request):
        pass