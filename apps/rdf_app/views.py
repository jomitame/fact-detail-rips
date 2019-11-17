from datetime import datetime

from django.db.models import Sum, F, FloatField
from django.shortcuts import render
from openpyxl import Workbook
from django.http import HttpResponse
from django.views.generic import TemplateView

from apps.rdf_app.models import Fact, DetailMedi, DetailLabo, DetailTreat
from .forms import CreatorForm

class CreatorXLSXView(TemplateView):
    model = Fact
    template_name = 'rdf/creator.html'
    form_class = CreatorForm


    def get(self, request):
        return render(request, self.template_name, {'form':self.form_class})

    def form_invalid(self, form):
        return super(CreatorXLSXView, self).form_invalid(form)


    def fill_detail(self, type, title, mylist, ws, fl, total, *args, **kwargs ):
        if mylist:
            ws.cell(row=fl, column=1).value = title
            ws.merge_cells(start_row=fl, start_column=1, end_row=fl, end_column=7)
            fl += 1
            ws.cell(row=fl, column=1).value = args[0]
            ws.cell(row=fl, column=2).value = args[1]
            ws.cell(row=fl, column=4).value = args[2]
            ws.cell(row=fl, column=5).value = args[3]
            ws.cell(row=fl, column=6).value = args[4]
            ws.cell(row=fl, column=7).value = args[5]
            fl += 1
            if type==1:
                for elem in mylist:
                    ws.cell(row=fl, column=1).value = str(getattr(getattr(elem, kwargs['a1']), kwargs['a2']))
                    ws.cell(row=fl, column=2).value = str(getattr(getattr(elem, kwargs['a1']), kwargs['a3']))
                    ws.cell(row=fl, column=4).value = str(getattr(elem, kwargs['a4']))
                    ws.cell(row=fl, column=5).value = int(getattr(elem, kwargs['a5']))
                    ws.cell(row=fl, column=6).value = round(float(getattr(getattr(elem, kwargs['a1']), kwargs['a6'])), 2)
                    ws.cell(row=fl, column=7).value = round(float(getattr(elem, kwargs['a7'])), 2)
                    fl += 1
            elif type==2:
                for elem in mylist:
                    ws.cell(row=fl, column=1).value = str(kwargs['a7'])
                    ws.cell(row=fl, column=2).value = str(getattr(getattr(elem, kwargs['a1']), kwargs['a2']))
                    ws.cell(row=fl, column=4).value = str(getattr(getattr(elem, kwargs['a1']), kwargs['a3']))
                    ws.cell(row=fl, column=5).value = int(getattr(elem, kwargs['a4']))
                    ws.cell(row=fl, column=6).value = round(float(getattr(getattr(elem, kwargs['a1']), kwargs['a5'])), 2)
                    ws.cell(row=fl, column=7).value = round(float(getattr(elem, kwargs['a6'])), 2)
                    fl += 1
            ws.cell(row=fl, column=1).value = args[6]
            ws.merge_cells(start_row=fl, start_column=1, end_row=fl, end_column=6)
            ws.cell(row=fl, column=7).value = round(float(total), 2)
            fl += 1
        return fl

    def post(self, request):
        my_form = CreatorForm(request.POST)
        if my_form.is_valid():
            workbook = Workbook()
            sheet_temp = workbook.active

            facts = Fact.objects.all()
            for fact in facts:
                worksheet = workbook.create_sheet(fact.patient.first_name)
                worksheet['A2'] = 'DETALLE DE CARGOS DE FACTURA'
                worksheet.merge_cells('A2:G2')
                worksheet['A3'] = 'FACTURA DE VENTA No '+str(fact.cod_fact)
                worksheet.merge_cells('A3:G3')
                worksheet['A4'] = 'Periodo facturado del '+str(fact.cut_ini)+' al '+str(fact.cut_end)
                worksheet.merge_cells('A4:G4')
                worksheet['A5'] = str(fact.regional.company.name)
                worksheet['D5'] = str(fact.regional.company.number_id)+str(fact.regional.company.cod_verify)
                worksheet['A6'] = 'NOMBRE'
                worksheet['B6'] = str(fact.patient.first_name) + (' ' + str(
                    fact.patient.second_last_name)) if fact.patient.second_name is not None else '' + ' ' + str(
                    fact.patient.first_last_name) + (' ' + str(
                    fact.patient.second_last_name)) if fact.patient.second_last_name is not None else ''
                worksheet['E6'] = 'HC'
                worksheet['F6'] = str(fact.patient.num_id)
                worksheet['A7'] = 'DIAGNOSTICO'
                worksheet['B7'] = str(fact.patient.diagnostic.name)
                worksheet['E7'] = 'EDAD'
                worksheet['F7'] = str(fact.patient.age)+' '+str(fact.patient.age_mess)
                worksheet['A8'] = 'EPS'
                worksheet['B8'] = str(fact.patient.eps.name)


                #obtención de datos
                frs_line = 10
                medis_pos = DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact, medicine__is_pos=True)
                medis_nopos = DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact, medicine__is_pos=False)
                labos = DetailLabo.objects.filter(fact__cod_fact=fact.cod_fact)
                treats = DetailTreat.objects.filter(fact__cod_fact=fact.cod_fact)
                #----------------------------cambiar total no pos
                total_medis_pos = DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact,
                                                            medicine__is_pos=True).aggregate(
                    sum=Sum(F('cant') * F('medicine__price'), output_field=FloatField()))
                total_medis_nopos = DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact,
                                                              medicine__is_pos=False).aggregate(
                    sum=Sum(F('cant') * F('medicine__price'), output_field=FloatField()))
                total_labos = DetailLabo.objects.filter(fact__cod_fact=fact.cod_fact).aggregate(
                    sum=Sum(F('cant') * F('laboratory__price'), output_field=FloatField()))
                total_treats = DetailTreat.objects.filter(fact__cod_fact=fact.cod_fact).aggregate(
                    sum=Sum(F('cant') * F('treatement__price'), output_field=FloatField()))


                # Medicamentos POS
                titulos = ['CODIGO CUM', 'MEDICAMENTOS', 'DOSIS', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL MEDICAMENTOS POS']
                valores = {'a1': 'medicine', 'a2': 'cod_cum', 'a3': 'name', 'a4': 'dosis', 'a5': 'cant', 'a6': 'price',
                           'a7': 'subtotal'}
                frs_line = self.fill_detail(1, 'SERVICIO FARMACEUTICO REMEO - MEDICAMENTOS POS', medis_pos, worksheet, frs_line,
                                 total_medis_pos['sum'], *titulos, **valores)

                # Medicamentos NO POS
                titulos = ['CODIGO CUM', 'MEDICAMENTOS', 'DOSIS', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL MEDICAMENTOS NO POS']
                valores = {'a1': 'medicine', 'a2': 'cod_cum', 'a3': 'name', 'a4': 'dosis', 'a5': 'cant', 'a6': 'price',
                           'a7': 'subtotal'}
                frs_line = self.fill_detail(1, 'SERVICIO FARMACEUTICO REMEO - MEDICAMENTOS NO POS', medis_nopos, worksheet,
                                            frs_line,
                                            total_medis_nopos['sum'], *titulos, **valores)

                # Laboratorios
                titulos = ['FECHA', 'NOMBRE', 'CODIGO', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL MEDICAMENTOS LABORATORIOS']
                valores = {'a1': 'laboratory', 'a2': 'name', 'a3': 'codigo', 'a4': 'cant', 'a5': 'price', 'a6': 'subtotal',
                           'a7': fact.cut_ini}
                frs_line = self.fill_detail(2, 'SERVICIO LABORATORIO CLINICO', labos,
                                            worksheet,
                                            frs_line,
                                            total_labos['sum'], *titulos, **valores)

                # Tratamientos
                titulos = ['FECHA', 'NOMBRE', 'CODIGO', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL MEDICAMENTOS LABORATORIOS']
                valores = {'a1': 'treatement', 'a2': 'name', 'a3': 'cod_treat', 'a4': 'cant', 'a5': 'price',
                           'a6': 'subtotal',
                           'a7': fact.cut_ini}
                frs_line = self.fill_detail(2, 'TRATAMIENTOS', treats,
                                            worksheet,
                                            frs_line,
                                            total_treats['sum'], *titulos, **valores)




            workbook.remove(sheet_temp)


            response = HttpResponse(
                content_type = 'application/ms-excel'
            )
            response['Content-Disposition'] = 'attachment; filename = report_{date}.xlsx'.format(
                date=datetime.now().strftime("%d%m%Y-%H%M%S"),
            )
            workbook.save(response)
            return response
        return render(request, "rdf/creator.html", {'form': my_form})