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
                medis_pos = DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact, medicine__is_pos=True)
                medis_nopos = DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact, medicine__is_pos=False)
                labos = DetailLabo.objects.filter(fact__cod_fact=fact.cod_fact)
                treats = DetailTreat.objects.filter(fact__cod_fact=fact.cod_fact)
                total_medis = DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact).aggregate(
                    sum=Sum(F('cant') * F('medicine__price'), output_field=FloatField()))
                total_labos = DetailLabo.objects.filter(fact__cod_fact=fact.cod_fact).aggregate(
                    sum=Sum(F('cant') * F('laboratory__price'), output_field=FloatField()))
                total_treats = DetailTreat.objects.filter(fact__cod_fact=fact.cod_fact).aggregate(
                    sum=Sum(F('cant') * F('treatement__price'), output_field=FloatField()))

                frs_line = 9
                # Medicamentos POS
                if medis_pos:
                    worksheet.cell(row=frs_line, column=1).value = 'SERVICIO FARMACEUTICO REMEO - MEDICAMENTOS POS'
                    worksheet.cell(row=frs_line, column=1).fill
                    worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)
                    frs_line += 1
                    worksheet.cell(row=frs_line, column=1).value = 'CODIGO CUM'
                    worksheet.cell(row=frs_line, column=2).value = 'MEDICAMENTOS'
                    worksheet.cell(row=frs_line, column=4).value = 'DOSIS'
                    worksheet.cell(row=frs_line, column=5).value = 'CANTIDAD'
                    worksheet.cell(row=frs_line, column=6).value = 'VALOR UNITARIO'
                    worksheet.cell(row=frs_line, column=7).value = 'VALOR TOTAL'
                    frs_line += 1
                    for medi in medis_pos:
                        worksheet.cell(row=frs_line, column=1).value = str(medi.medicine.cod_cum)
                        worksheet.cell(row=frs_line, column=2).value = str(medi.medicine.name)
                        worksheet.cell(row=frs_line, column=4).value = int(medi.dosis)
                        worksheet.cell(row=frs_line, column=5).value = int(medi.cant)
                        worksheet.cell(row=frs_line, column=6).value = round(float(medi.medicine.price),2)
                        worksheet.cell(row=frs_line, column=7).value = round(float(medi.subtotal),2)
                        frs_line += 1
                    worksheet.cell(row=frs_line, column=1).value = 'TOTAL MEDICAMENTOS POS'
                    worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = round(float(total_medis['sum']),2)
                    frs_line += 1

                # MEdicamentos NO POS
                if medis_nopos:
                    worksheet.cell(row=frs_line, column=1).value = 'SERVICIO FARMACEUTICO REMEO - MEDICAMENTOS NO POS'
                    worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)
                    frs_line += 1
                    worksheet.cell(row=frs_line, column=1).value = 'CODIGO CUM'
                    worksheet.cell(row=frs_line, column=2).value = 'MEDICAMENTOS'
                    worksheet.cell(row=frs_line, column=4).value = 'DOSIS'
                    worksheet.cell(row=frs_line, column=5).value = 'CANTIDAD'
                    worksheet.cell(row=frs_line, column=6).value = 'VALOR UNITARIO'
                    worksheet.cell(row=frs_line, column=7).value = 'VALOR TOTAL'
                    frs_line += 1
                    for medi in medis_pos:
                        worksheet.cell(row=frs_line, column=1).value = str(medi.medicine.cod_cum)
                        worksheet.cell(row=frs_line, column=2).value = str(medi.medicine.name)
                        worksheet.cell(row=frs_line, column=4).value = int(medi.dosis)
                        worksheet.cell(row=frs_line, column=5).value = int(medi.cant)
                        worksheet.cell(row=frs_line, column=6).value = round(float(medi.medicine.price),2)
                        worksheet.cell(row=frs_line, column=7).value = round(float(medi.subtotal),2)
                        frs_line += 1
                    worksheet.cell(row=frs_line, column=1).value = 'TOTAL MEDICAMENTOS NO POS'
                    worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = round(float(total_medis['sum']),2)
                    frs_line += 1

                # Laboratorios
                if labos:
                    worksheet.cell(row=frs_line, column=1).value = 'SERVICIO LABORATORIO CLINICO'
                    worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)
                    frs_line += 1
                    worksheet.cell(row=frs_line, column=1).value = 'FECHA'
                    worksheet.cell(row=frs_line, column=2).value = 'NOMBRE'
                    worksheet.cell(row=frs_line, column=4).value = 'CODIGO'
                    worksheet.cell(row=frs_line, column=5).value = 'CANTIDAD'
                    worksheet.cell(row=frs_line, column=6).value = 'VALOR UNITARIO'
                    worksheet.cell(row=frs_line, column=7).value = 'VALOR TOTAL'
                    frs_line += 1
                    for labo in labos:
                        worksheet.cell(row=frs_line, column=1).value = str(fact.cut_ini)
                        worksheet.cell(row=frs_line, column=2).value = str(labo.laboratory.name)
                        worksheet.cell(row=frs_line, column=4).value = str(labo.laboratory.codigo)
                        worksheet.cell(row=frs_line, column=5).value = int(labo.cant)
                        worksheet.cell(row=frs_line, column=6).value = round(float(labo.laboratory.price), 2)
                        worksheet.cell(row=frs_line, column=7).value = round(float(labo.subtotal), 2)
                        frs_line += 1
                    worksheet.cell(row=frs_line, column=1).value = 'TOTAL SERVICIO DE LABORATORIO CLINICO'
                    worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = round(float(total_labos['sum']), 2)
                    frs_line += 1

                if treats:
                    worksheet.cell(row=frs_line, column=1).value = 'TRATAMIENTOS'
                    worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)
                    frs_line += 1
                    worksheet.cell(row=frs_line, column=1).value = 'FECHA'
                    worksheet.cell(row=frs_line, column=2).value = 'NOMBRE'
                    worksheet.cell(row=frs_line, column=4).value = 'CODIGO'
                    worksheet.cell(row=frs_line, column=5).value = 'CANTIDAD'
                    worksheet.cell(row=frs_line, column=6).value = 'VALOR UNITARIO'
                    worksheet.cell(row=frs_line, column=7).value = 'VALOR TOTAL'
                    frs_line += 1
                    for treat in treats:
                        worksheet.cell(row=frs_line, column=1).value = str(fact.cut_ini)
                        worksheet.cell(row=frs_line, column=2).value = str(treat.treatement.name)
                        worksheet.cell(row=frs_line, column=4).value = str(treat.treatement.cod_treat)
                        worksheet.cell(row=frs_line, column=5).value = int(treat.cant)
                        worksheet.cell(row=frs_line, column=6).value = round(float(treat.treatement.price), 2)
                        worksheet.cell(row=frs_line, column=7).value = round(float(treat.subtotal), 2)
                        frs_line += 1
                    worksheet.cell(row=frs_line, column=1).value = 'TOTAL SERVICIO DE TRATAMIENTOS'
                    worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = round(float(total_treats['sum']), 2)
                    frs_line += 1

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