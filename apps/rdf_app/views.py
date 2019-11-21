from datetime import datetime

from django.db.models import Sum, F, FloatField, Q
from django.shortcuts import render
from openpyxl import Workbook
from django.http import HttpResponse
from django.views.generic import TemplateView

from apps.rdf_app.models import Fact, DetailMedi, DetailLabo, DetailDispo, DetailService, DetailMediNoPos
from apps.rdf_app.forms import CreatorForm
from apps.rdf_app.utils.style import give_style

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
            ws.cell(row=fl, column=1).fill = give_style('title').get('relleno')
            ws.cell(row=fl, column=1).alignment = give_style('title').get('alineacion')
            ws.cell(row=fl, column=1).font = give_style('title').get('fuente')
            ws.cell(row=fl, column=1).border = give_style('uniblock-top').get('borde')
            ws.merge_cells(start_row=fl, start_column=1, end_row=fl, end_column=7)
            fl += 1
            ws.cell(row=fl, column=1).value = args[0]
            ws.cell(row=fl, column=1).alignment = give_style('enc-sinwrap').get('alineacion')
            ws.cell(row=fl, column=1).font = give_style('encabezado').get('fuente')
            ws.cell(row=fl, column=1).border = give_style('block-left').get('borde')
            ws.cell(row=fl, column=2).value = args[1]
            ws.cell(row=fl, column=2).alignment = give_style('enc-sinwrap').get('alineacion')
            ws.cell(row=fl, column=2).font = give_style('encabezado').get('fuente')
            ws.cell(row=fl, column=4).value = args[2]
            ws.cell(row=fl, column=4).alignment = give_style('enc-sinwrap').get('alineacion')
            ws.cell(row=fl, column=4).font = give_style('encabezado').get('fuente')
            ws.cell(row=fl, column=5).value = args[3]
            ws.cell(row=fl, column=5).alignment = give_style('enc-sinwrap').get('alineacion')
            ws.cell(row=fl, column=5).font = give_style('encabezado').get('fuente')
            ws.cell(row=fl, column=6).value = args[4]
            ws.cell(row=fl, column=6).alignment = give_style('enc-sinwrap').get('alineacion')
            ws.cell(row=fl, column=6).font = give_style('encabezado').get('fuente')
            ws.cell(row=fl, column=7).value = args[5]
            ws.cell(row=fl, column=7).alignment = give_style('enc-sinwrap').get('alineacion')
            ws.cell(row=fl, column=7).font = give_style('encabezado').get('fuente')
            ws.cell(row=fl, column=7).border = give_style('block-right').get('borde')
            fl += 1
            if type==1:
                for elem in mylist:
                    ws.cell(row=fl, column=1).value = str(getattr(getattr(elem, kwargs['a1']), kwargs['a2']))
                    ws.cell(row=fl, column=1).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=1).alignment = give_style('title').get('alineacion')
                    ws.cell(row=fl, column=1).border = give_style('block-left').get('borde')
                    ws.cell(row=fl, column=2).value = str(getattr(getattr(elem, kwargs['a1']), kwargs['a3']))
                    ws.cell(row=fl, column=2).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=2).alignment = give_style('normal-wrap').get('alineacion')
                    ws.cell(row=fl, column=4).value = str(getattr(elem, kwargs['a4']))
                    ws.cell(row=fl, column=4).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=4).alignment = give_style('title').get('alineacion')
                    ws.cell(row=fl, column=5).value = int(getattr(elem, kwargs['a5']))
                    ws.cell(row=fl, column=5).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=5).alignment = give_style('title').get('alineacion')
                    ws.cell(row=fl, column=6).value = round(float(getattr(elem, kwargs['a6'])), 2)
                    ws.cell(row=fl, column=6).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=6).number_format = '"$"#,##0_);("$"#,##0)'
                    ws.cell(row=fl, column=6).alignment = give_style('total').get('alineacion')
                    ws.cell(row=fl, column=7).value = round(float(getattr(elem, kwargs['a7'])), 2)
                    ws.cell(row=fl, column=7).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=7).number_format = '"$"#,##0_);("$"#,##0)'
                    ws.cell(row=fl, column=7).alignment = give_style('total').get('alineacion')
                    ws.cell(row=fl, column=7).border = give_style('block-right').get('borde')
                    fl += 1
            elif type==2:
                for elem in mylist:
                    ws.cell(row=fl, column=1).value = str(kwargs['a7'])
                    ws.cell(row=fl, column=1).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=1).alignment = give_style('title').get('alineacion')
                    ws.cell(row=fl, column=1).border = give_style('block-left').get('borde')
                    ws.cell(row=fl, column=2).value = str(getattr(getattr(elem, kwargs['a1']), kwargs['a2']))
                    ws.cell(row=fl, column=2).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=2).alignment = give_style('normal-wrap').get('alineacion')
                    ws.cell(row=fl, column=4).value = str(getattr(getattr(elem, kwargs['a1']), kwargs['a3']))
                    ws.cell(row=fl, column=4).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=4).alignment = give_style('title').get('alineacion')
                    ws.cell(row=fl, column=5).value = int(getattr(elem, kwargs['a4']))
                    ws.cell(row=fl, column=5).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=5).alignment = give_style('title').get('alineacion')
                    ws.cell(row=fl, column=6).value = round(float(getattr(elem, kwargs['a5'])), 2)
                    ws.cell(row=fl, column=6).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=6).number_format = '"$"#,##0_);("$"#,##0)'
                    ws.cell(row=fl, column=6).alignment = give_style('total').get('alineacion')
                    ws.cell(row=fl, column=7).value = round(float(getattr(elem, kwargs['a6'])), 2)
                    ws.cell(row=fl, column=7).font = give_style('normal').get('fuente')
                    ws.cell(row=fl, column=7).number_format = '"$"#,##0_);("$"#,##0)'
                    ws.cell(row=fl, column=7).alignment = give_style('total').get('alineacion')
                    ws.cell(row=fl, column=7).border = give_style('block-right').get('borde')
                    fl += 1
            ws.cell(row=fl, column=1).value = args[6]
            ws.cell(row=fl, column=1).fill = give_style('total').get('relleno')
            ws.cell(row=fl, column=1).alignment = give_style('total').get('alineacion')
            ws.cell(row=fl, column=1).font = give_style('total').get('fuente')
            ws.cell(row=fl, column=1).border = give_style('block-btmlft').get('borde')
            ws.merge_cells(start_row=fl, start_column=1, end_row=fl, end_column=6)
            ws.cell(row=fl, column=7).value = round(float(total), 2)
            ws.cell(row=fl, column=7).font = give_style('normal-bold').get('fuente')
            ws.cell(row=fl, column=7).alignment = give_style('total').get('alineacion')
            ws.cell(row=fl, column=7).number_format = '"$"#,##0_);("$"#,##0)'
            ws.cell(row=fl, column=7).border = give_style('block-btmrgt').get('borde')
            fl += 2
        return fl

    def post(self, request):
        my_form = CreatorForm(request.POST)
        if my_form.is_valid():
            workbook = Workbook()
            sheet_temp = workbook.active

            facts = Fact.objects.all()
            for fact in facts:
                # Sheet Creation
                worksheet = workbook.create_sheet(fact.patient.first_name)

                frs_line = 2

                worksheet.cell(row=frs_line, column=1).value = 'DETALLE DE CARGOS DE FACTURA'
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                worksheet.cell(row=frs_line, column=1).border = give_style('uniblock-top').get('borde')
                worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)

                frs_line += 1

                worksheet.cell(row=frs_line, column=1).value = 'FACTURA DE VENTA No '+str(fact.cod_fact)
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                worksheet.cell(row=frs_line, column=1).border = give_style('uniblock-center').get('borde')
                worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)
                frs_line += 1

                worksheet.cell(row=frs_line, column=1).value = 'Periodo facturado del ' + str(
                    fact.cut_ini) + ' al ' + str(fact.cut_end)
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                worksheet.cell(row=frs_line, column=1).border = give_style('uniblock-center').get('borde')
                worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)
                frs_line += 1

                worksheet.cell(row=frs_line, column=1).value = str(fact.regional.company.name) + ' NIT: ' + str(
                    fact.regional.company.number_id) + str(fact.regional.company.cod_verify)
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                worksheet.cell(row=frs_line, column=1).border = give_style('uniblock-center').get('borde')
                worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)
                frs_line += 1

                worksheet.cell(row=frs_line, column=1).value = 'NOMBRE'
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).border = give_style('block-left').get('borde')
                worksheet.cell(row=frs_line, column=2).value = '{}{} {}{}'.format(str(fact.patient.first_name), (
                            ' ' + str(fact.patient.second_name)) if fact.patient.second_name is not None else '',
                                                     str(fact.patient.first_last_name), (' ' + str(
                        fact.patient.second_last_name)) if fact.patient.second_last_name is not None else '')
                worksheet.cell(row=frs_line, column=2).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=6).value = 'HC'
                worksheet.cell(row=frs_line, column=6).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=7).value = str(fact.patient.num_id)
                worksheet.cell(row=frs_line, column=7).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=7).border = give_style('block-right').get('borde')
                frs_line += 1

                worksheet.cell(row=frs_line, column=1).value = 'DIAGNOSTICO'
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).border = give_style('block-left').get('borde')
                worksheet.cell(row=frs_line, column=2).value = str(fact.patient.diagnostic.name)
                worksheet.cell(row=frs_line, column=2).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=6).value = 'EDAD'
                worksheet.cell(row=frs_line, column=6).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=7).value = str(fact.patient.age)+' '+str(fact.patient.age_mess)
                worksheet.cell(row=frs_line, column=7).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=7).border = give_style('block-right').get('borde')
                frs_line += 1

                worksheet.cell(row=frs_line, column=1).value = 'EPS'
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).border = give_style('block-btmlft').get('borde')
                worksheet.cell(row=frs_line, column=2).value = str(fact.patient.eps.name)
                worksheet.cell(row=frs_line, column=2).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=2).alignment = give_style('normal').get('alineacion')
                worksheet.cell(row=frs_line, column=2).border = give_style('block-btmrgt').get('borde')
                worksheet.merge_cells(start_row=frs_line, start_column=2, end_row=frs_line, end_column=7)
                frs_line += 2

                #obtención de datos
                medis_pos = DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact)
                medis_nopos = DetailMediNoPos.objects.filter(fact__cod_fact=fact.cod_fact)
                labos = DetailLabo.objects.filter(fact__cod_fact=fact.cod_fact)
                dispos = DetailDispo.objects.filter(fact__cod_fact=fact.cod_fact)
                servis = DetailService.objects.filter(fact__cod_fact=fact.cod_fact)

                # totales

                total_medis_pos = sum([detalle.subtotal for detalle in
                                       DetailMedi.objects.filter(fact__cod_fact=fact.cod_fact)])

                total_medis_nopos = sum([detalle.subtotal for detalle in
                                         DetailMediNoPos.objects.filter(fact__cod_fact=fact.cod_fact)])
                total_labos = sum(
                    [detalle.subtotal for detalle in DetailLabo.objects.filter(fact__cod_fact=fact.cod_fact)])

                total_dispos = sum(
                    [detalle.subtotal for detalle in DetailDispo.objects.filter(fact__cod_fact=fact.cod_fact)])

                total_services = sum(
                    [detalle.subtotal for detalle in DetailService.objects.filter(fact__cod_fact=fact.cod_fact)])

                '''
                Asi funciona si el valor lo tiene el servicio o producto
                total_services = DetailService.objects.filter(fact__cod_fact=fact.cod_fact).aggregate(
                    sum=Sum(F('cant') * F('service__price'), output_field=FloatField()))
                '''

                # Estancia
                titulos = ['FECHA', 'NOMBRE', 'CODIGO', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL SERVICIO ESTANCIA HOSPITALARIA: ']
                valores = {'a1': 'service', 'a2': 'name', 'a3': 'codigo', 'a4': 'cant', 'a5': 'price',
                           'a6': 'subtotal',
                           'a7': fact.cut_ini}
                frs_line = self.fill_detail(2, 'SERVICIO ESTANCIA HOSPITALARIA', servis,
                                            worksheet,
                                            frs_line,
                                            total_services, *titulos, **valores)
                #frs_line+=1
                # Medicamentos POS
                titulos = ['CODIGO CUM', 'MEDICAMENTOS', 'DOSIS', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL SERVICIO MEDICAMENTOS POS: ']
                valores = {'a1': 'medicine', 'a2': 'cod_cum', 'a3': 'name', 'a4': 'dosis', 'a5': 'cant', 'a6': 'price',
                           'a7': 'subtotal'}
                frs_line = self.fill_detail(1, 'SERVICIO MEDICAMENTOS POS', medis_pos, worksheet, frs_line,
                                 total_medis_pos, *titulos, **valores)

                #frs_line += 1
                # Medicamentos NO POS
                titulos = ['CODIGO CUM', 'MEDICAMENTOS', 'DOSIS', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL SERVICIO MEDICAMENTOS NO POS: ']
                valores = {'a1': 'medicine', 'a2': 'cod_cum', 'a3': 'name', 'a4': 'dosis', 'a5': 'cant', 'a6': 'price',
                           'a7': 'subtotal'}
                frs_line = self.fill_detail(1, 'SERVICIO MEDICAMENTOS NO POS', medis_nopos, worksheet,
                                            frs_line,
                                            total_medis_nopos, *titulos, **valores)

                #frs_line += 1
                # Dispositivos
                titulos = ['FECHA', 'NOMBRE', 'CODIGO', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL SERVICIO DISPOSITIVOS MEDICO-QUIRURGICOS: ']
                valores = {'a1': 'dispositive', 'a2': 'name', 'a3': 'codigo', 'a4': 'cant', 'a5': 'price',
                           'a6': 'subtotal',
                           'a7': fact.cut_ini}
                frs_line = self.fill_detail(2, 'SERVICIO DISPOSITIVOS MEDICO-QUIRURGICOS', dispos,
                                            worksheet,
                                            frs_line,
                                            total_dispos, *titulos, **valores)

                #frs_line += 1
                # Laboratorios
                titulos = ['FECHA', 'NOMBRE', 'CODIGO', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL',
                           'TOTAL SERVICIO LABORATORIOS: ']
                valores = {'a1': 'laboratory', 'a2': 'name', 'a3': 'codigo', 'a4': 'cant', 'a5': 'price', 'a6': 'subtotal',
                           'a7': fact.cut_ini}
                frs_line = self.fill_detail(2, 'SERVICIO LABORATORIO CLINICO', labos,
                                            worksheet,
                                            frs_line,
                                            total_labos, *titulos, **valores)

                #frs_line += 1

                worksheet.cell(row=frs_line, column=1).value = '{}{}{}'.format(
                    'SERVICIO REMEO - PAQUETE/ESTANCIA - ORDEN PQTE: ' + str(fact.aut_number),
                    (' PIN-ELECTRONICO:  ' + str(fact.pin_elect)) if fact.pin_elect is not None else '',
                    (' VALIDACION: ' + str(fact.validation)) if fact.validation is not None else '')
                worksheet.cell(row=frs_line, column=1).fill = give_style('title').get('relleno')
                worksheet.cell(row=frs_line, column=1).alignment = give_style('title-wrap').get('alineacion')
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).border = give_style('uniblock-top').get('borde')
                worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=7)
                frs_line += 1
                total_factura = 0
                worksheet.cell(row=frs_line, column=1).value = 'FECHA'
                worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                worksheet.cell(row=frs_line, column=1).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=1).border = give_style('block-left').get('borde')
                worksheet.cell(row=frs_line, column=2).value = 'SERVICIO'
                worksheet.cell(row=frs_line, column=2).alignment = give_style('title').get('alineacion')
                worksheet.cell(row=frs_line, column=2).font = give_style('title').get('fuente')
                worksheet.merge_cells(start_row=frs_line, start_column=2, end_row=frs_line, end_column=6)
                worksheet.cell(row=frs_line, column=7).value = 'TOTAL'
                worksheet.cell(row=frs_line, column=7).alignment = give_style('title').get('alineacion')
                worksheet.cell(row=frs_line, column=7).font = give_style('title').get('fuente')
                worksheet.cell(row=frs_line, column=7).border = give_style('block-right').get('borde')
                frs_line += 1
                if servis:
                    worksheet.cell(row=frs_line, column=1).value = str(fact.cut_ini)
                    worksheet.cell(row=frs_line, column=1).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=1).border = give_style('block-left').get('borde')
                    worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                    worksheet.cell(row=frs_line, column=2).value = 'SERVICIO ESTANCIA HOSPITALARIA'
                    worksheet.cell(row=frs_line, column=2).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=2).alignment = give_style('normal').get('alineacion')
                    worksheet.merge_cells(start_row=frs_line, start_column=2, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = total_services
                    worksheet.cell(row=frs_line, column=7).font = give_style('total').get('fuente')
                    worksheet.cell(row=frs_line, column=7).number_format = '"$"#,##0_);("$"#,##0)'
                    worksheet.cell(row=frs_line, column=7).border = give_style('block-right').get('borde')
                    worksheet.cell(row=frs_line, column=7).alignment = give_style('total').get('alineacion')
                    total_factura+=total_services
                    frs_line+=1
                if medis_pos:
                    worksheet.cell(row=frs_line, column=1).value = str(fact.cut_ini)
                    worksheet.cell(row=frs_line, column=1).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=1).border = give_style('block-left').get('borde')
                    worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                    worksheet.cell(row=frs_line, column=2).value = 'SERVICIO MEDICAMENTOS POS'
                    worksheet.cell(row=frs_line, column=2).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=2).alignment = give_style('normal').get('alineacion')
                    worksheet.merge_cells(start_row=frs_line, start_column=2, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = total_medis_pos
                    worksheet.cell(row=frs_line, column=7).font = give_style('total').get('fuente')
                    worksheet.cell(row=frs_line, column=7).number_format = '"$"#,##0_);("$"#,##0)'
                    worksheet.cell(row=frs_line, column=7).border = give_style('block-right').get('borde')
                    worksheet.cell(row=frs_line, column=7).alignment = give_style('total').get('alineacion')
                    total_factura += total_medis_pos
                    frs_line += 1
                if medis_nopos:
                    worksheet.cell(row=frs_line, column=1).value = str(fact.cut_ini)
                    worksheet.cell(row=frs_line, column=1).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=1).border = give_style('block-left').get('borde')
                    worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                    worksheet.cell(row=frs_line, column=2).value = 'SERVICIO MEDICAMENTOS NO POS'
                    worksheet.cell(row=frs_line, column=2).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=2).alignment = give_style('normal').get('alineacion')
                    worksheet.merge_cells(start_row=frs_line, start_column=2, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = total_medis_nopos
                    worksheet.cell(row=frs_line, column=7).font = give_style('total').get('fuente')
                    worksheet.cell(row=frs_line, column=7).number_format = '"$"#,##0_);("$"#,##0)'
                    worksheet.cell(row=frs_line, column=7).border = give_style('block-right').get('borde')
                    worksheet.cell(row=frs_line, column=7).alignment = give_style('total').get('alineacion')
                    total_factura += total_medis_nopos
                    frs_line += 1
                if dispos:
                    worksheet.cell(row=frs_line, column=1).value = str(fact.cut_ini)
                    worksheet.cell(row=frs_line, column=1).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=1).border = give_style('block-left').get('borde')
                    worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                    worksheet.cell(row=frs_line, column=2).value = 'SERVICIO DISPOSITIVOS MEDICO-QUIRURGICOS'
                    worksheet.cell(row=frs_line, column=2).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=2).alignment = give_style('normal').get('alineacion')
                    worksheet.merge_cells(start_row=frs_line, start_column=2, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = total_dispos
                    worksheet.cell(row=frs_line, column=7).font = give_style('total').get('fuente')
                    worksheet.cell(row=frs_line, column=7).number_format = '"$"#,##0_);("$"#,##0)'
                    worksheet.cell(row=frs_line, column=7).border = give_style('block-right').get('borde')
                    worksheet.cell(row=frs_line, column=7).alignment = give_style('total').get('alineacion')
                    total_factura += total_dispos
                    frs_line += 1
                if labos:
                    worksheet.cell(row=frs_line, column=1).value = str(fact.cut_ini)
                    worksheet.cell(row=frs_line, column=1).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=1).border = give_style('block-left').get('borde')
                    worksheet.cell(row=frs_line, column=1).alignment = give_style('title').get('alineacion')
                    worksheet.cell(row=frs_line, column=2).value = 'SERVICIO LABORATORIO CLINICO'
                    worksheet.cell(row=frs_line, column=2).font = give_style('normal').get('fuente')
                    worksheet.cell(row=frs_line, column=2).alignment = give_style('normal').get('alineacion')
                    worksheet.merge_cells(start_row=frs_line, start_column=2, end_row=frs_line, end_column=6)
                    worksheet.cell(row=frs_line, column=7).value = total_labos
                    worksheet.cell(row=frs_line, column=7).font = give_style('total').get('fuente')
                    worksheet.cell(row=frs_line, column=7).number_format = '"$"#,##0_);("$"#,##0)'
                    worksheet.cell(row=frs_line, column=7).border = give_style('block-right').get('borde')
                    worksheet.cell(row=frs_line, column=7).alignment = give_style('total').get('alineacion')
                    total_factura += total_labos
                    frs_line += 1
                worksheet.cell(row=frs_line, column=1).value = 'TOTAL VALOR FACTURADO: '
                worksheet.cell(row=frs_line, column=1).fill = give_style('total').get('relleno')
                worksheet.cell(row=frs_line, column=1).alignment = give_style('total').get('alineacion')
                worksheet.cell(row=frs_line, column=1).font = give_style('total').get('fuente')
                worksheet.cell(row=frs_line, column=1).border = give_style('block-btmlft').get('borde')
                worksheet.merge_cells(start_row=frs_line, start_column=1, end_row=frs_line, end_column=6)
                worksheet.cell(row=frs_line, column=7).value = round(total_factura,2)
                worksheet.cell(row=frs_line, column=7).number_format = '"$"#,##0_);("$"#,##0)'
                worksheet.cell(row=frs_line, column=7).border = give_style('block-btmrgt').get('borde')
                worksheet.cell(row=frs_line, column=7).font = give_style('enc-sinwrap').get('fuente')
                worksheet.cell(row=frs_line, column=7).alignment = give_style('total').get('alineacion')


            workbook.remove(sheet_temp)


            response = HttpResponse(
                content_type = 'application/ms-excel'
            )
            response['Content-Disposition'] = 'attachment; filename = detallado_{date}.xlsx'.format(
                date=datetime.now().strftime("%d%m%Y-%H%M%S"),
            )
            workbook.save(response)
            return response
        return render(request, "rdf/creator.html", {'form': my_form})