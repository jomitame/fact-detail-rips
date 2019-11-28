from django.contrib import admin

from .models import Company, Regional, EPS, Diagnostic, Patient, Medicine, Fact, DetailMedi, Dispositive, DetailDispo, \
    Laboratory, DetailLabo, Service, DetailService, Presentation, Concentration, Departament, Municipe, PriceMedicine, \
    PriceService, PriceDispositive, PriceLabo, DetailMediNoPos

class servi_inline(admin.TabularInline):
    model = (DetailService)
    extra = 0
    readonly_fields = ('price','subtotal',)
    autocomplete_fields = ('service',)

class medi_inline(admin.TabularInline):
    model =  (DetailMedi)
    extra = 0
    readonly_fields = ('price','subtotal',)
    autocomplete_fields = ('medicine',)

class medinopos_inline(admin.TabularInline):
    model = (DetailMediNoPos)
    extra = 0
    readonly_fields = ('price','subtotal',)
    autocomplete_fields = ('medicine',)

class labo_inline(admin.TabularInline):
    model = (DetailLabo)
    extra = 0
    readonly_fields = ('price','subtotal',)
    autocomplete_fields = ('laboratory',)

class dispo_inline(admin.TabularInline):
    model = (DetailDispo)
    extra = 0
    readonly_fields = ('price','subtotal',)
    autocomplete_fields = ('dispositive',)


@admin.register(Fact)
class factAdmin(admin.ModelAdmin):
    inlines = (servi_inline, medi_inline, medinopos_inline, dispo_inline, labo_inline, )
    autocomplete_fields = ('patient',)
    ordering = ('cod_fact',)


@admin.register(Service)
class ServiceAdmin(admin.ModelAdmin):
    ordering = ('name',)
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(Medicine)
class MedicineAdmin(admin.ModelAdmin):
    ordering = ('name',)
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(Dispositive)
class DispositiveAdmin(admin.ModelAdmin):
    ordering = ('name',)
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(Laboratory)
class LaboratoryAdmin(admin.ModelAdmin):
    ordering = ('name',)
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(Patient)
class PatientAdmin(admin.ModelAdmin):
    ordering = ('first_name',)
    list_display = ('first_name',)
    search_fields = ('first_name',)
    readonly_fields = ('age','age_mess')

@admin.register(PriceMedicine)
class PriceMedicine(admin.ModelAdmin):
    ordering = ('medto__name_rips',)
    search_fields = ('medto__name_rips',)

@admin.register(PriceLabo)
class PriceMedicine(admin.ModelAdmin):
    ordering = ('labo__name_rips',)
    search_fields = ('labo__name_rips',)

@admin.register(PriceDispositive)
class PriceMedicine(admin.ModelAdmin):
    ordering = ('dispo__name_rips',)
    search_fields = ('dispo__name_rips',)

@admin.register(PriceService)
class PriceMedicine(admin.ModelAdmin):
    ordering = ('servi__name_rips',)
    search_fields = ('servi__name_rips',)
'''
class PapientAdmin(admin.ModelAdmin):
    list_display = ['first_name', 'first_last_name','age', 'age_mess']

'''



admin.site.register(Company)
admin.site.register(Regional)
admin.site.register(EPS)
admin.site.register(Diagnostic)
admin.site.register(DetailMedi)
admin.site.register(DetailMediNoPos)
admin.site.register(DetailLabo)
admin.site.register(DetailDispo)
admin.site.register(DetailService)
admin.site.register(Presentation)
admin.site.register(Concentration)
admin.site.register(Departament)
admin.site.register(Municipe)
#admin.site.register(PriceMedicine)
#admin.site.register(PriceLabo)
#admin.site.register(PriceDispositive)
#admin.site.register(PriceService)

