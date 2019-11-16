from django.contrib import admin


from .models import Company, Regional, EPS, Diagnostic, Patient, Medicine, Fact, DetailMedi, Treatement, DetailTreat, \
    Laboratory, DetailLabo

class medi_inline(admin.TabularInline):
    model =  (DetailMedi)
    extra = 0
    readonly_fields = ('subtotal',)
    autocomplete_fields = ('medicine',)

class labo_inline(admin.TabularInline):
    model = (DetailLabo)
    extra = 0
    readonly_fields = ('subtotal',)
    autocomplete_fields = ('laboratory',)

class treat_inline(admin.TabularInline):
    model = (DetailTreat)
    extra = 0
    readonly_fields = ('subtotal',)
    autocomplete_fields = ('treatement',)


@admin.register(Fact)
class factAdmin(admin.ModelAdmin):
    inlines = (medi_inline, labo_inline, treat_inline)
    autocomplete_fields = ('patient',)



@admin.register(Medicine)
class MedicineAdmin(admin.ModelAdmin):
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(Treatement)
class TreatementAdmin(admin.ModelAdmin):
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(Laboratory)
class TreatementAdmin(admin.ModelAdmin):
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(Patient)
class PatientAdmin(admin.ModelAdmin):
    list_display = ('first_name',)
    search_fields = ('first_name',)
    readonly_fields = ('age','age_mess')

'''
class PapientAdmin(admin.ModelAdmin):
    list_display = ['first_name', 'first_last_name','age', 'age_mess']

'''



admin.site.register(Company)
admin.site.register(Regional)
admin.site.register(EPS)
admin.site.register(Diagnostic)
admin.site.register(DetailMedi)
admin.site.register(DetailLabo)
admin.site.register(DetailTreat)
#admin.site.register(Fact, factAdmin)
#admin.site.register(Patient, PapientAdmin)
#admin.site.register(Patient)
#admin.site.register(Medicine)
#admin.site.register(Fact)
#admin.site.register(Treatement)
