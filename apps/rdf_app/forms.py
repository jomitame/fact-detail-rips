import datetime

from django import forms


class CreatorForm(forms.Form):

    #fecha_ini = forms.DateField(label='Fecha Inicial', widget=DatePickerInput(format='%Y-%m-%d'))
    #fecha_fin = forms.DateField(label='Fecha Final', widget=DatePickerInput(format='%Y-%m-%d'))
    fecha_ini =  forms.DateField(label='Fecha Inicial')
    fecha_fin = forms.DateField(label='Fecha Final')

    def clean(self):
        cleaned_data = super(CreatorForm, self).clean()
        fecha_ini = cleaned_data.get('fecha_ini')
        fecha_fin = cleaned_data.get('fecha_fin')
        if fecha_fin <= datetime.date.today():
            if fecha_ini > fecha_fin:
                raise forms.ValidationError("Fecha inicial no puede ser mayor a la fecha final")
        else:
            raise forms.ValidationError("Fecha final no puede ser mayor a hoy")

        return cleaned_data