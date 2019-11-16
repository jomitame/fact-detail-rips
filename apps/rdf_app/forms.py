import datetime

from django import forms


class CreatorForm(forms.Form):

    fecha_ini = forms.DateField(
        label='Fecha Inicial',
        input_formats=['%d/%m/%Y'],
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input',
            'data-target': '#datetimepicker1'
        })
    )
    fecha_fin = forms.DateField(#DateTimeField
        label='Fecha Final',
        input_formats=['%d/%m/%Y'],# %H:%M
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input',
            'data-target': '#datetimepicker2'
        })
    )

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