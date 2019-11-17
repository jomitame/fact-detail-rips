import datetime

from django.db import models

from .choices import CEDULA, NIT, TYPE_ID, COD_DPTO, COD_MUNI, URBA, URBA_RUL, GEN, YEAR, AGE_MESS, MEDI_PREST, \
    TYPE_CONCENT


class Company(models.Model):
    name = models.CharField(max_length=50)
    type_id = models.CharField(max_length=10, choices=TYPE_ID, default=NIT)
    number_id = models.CharField(max_length=50)
    cod_verify = models.CharField(max_length=50, null=True, blank=True)
    name_rips = models.CharField(max_length=50)

    def __str__(self):
        return self.name


class Regional(models.Model):
    company = models.ForeignKey(Company, on_delete=models.CASCADE)
    name = models.CharField(max_length=50)
    cod_regional = models.IntegerField(unique=True)
    cod_hab = models.IntegerField()
    cod_dpto = models.CharField(max_length=10, choices=COD_DPTO)
    cod_muni = models.CharField(max_length=10, choices=COD_MUNI)
    urba_rul = models.CharField(max_length=10, choices=URBA_RUL, default=URBA)

    def __str__(self):
        return str(self.company)+'-'+self.name


class EPS(models.Model):
    cod_eps = models.CharField(max_length=50, unique=True)
    name = models.CharField(max_length=50)
    name_rips = models.CharField(max_length=50)

    def __str__(self):
        return self.name_rips


class Diagnostic(models.Model):
    cod_dx = models.CharField(max_length=50, unique=True)
    name = models.CharField(max_length=100)

    def __str__(self):
        return self.name


class Patient(models.Model):
    first_name = models.CharField(max_length=50)
    second_name = models.CharField(max_length=50)
    first_last_name = models.CharField(max_length=50)
    second_last_name = models.CharField(max_length=50)
    eps = models.ForeignKey(EPS, on_delete=models.CASCADE)
    gene = models.CharField(max_length=10, choices=GEN)
    type_id = models.CharField(max_length=10, choices=TYPE_ID, default=CEDULA)
    num_id = models.CharField(max_length=10, default=0)
    born_date = models.DateField(default=datetime.date(2000,1,1))
    diagnostic = models.ForeignKey(Diagnostic, on_delete=models.CASCADE)

    def _get_age (self):
        tiempo = (datetime.date.today() - self.born_date).days
        if tiempo // 365.2425 > 0:
            return int(tiempo // 365.2425)
        elif tiempo % 365.2425  > 30:
            return int((tiempo % 365.2425) // 30.64)
        else:
            return int(tiempo % 365.2425)

    def _get_age_messure(self):
        tiempo = (datetime.date.today() - self.born_date).days
        years =  tiempo // 365.2425
        months = tiempo % 365.2425
        if years > 0:
            mesure = 'años'
        elif months >= 30:
            mesure = 'meses'
        else:
            mesure = 'dias'
        return mesure

    def __str__(self):
        return self.first_name+' '+self.first_last_name+' - '+str(self.eps)

    age = property (_get_age)
    age_mess = property(_get_age_messure)



class Medicine(models.Model):
    cod_cum = models.CharField(max_length=50, unique=True)
    name = models.CharField(max_length=100)
    name_rips = models.CharField(max_length=50)
    is_pos = models.BooleanField()
    presentation = models.CharField(max_length=50, choices=MEDI_PREST)
    concentration = models.IntegerField()
    type_concent = models.CharField(max_length=50, choices=TYPE_CONCENT)
    price = models.FloatField(default=0.0)

    def __str__(self):
        return self.name


class Fact(models.Model):
    cod_fact = models.CharField(max_length=10, unique=True)
    patient = models.ForeignKey(Patient, on_delete=models.CASCADE)
    regional = models.ForeignKey(Regional, on_delete=models.CASCADE, default=2)
    aut_number = models.IntegerField()
    date_fact = models.DateField()
    cut_ini = models.DateField()
    cut_end = models.DateField()
    cero1to6 = models.IntegerField()

    def __str__(self):
        return self.cod_fact+' - '+str(self.patient)

class DetailMedi(models.Model):
    fact = models.ForeignKey(Fact, on_delete=models.CASCADE)
    medicine = models.ForeignKey(Medicine, on_delete=models.CASCADE)
    dosis = models.CharField(max_length=50)
    cant = models.IntegerField()

    def _subtotal(self):
        return self.cant * self.medicine.price

    subtotal = property(_subtotal)

    def __str__(self):
        return str(self.fact)+' - '+str(self.medicine)

class Treatement(models.Model):
    name = name = models.CharField(max_length=100)
    cod_treat = models.CharField(max_length=10, default=0)
    is_pos = models.BooleanField()
    especial = models.BooleanField(default=True)
    price = models.FloatField(default=0.0)

    def __str__(self):
        return self.name

class DetailTreat(models.Model):
    fact = models.ForeignKey(Fact, on_delete=models.CASCADE)
    treatement = models.ForeignKey(Treatement, on_delete=models.CASCADE)
    cant = models.IntegerField()

    def _subtotal(self):
        return self.cant * self.treatement.price

    subtotal = property(_subtotal)

    def __str__(self):
        return str(self.fact)+' - '+str(self.treatement)


class Laboratory(models.Model):
    name = models.CharField(max_length=100)
    codigo = models.CharField(max_length=100)
    price = models.FloatField()

    def __str__(self):
        return self.name


class DetailLabo(models.Model):
    fact = models.ForeignKey(Fact, on_delete=models.CASCADE)
    laboratory = models.ForeignKey(Laboratory, on_delete=models.CASCADE)
    cant = models.IntegerField()

    def _subtotal(self):
        return self.cant * self.laboratory.price


    def __str__(self):
        return str(self.fact) + ' - ' + str(self.laboratory)

    subtotal = property(_subtotal)