import datetime

from django.db import models

from .choices import CEDULA, NIT, TYPE_ID, URBA, URBA_RUL, GEN


class Company(models.Model):
    name = models.CharField(max_length=50)
    type_id = models.CharField(max_length=10, choices=TYPE_ID, default=NIT)
    number_id = models.CharField(max_length=50)
    cod_verify = models.CharField(max_length=50, null=True, blank=True)
    name_rips = models.CharField(max_length=50)

    def __str__(self):
        return self.name

class Departament(models.Model):
    name = models.CharField(max_length=50)
    name_rips = models.CharField(max_length=50, null=True)
    codigo = models.CharField(max_length=10)

    def __str__(self):
        return self.name

class Municipe(models.Model):
    name = models.CharField(max_length=50)
    name_rips = models.CharField(max_length=50, null=True)
    codigo = models.CharField(max_length=10)

    def __str__(self):
        return self.name

class Regional(models.Model):
    company = models.ForeignKey(Company, on_delete=models.CASCADE)
    name = models.CharField(max_length=50)
    name_rips = models.CharField(max_length=50, null=True)
    cod_regional = models.IntegerField(unique=True)
    cod_hab = models.IntegerField()
    dpto = models.ForeignKey(Departament, on_delete=models.CASCADE)
    municipe = models.ForeignKey(Municipe, on_delete=models.CASCADE)
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
    name_rips = models.CharField(max_length=50, null=True)

    def __str__(self):
        return self.name


class Patient(models.Model):
    first_name = models.CharField(max_length=50)
    second_name = models.CharField(max_length=50, null=True, blank=True)
    first_last_name = models.CharField(max_length=50)
    second_last_name = models.CharField(max_length=50, null=True, blank=True)
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


class Presentation(models.Model):
    name = models.CharField(max_length=50)
    name_rips = models.CharField(max_length=50, null=True)

    def __str__(self):
        return self.name

class Concentration(models.Model):
    name = models.CharField(max_length=50)
    name_rips = models.CharField(max_length=50, null=True)

    def __str__(self):
        return self.name


class Fact(models.Model):
    cod_fact = models.CharField(max_length=10, unique=True)
    patient = models.ForeignKey(Patient, on_delete=models.CASCADE)
    regional = models.ForeignKey(Regional, on_delete=models.CASCADE)
    aut_number = models.PositiveIntegerField()
    date_fact = models.DateField()
    cut_ini = models.DateField()
    cut_end = models.DateField()
    pin_elect = models.PositiveIntegerField(null=True, blank=True)
    validation = models.PositiveIntegerField(null=True, blank=True)
    cero1to6 = models.IntegerField()

    def __str__(self):
        return self.cod_fact+' - '+str(self.patient)


class Medicine(models.Model):
    cod_cum = models.CharField(max_length=50, unique=True)
    name = models.CharField(max_length=100)
    name_rips = models.CharField(max_length=50)
    is_pos = models.BooleanField()
    presentation = models.ForeignKey(Presentation, on_delete=models.CASCADE)
    concentration = models.ForeignKey(Concentration, on_delete=models.CASCADE)
    cant_concent = models.PositiveIntegerField()
    #price = models.FloatField(default=0.0)

    def __str__(self):
        return self.name


class PriceMedicine(models.Model):
    eps = models.ForeignKey(EPS, on_delete=models.CASCADE)
    medto = models.ForeignKey(Medicine, on_delete=models.CASCADE, related_name='pri_med')
    price = models.FloatField(default=0.0)

    class Meta:
        unique_together = ('eps', 'medto')

    def __str__(self):
        return self.eps.name_rips+' - '+self.medto.name_rips


class DetailMedi(models.Model):
    fact = models.ForeignKey(Fact, on_delete=models.CASCADE)
    medicine = models.ForeignKey(Medicine, on_delete=models.CASCADE)
    dosis = models.CharField(max_length=50)
    cant = models.IntegerField()

    def __str__(self):
        return str(self.fact)+' - '+str(self.medicine)

    def _price(self):
        return PriceMedicine.objects.filter(eps=self.fact.patient.eps, medto=self.medicine).first().price

    def _subtotal(self):
        return self.cant * self._price()

    subtotal = property(_subtotal)
    price = property(_price)

class DetailMediNoPos(models.Model):
    fact = models.ForeignKey(Fact, on_delete=models.CASCADE)
    medicine = models.ForeignKey(Medicine, on_delete=models.CASCADE)
    mipres = models.CharField(max_length=50)
    dosis = models.CharField(max_length=50)
    cant = models.IntegerField()

    def __str__(self):
        return str(self.fact)+' - '+str(self.medicine)+' No Pos'

    def _price(self):
        return PriceMedicine.objects.filter(eps=self.fact.patient.eps, medto=self.medicine).first().price

    def _subtotal(self):
        return self.cant * self._price()

    price = property(_price)
    subtotal = property(_subtotal)


class Dispositive(models.Model):
    name = models.CharField(max_length=100)
    name_rips = models.CharField(max_length=50, null=True)
    codigo = models.CharField(max_length=10, default=0)
    is_pos = models.BooleanField()
    especial = models.BooleanField(default=True)
    #price = models.FloatField(default=0.0)

    def __str__(self):
        return self.name


class PriceDispositive(models.Model):
    eps = models.ForeignKey(EPS, on_delete=models.CASCADE)
    dispo = models.ForeignKey(Dispositive, on_delete=models.CASCADE)
    price = models.FloatField(default=0.0)

    class Meta:
        unique_together = ('eps', 'dispo')

    def __str__(self):
        return self.eps.name_rips+' - '+self.dispo.name_rips



class DetailDispo(models.Model):
    fact = models.ForeignKey(Fact, on_delete=models.CASCADE)
    dispositive = models.ForeignKey(Dispositive, on_delete=models.CASCADE)
    cant = models.IntegerField()

    def __str__(self):
        return str(self.fact)+' - '+str(self.dispositive)

    def _price(self):
        return PriceDispositive.objects.filter(eps=self.fact.patient.eps, dispo=self.dispositive).first().price

    def _subtotal(self):
        return self.cant * self._price()

    price = property(_price)
    subtotal = property(_subtotal)



class Laboratory(models.Model):
    name = models.CharField(max_length=100)
    name_rips = models.CharField(max_length=50, null=True)
    codigo = models.CharField(max_length=100)
    #price = models.FloatField()

    def __str__(self):
        return self.name


class PriceLabo(models.Model):
    eps = models.ForeignKey(EPS, on_delete=models.CASCADE)
    labo = models.ForeignKey(Laboratory, on_delete=models.CASCADE)
    price = models.FloatField(default=0.0)

    class Meta:
        unique_together = ('eps', 'labo')

    def __str__(self):
        return self.eps.name_rips+' - '+self.labo.name_rips


class DetailLabo(models.Model):
    fact = models.ForeignKey(Fact, on_delete=models.CASCADE)
    laboratory = models.ForeignKey(Laboratory, on_delete=models.CASCADE)
    cant = models.IntegerField()

    def __str__(self):
        return str(self.fact) + ' - ' + str(self.laboratory)

    def _price(self):
        return PriceLabo.objects.filter(eps=self.fact.patient.eps, labo=self.laboratory).first().price

    def _subtotal(self):
        return self.cant * self._price()

    price = property(_price)
    subtotal = property(_subtotal)



class Service(models.Model):
    name = models.CharField(max_length=100)
    name_rips = models.CharField(max_length=100)
    codigo = models.CharField(max_length=100)
    #price = models.FloatField()


    def __str__(self):
        return self.name


class PriceService(models.Model):
    eps = models.ForeignKey(EPS, on_delete=models.CASCADE)
    servi = models.ForeignKey(Service, on_delete=models.CASCADE)
    price = models.FloatField(default=0.0)

    class Meta:
        unique_together = ('eps', 'servi')

    def __str__(self):
        return self.eps.name_rips+' - '+self.servi.name_rips


class DetailService(models.Model):
    fact = models.ForeignKey(Fact, on_delete=models.CASCADE)
    service = models.ForeignKey(Service, on_delete=models.CASCADE)
    cant = models.IntegerField()

    def __str__(self):
        return str(self.fact) + ' - ' + str(self.service)

    def _price(self):
        return PriceService.objects.filter(eps=self.fact.patient.eps, servi=self.service).first().price

    def _subtotal(self):
        return self.cant * self._price()

    price = property(_price)
    subtotal = property(_subtotal)
