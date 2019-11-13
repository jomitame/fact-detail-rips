CEDULA = 'CC'
REGISTRO = 'RC'
CED_EXT = 'CE'
TARJETA = 'TI'
NIT = 'NI'
URBA = 'U'
YEAR = '1'
TABLETA = 'T' # ojo cambiar
GRAJEA = 'G'



TYPE_ID = (
    (CEDULA,'Cédula'),
    (REGISTRO, 'Registo Civil'),
    (CED_EXT, 'Cedula Extranjera'),
    (TARJETA, 'Tarjeta Identidad'),
    (NIT, 'Nit')
)

COD_DPTO = (
    ('13','Bolivar'),
    ('08','Atlantico'),
)

COD_MUNI = (
    ('13001','Cartagena'),
    ('08001','Barranquilla'),
)

URBA_RUL = (
    ('U','Urbano'),
    ('R','Rural'),
)

GEN = (
    ('M', 'Masculino'),
    ('F', 'Femenino')
)

AGE_MESS = (
    (YEAR,'Años'),
    ('2','Meses'),
    ('3','Dias')
)

MEDI_PREST = (
    (TABLETA, 'Tableta'),
    (GRAJEA, 'Grajea')
)# ojo ampliar

TYPE_CONCENT = (
    ('ML', 'Mililitro'),
    ('MG', 'Miligramo')
)