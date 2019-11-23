CEDULA = 'CC'
REGISTRO = 'RC'
CED_EXT = 'CE'
TARJETA = 'TI'
NIT = 'NI'
URBA = 'U'
TABLETA = 'T' # ojo cambiar
GRAJEA = 'G'
CONTRIBUTIVO = '1'



TYPE_ID = (
    (CEDULA,'Cédula'),
    (REGISTRO, 'Registo Civil'),
    (CED_EXT, 'Cedula Extranjera'),
    (TARJETA, 'Tarjeta Identidad'),
    (NIT, 'Nit')
)

URBA_RUL = (
    ('U','Urbano'),
    ('R','Rural'),
)

GEN = (
    ('M', 'Masculino'),
    ('F', 'Femenino')
)

REG = (
    (CONTRIBUTIVO,'Contributivo'),
    ('2','Subsidiado')
)