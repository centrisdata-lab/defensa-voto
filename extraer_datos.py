import openpyxl, json, unicodedata

def normalize(s):
    s = str(s).strip().upper()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s

wb = openpyxl.load_workbook(
    'c:/Users/sara.munoz/OneDrive - arus.com.co/Analitica EO/AnalisisDescriptivo_Info/Lideres/Defensa del voto/TESTIGOS ESCRUTADORES.xlsx'
)

data = {}

# ---- OTROS MUNICIPIOS (col0=municipio, col1=comision texto, col2=nombre) ----
ws = wb['OTROS MUNICIPIOS']
for row in ws.iter_rows(min_row=2, values_only=True):
    municipio = row[0]; comision = row[1]; nombre = row[2]
    if not municipio or not comision or not nombre:
        continue
    municipio = normalize(municipio)
    comision_str = normalize(comision)
    nombre_str = normalize(nombre)

    if municipio not in data:
        data[municipio] = {'total_personas': 0, 'municipales': 0, 'auxiliares': 0,
                           'comisiones_unicas': set(), 'personas': []}
    data[municipio]['total_personas'] += 1
    if 'MUNICIPAL' in comision_str:
        data[municipio]['municipales'] += 1
    if 'AUXILIAR' in comision_str:
        data[municipio]['auxiliares'] += 1
    data[municipio]['comisiones_unicas'].add(comision_str)
    data[municipio]['personas'].append({'nombre': nombre_str, 'comision': comision_str})

# ---- MEDELLIN 2 (col0=numero comision, col1=nombre) ----
# Las comisiones son numeradas (1,2,3...) = auxiliares de Medellín
ws2 = wb['MEDELLIN 2']
municipio = 'MEDELLIN'
data[municipio] = {'total_personas': 0, 'municipales': 0, 'auxiliares': 0,
                   'comisiones_unicas': set(), 'personas': []}
for row in ws2.iter_rows(min_row=2, values_only=True):
    com_num = row[0]; nombre = row[1]
    if not nombre:
        continue
    nombre_str = normalize(nombre)
    comision_str = f'COMISION {com_num}' if com_num is not None else 'SIN COMISION'
    data[municipio]['total_personas'] += 1
    data[municipio]['auxiliares'] += 1
    data[municipio]['comisiones_unicas'].add(comision_str)
    data[municipio]['personas'].append({'nombre': nombre_str, 'comision': comision_str})

# Serializar
output = {}
for mun, d in data.items():
    output[mun] = {
        'total_personas': d['total_personas'],
        'municipales': d['municipales'],
        'auxiliares': d['auxiliares'],
        'total_comisiones': len(d['comisiones_unicas']),
        'comisiones': sorted(list(d['comisiones_unicas'])),
        'personas': d['personas']
    }

with open(
    'c:/Users/sara.munoz/OneDrive - arus.com.co/Analitica EO/AnalisisDescriptivo_Info/Lideres/Defensa del voto/escrutadores_data.json',
    'w', encoding='utf-8'
) as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

total = sum(d['total_personas'] for d in output.values())
print(f'OK - Municipios: {len(output)}, Total personas: {total}')
for mun, d in sorted(output.items()):
    print(f'  {mun}: {d["total_personas"]} personas, {d["municipales"]} municipales, {d["auxiliares"]} auxiliares')
