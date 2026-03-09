import json, unicodedata

def normalize(s):
    s = str(s).strip().upper()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s

name_map = {'CARMEN DE VIBORAL': 'EL CARMEN DE VIBORAL'}

with open('c:/Users/sara.munoz/OneDrive - arus.com.co/Analitica EO/AnalisisDescriptivo_Info/Lideres/Lideres-Referidos- mapa/antioquia_municipios_clean.json', 'r', encoding='utf-8') as f:
    gj = json.load(f)

with open('c:/Users/sara.munoz/OneDrive - arus.com.co/Analitica EO/AnalisisDescriptivo_Info/Lideres/Defensa del voto/escrutadores_data.json', 'r', encoding='utf-8') as f:
    esc = json.load(f)

geo_name_to_id = {}
for feat in gj['features']:
    geo_name_to_id[feat['properties']['MPIO_CNMBR']] = feat['id']

id_to_data = {}
for mun, d in esc.items():
    geo_name = name_map.get(mun, mun)
    gid = geo_name_to_id.get(geo_name)
    if gid:
        id_to_data[str(gid)] = {**d, 'display_name': mun}

total_personas = sum(d['total_personas'] for d in esc.values())
total_municipales = sum(d['municipales'] for d in esc.values())
total_auxiliares = sum(d['auxiliares'] for d in esc.values())
n_municipios = len(esc)

ranking = sorted(
    [(mun, d['total_comisiones'], d['municipales'], d['auxiliares']) for mun, d in esc.items()],
    key=lambda x: -x[1]
)

ranking_js_rows = ''
for mun, tc, muni, aux in ranking:
    geo_name = name_map.get(mun, mun)
    gid = geo_name_to_id.get(geo_name, '')
    ranking_js_rows += f"['{mun}', {tc}, {muni}, {aux}, {gid}],"

js_data = json.dumps(id_to_data, ensure_ascii=False)
geojson_str = json.dumps(gj, ensure_ascii=False)

tp_fmt = f"{total_personas:,}"
nm_fmt = str(n_municipios)
tm_fmt = str(total_municipales)
ta_fmt = str(total_auxiliares)

html_template = open('c:/Users/sara.munoz/OneDrive - arus.com.co/Analitica EO/AnalisisDescriptivo_Info/Lideres/Defensa del voto/template_mapa.html', 'r', encoding='utf-8').read()

html = html_template
html = html.replace('{{TOTAL_PERSONAS}}', tp_fmt)
html = html.replace('{{N_MUNICIPIOS}}', nm_fmt)
html = html.replace('{{TOTAL_MUNICIPALES}}', tm_fmt)
html = html.replace('{{TOTAL_AUXILIARES}}', ta_fmt)
html = html.replace('{{TOTAL_PERSONAS_2}}', tp_fmt)
html = html.replace('{{TOTAL_MUNICIPALES_2}}', tm_fmt)
html = html.replace('{{TOTAL_AUXILIARES_2}}', ta_fmt)
html = html.replace('ESCRUTADORES_DATA_PLACEHOLDER', js_data)
html = html.replace('RANKING_PLACEHOLDER', ranking_js_rows)
# Reemplazar primero PLACEHOLDER2 (más específico) y luego el genérico
html = html.replace('GEOJSON_PLACEHOLDER2', geojson_str)
html = html.replace('GEOJSON_PLACEHOLDER', geojson_str)

with open('c:/Users/sara.munoz/OneDrive - arus.com.co/Analitica EO/AnalisisDescriptivo_Info/Lideres/Defensa del voto/index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print('HTML generado OK - tamanio:', len(html), 'bytes')
