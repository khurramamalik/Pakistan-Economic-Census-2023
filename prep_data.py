# -*- coding: utf-8 -*-
"""Parse PBS Economic Census Excel files + admin boundaries into JS data files
for the dashboard. Run:  py prep_data.py
"""
import json, re, glob, os
import pandas as pd

BASE = r'E:\Data\PBS Economic Census\DB'
OUT = os.path.join(BASE, 'Dashboard', 'data')
os.makedirs(OUT, exist_ok=True)

# ---------------------------------------------------------------- sectors
SECTORS = {
    1:  ('Agriculture, Forestry & Fishing', '01-03'),
    2:  ('Mining & Quarrying', '05-09'),
    3:  ('Manufacturing', '10-33'),
    4:  ('Electricity, Gas & AC Supply', '35'),
    5:  ('Water Supply & Waste Mgmt', '36-39'),
    6:  ('Construction', '41-43'),
    7:  ('Wholesale & Retail Trade', '45-47'),
    8:  ('Transportation & Storage', '49-53'),
    9:  ('Accommodation & Food Services', '55-56'),
    10: ('Information & Communication', '58-63'),
    11: ('Financial & Insurance', '64-66'),
    12: ('Real Estate', '68'),
    13: ('Professional & Technical', '69-75'),
    14: ('Admin & Support Services', '77-82'),
    15: ('Public Admin & Defence', '84'),
    16: ('Education', '85'),
    17: ('Health & Social Work', '86-88'),
    18: ('Arts & Recreation', '90-93'),
    19: ('Other Services', '94-96'),
    20: ('Extraterritorial Bodies', '99'),
    21: ('Others / Unclassified', 'n.e.c.'),
}
DESC2SEC = {
    'agriculture, forestry and fishing': 1,
    'mining and quarrying': 2,
    'manufacturing': 3,
    'electricity, gas, steam and air conditioning supply': 4,
    'water supply; sewerage, waste management and remediation activities': 5,
    'wholesale and retail trade; repair of motor vehicles and motorcycles': 7,
    'construction': 6,
    'transportation and storage': 8,
    'accommodation and food service activities': 9,
    'information and communication': 10,
    'financial and insurance activities': 11,
    'real estate activities': 12,
    'professional, scientific and technical activities': 13,
    'administrative and support service activities': 14,
    'public administration and defence; compulsory social security': 15,
    'education': 16,
    'human health and social work activities': 17,
    'arts, entertainment and recreation': 18,
    'other service activities': 19,
    'activities of extraterritorial organizations and bodies': 20,
    'others': 21,
}
UNITS = {
    2: 'Old Homes & Orphanages', 3: 'Hostels', 4: 'Hotels',
    6: 'Health Establishments', 7: 'Madrassas', 8: 'Schools', 9: 'Colleges',
    10: 'Universities', 11: 'Masjids', 12: 'Retail Shops', 13: 'Wholesale Shops',
    14: 'Service Shops', 15: 'Production Shops', 16: 'Factories',
    17: 'Semi-Govt. Offices', 18: 'Govt. Offices', 19: 'Post Offices & Couriers',
    20: 'Banks', 21: 'Police Stations', 22: 'Cattle Farms', 24: 'Others',
}

PSIC_FILES = {
    'Punjab': 'PUNJAB-DISTRICTS-PSIC-wise-1.xlsx',
    'Sindh': 'SINDH-DISTRICTS-PSIC-WISE.xlsx',
    'Khyber Pakhtunkhwa': 'KPK-DISTRICTPSIC-WISE.xlsx',
    'Balochistan': 'BALOCHISTAN-DISTRICTPSIC-WISE.xlsx',
}
UNIT_FILES = {
    'Punjab': 'Punjab-district-unit-type.xlsx',
    'Sindh': 'Sindh-districtsunit-type-1.xlsx',
    'Khyber Pakhtunkhwa': 'kpk-districtsUnit-Type.xlsx',
    'Balochistan': 'district-balochistan-unit-type.xlsx',
}

def normkey(s):
    return re.sub(r'[^A-Z]', '', s.upper())

# census-name (normalized) -> geojson adm2_name; None = no polygon exists
ALIAS = {
    # Balochistan
    'GAWADAR': 'Gwadar', 'SOHBATOUR': 'Sohbatpur', 'ZOHB': 'Zhob',
    'SURAB': None,
    # KPK
    'ABBOTABAD': 'Abbottabad', 'BATGRAM': 'Batagram',
    'LOWERCHITRAL': 'Chitral Lower', 'UPPERCHITRAL': 'Chitral Upper',
    'DERAISMAILKHAN': 'D. I. Khan',
    'LOWERKOHISTAN': 'Kohistan Lower', 'UPPERKOHISTAN': 'Kohistan Upper',
    'NORTHWAZIRSTAN': 'North Waziristan', 'SOUTHWAZIRSTAN': 'South Waziristan',
    'NORTHWAZIRISTAN': 'North Waziristan', 'SOUTHWAZIRISTAN': 'South Waziristan',
    'NOWSHEHRA': 'Nowshera', 'SAWABI': 'Swabi', 'SAWAT': 'Swat',
    'TORGHAR': 'Tor Ghar', 'MALAKANDPROTECTEDAREA': 'Malakand',
    # Punjab
    'JHUNG': 'Jhang', 'JEHLUM': 'Jhelum', 'LAYYAH': 'Leiah',
    'MANDIBAHAUDIN': 'Mandi Bahauddin', 'PAKPATAN': 'Pakpattan',
    # Sindh
    'KARACHIEAST': 'East Karachi', 'KARACHIWEST': 'West Karachi',
    'KARACHICENTRAL': 'Central Karachi', 'KARACHISOUTH': 'South Karachi',
    'MALIR': 'Malir Karachi', 'KORANGI': 'Korangi Karachi',
    'SHAHEEDBENAZIRABAD': 'Shaheed Benazir Abad',
    'MIRPUR': 'Mirpur Khas', 'NAUSHAHRO': 'Naushahro Feroze',
    'KEAMARI': None,
}
# display names for districts without a polygon
NO_GEO_NAME = {'SURAB': 'Surab', 'KEAMARI': 'Keamari'}

# ---------------------------------------------------------------- geo names
with open(os.path.join(BASE, 'Administrative Boundaries', 'pak_admin2.geojson'), encoding='utf-8') as fh:
    ADM2 = json.load(fh)
with open(os.path.join(BASE, 'Administrative Boundaries', 'pak_admin1.geojson'), encoding='utf-8') as fh:
    ADM1 = json.load(fh)

GEO_BY_PROV = {}
for ft in ADM2['features']:
    p = ft['properties']
    GEO_BY_PROV.setdefault(p['adm1_name'], {})[normkey(p['adm2_name'])] = p['adm2_name']

def resolve(census_name, province):
    """-> (district_key, has_geo) ; district_key is the geojson name when matched."""
    nm = census_name.strip()
    # strip repeated 'DISTRICT' suffixes and footnote junk
    while re.search(r'\s+DISTRICT\s*$', nm, re.I):
        nm = re.sub(r'\s+DISTRICT\s*$', '', nm, flags=re.I)
    k = normkey(nm)
    if k in ALIAS:
        tgt = ALIAS[k]
        if tgt is None:
            return NO_GEO_NAME[k], False
        return tgt, True
    prov_geo = GEO_BY_PROV[province]
    if k in prov_geo:
        return prov_geo[k], True
    raise KeyError(f'UNMATCHED district {census_name!r} ({province})')

def is_footnote(s):
    s2 = s.lower()
    return s.startswith('v\xa0') or 'unit type includes' in s2 or 'post office' in s2 and 'encompasses' in s2 or 'police station includes' in s2

def num(v):
    if pd.isna(v):
        return 0
    return int(round(float(v)))

districts = {}  # key -> record

def rec(key, province, has_geo):
    if key not in districts:
        districts[key] = {'name': key, 'province': province, 'geo': has_geo,
                          'psic': {}, 'units': {}}
    return districts[key]

# ---------------------------------------------------------------- parse PSIC
for province, fname in PSIC_FILES.items():
    df = pd.read_excel(os.path.join(BASE, 'PSIC Code', fname), header=None)
    cur = None
    for _, r in df.iterrows():
        c0, c1, c3, c4 = r[0], r[1], r[3], r[4]
        if isinstance(c0, str) and pd.isna(c1) and pd.isna(c3):
            s = c0.strip()
            if s.startswith('Table') or s == 'Sr. No.' or is_footnote(s):
                continue
            key, has_geo = resolve(s, province)
            cur = rec(key, province, has_geo)
        elif cur is not None and isinstance(c1, str) and pd.notna(c3):
            desc = re.sub(r'\s+', ' ', c1.strip().lower())
            sec = DESC2SEC.get(desc)
            if sec is None:
                print(f'  !! unknown PSIC description: {c1!r} in {fname}')
                continue
            cur['psic'][sec] = [num(c3), num(c4)]

# ---------------------------------------------------------------- parse units
for province, fname in UNIT_FILES.items():
    df = pd.read_excel(os.path.join(BASE, 'Unit Type', fname), header=None)
    cur = None
    for _, r in df.iterrows():
        c0, c1, c2, c3 = r[0], r[1], r[2], r[3]
        if isinstance(c0, str) and pd.isna(c1) and pd.isna(c2):
            s = c0.strip()
            if s.upper().startswith('TABLE') or is_footnote(s):
                continue
            key, has_geo = resolve(s, province)
            cur = rec(key, province, has_geo)
        elif cur is not None and isinstance(c0, str) and pd.notna(c1) and c0.strip() != 'Description':
            try:
                code = int(c1)
            except (TypeError, ValueError):
                continue
            if code not in UNITS:
                print(f'  !! unknown unit code {c1!r} ({c0!r}) in {fname}')
                continue
            prev = cur['units'].get(code, [0, 0])
            cur['units'][code] = [prev[0] + num(c2), prev[1] + num(c3)]

# ---------------------------------------------------------------- report
provinces = ['Punjab', 'Sindh', 'Khyber Pakhtunkhwa', 'Balochistan']
print(f'parsed {len(districts)} districts')
for p in provinces:
    ds = [d for d in districts.values() if d['province'] == p]
    n_psic = sum(1 for d in ds if d['psic'])
    n_unit = sum(1 for d in ds if d['units'])
    n_geo = sum(1 for d in ds if d['geo'])
    print(f'  {p}: {len(ds)} districts | psic {n_psic} | units {n_unit} | with geometry {n_geo}')
    for d in ds:
        if not d['psic']:
            print(f'    (no PSIC data) {d["name"]}')
        if not d['units']:
            print(f'    (no unit data) {d["name"]}')

# cross-check district totals vs stated totals is skipped (stated totals row parsed out);
# print national sums as sanity check
te = sum(v[0] for d in districts.values() for v in d['psic'].values())
tw = sum(v[1] for d in districts.values() for v in d['psic'].values())
print(f'PSIC national: establishments={te:,} workforce={tw:,}')
tue = sum(v[0] for d in districts.values() for v in d['units'].values())
tuw = sum(v[1] for d in districts.values() for v in d['units'].values())
print(f'Units national: establishments={tue:,} workforce={tuw:,}')

# ---------------------------------------------------------------- geometry
def q(x):
    return round(x, 4)

def quantize_ring(ring):
    out = []
    prev = None
    for pt in ring:
        p2 = [q(pt[0]), q(pt[1])]
        if p2 != prev:
            out.append(p2)
            prev = p2
    if len(out) > 1 and out[0] != out[-1]:
        out.append(out[0])
    return out

def quantize_geom(geom):
    t = geom['type']
    if t == 'Polygon':
        return {'type': t, 'coordinates': [quantize_ring(r) for r in geom['coordinates']]}
    if t == 'MultiPolygon':
        return {'type': t, 'coordinates': [[quantize_ring(r) for r in poly] for poly in geom['coordinates']]}
    return geom

geo_feats = []
matched_geo = set()
for ft in ADM2['features']:
    p = ft['properties']
    if p['adm1_name'] not in provinces:
        continue
    name = p['adm2_name']
    has_data = name in districts
    if has_data:
        matched_geo.add(name)
    geo_feats.append({'type': 'Feature',
                      'properties': {'n': name, 'p': p['adm1_name']},
                      'geometry': quantize_geom(ft['geometry'])})
print('geo polygons without census data:',
      sorted(f['properties']['n'] for f in geo_feats if f['properties']['n'] not in districts))

prov_feats = []
for ft in ADM1['features']:
    prov_feats.append({'type': 'Feature',
                       'properties': {'n': ft['properties']['adm1_name']},
                       'geometry': quantize_geom(ft['geometry'])})

# ---------------------------------------------------------------- write
census = {
    'sectors': {str(k): {'label': v[0], 'psic': v[1]} for k, v in SECTORS.items()},
    'units': {str(k): v for k, v in UNITS.items()},
    'provinces': provinces,
    'districts': districts,
    'merge_into': {'Keamari': 'West Karachi'},  # Keamari split from West Karachi after boundary vintage
    'no_geo_note': {'Surab': 'Carved out of Kalat; boundary not in 2022 admin file',
                    'Keamari': 'Shown merged into West Karachi on the map'},
}
with open(os.path.join(OUT, 'census.js'), 'w', encoding='utf-8') as fh:
    fh.write('window.CENSUS=')
    json.dump(census, fh, separators=(',', ':'))
    fh.write(';')

with open(os.path.join(OUT, 'geo.js'), 'w', encoding='utf-8') as fh:
    fh.write('window.GEO_DISTRICTS=')
    json.dump({'type': 'FeatureCollection', 'features': geo_feats}, fh, separators=(',', ':'))
    fh.write(';\nwindow.GEO_PROVINCES=')
    json.dump({'type': 'FeatureCollection', 'features': prov_feats}, fh, separators=(',', ':'))
    fh.write(';')

for f in ('census.js', 'geo.js'):
    sz = os.path.getsize(os.path.join(OUT, f))
    print(f'{f}: {sz/1e6:.2f} MB')
print('done')
