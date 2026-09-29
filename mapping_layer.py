from __future__ import annotations
import re
import pandas as pd
from typing import Any

VENDOR_PROFILES = {'Assurance': {'parcel_number': 'Parcel Numbers', 'use_code': 'Land Use Code', 'neighborhood': 'Valuation Zone', 'sale_price': 'Sale Price', 'building_class': 'Building Class', 'adjusted_area': 'Total Adj Area', 'grantor': 'Grantor', 'grantee': 'Grantee', 'deed_book': 'Deed Book', 'deed_page': 'Deed Page', 'sale_date': 'Sale Date', 'land_acreage': 'Land Acreage', 'living_area': 'Living Area', 'land_value': 'Land Value', 'improvement_value': 'Improvement Value', 'misc_value': 'Miscellaneous Improvement Value', 'total_value': 'Total Value of Property', 'qualification': 'Qualified Sale', 'sales_ratio': 'Individual Ratio', 'comments': 'Sale Notes'}, 'Assurance2': {'parcel_number': "Pcl #'s", 'sale_date': 'Sale Date', 'sales_ratio': 'Sale Ratio', 'grantor': 'Grantor', 'grantee': 'Grantee', 'type': 'Sale Type', 'qualification': 'Included In Study', 'sale_price': 'Price', 'use_code': 'Land Use Code', 'neighborhood': 'Valuation Zone', 'deed_book': 'Deed Book', 'deed_page': 'Deed Page', 'building_class': 'Building Class', 'living_area': 'Base Area', 'adjusted_area': 'Adj Area', 'improvement_value': 'Bld Value', 'land_value': 'Land Value', 'misc_value': 'Oby Value', 'total_value': 'Total Value', 'comments': 'Notes', 'land_acreage': 'Acre'}, 'Delta': {'parcel_number': 'PARCEL_NUMBER', 'use_code': 'USE CODE', 'sale_price': 'SALE_PRICE', 'building_class': 'IMPROVEMENT_CLASS', 'living_area': 'LIVING_AREA', 'adjusted_area': 'TOTAL_ADJUSTED_AREA', 'grantor': 'GRANTOR', 'grantee': 'GRANTEE', 'deed_book': 'DEED_BOOK', 'deed_page': 'DEED_BOOK_PAGE', 'sale_date_year': 'SALE_DATE_YEAR', 'sale_date_month': 'SALE_DATE_MONTH', 'sale_date_day': 'SALE_DATE_DAY', 'land_acreage': 'LAND_ACRES', 'comments': 'COMMENTS', 'total_value': 'APPRAISED_VALUE', 'qualification': 'QUALIFIED_SALE', 'land_value': 'LAND_VALUE', 'improvement_value': 'IMPROVEMENT_VALUE', 'misc_value': 'MISC_VALUE', 'type': 'SALE_TYPE', 'neighborhood': 'NEIGHBORHOOD'}, 'S&W': {'parcel_number': 'PARCELNO', 'use_code': 'USECODE', 'neighborhood': 'NBHD', 'deed_book': 'D_BOOK', 'deed_page': 'D_BPAGE', 'sale_date': 'D_DATE', 'land_acreage': 'C_ACRE', 'grantor': 'GONAME', 'grantee': 'GENAME', 'building_class': 'BLDG_CLASS', 'living_area': 'LIVINGAREA', 'adjusted_area': 'TAA', 'total_value': 'APPRVAL', 'land_value': 'LANDVAL', 'improvement_value': 'IMPRVAL', 'sale_price': 'SALE_PRICE', 'sales_ratio': 'RATIO', 'qualification': 'USABLE_SALE'}, 'Capture': {'parcel_number': 'PARCEL', 'neighborhood': 'NBRHD', 'building_class': 'GRADE', 'adjusted_area': 'TOTAL ADJ AREA', 'use_code': 'LAND USE', 'land_value': 'LAND VALUE', 'improvement_value': 'BLDG VALUE', 'misc_value': 'MISC VALUE', 'total_value': 'APPR VALUE', 'sale_price': 'SALE PRICE', 'sale_date': 'SALE DATE', 'deed_book': 'RECD BOOK', 'deed_page': 'RECD PAGE', 'comments': 'SALE COMMENT 1', 'sales_ratio': 'SALE RATIO', 'type': 'SALE TYPE', 'qualification': 'RATIOABLE', 'grantor': 'GRANTOR', 'grantee': 'GRANTEE', 'living_area': 'TOTAL LIVING AREA'}, 'Ingeunity': {'parcel_number': 'Parcels', 'use_code': 'LandUseCodes', 'neighborhood': 'NbhCode', 'sale_price': 'Sale Price', 'building_class': 'Bldg Class', 'adjusted_area': 'TAA', 'grantor': 'Grantor', 'grantee': 'Grantee', 'deed_book': 'Book', 'deed_page': 'Page', 'sale_date': 'Sale Date', 'land_acreage': 'Land Acreage', 'living_area': 'Living Area', 'land_value': 'Land Value', 'improvement_value': 'Improvement Value', 'misc_value': 'Misc Improvements Value', 'total_value': 'Total Value', 'qualification': 'Qualified Sale', 'comments': 'Comment'}}
FIELD_ALIASES = {'parcel_number': ['Parcel Numbers', "Pcl #'s", 'PARCEL_NUMBER', 'PARCELNO', 'PARCEL', 'Parcels', 'Parcel Number'], 'use_code': ['Land Use Code', 'USE CODE', 'USECODE', 'LAND USE', 'LandUseCodes', 'Improvement Code', 'Property Use Code'], 'neighborhood': ['Valuation Zone', 'NEIGHBORHOOD', 'NBHD', 'NBRHD', 'NbhCode', 'Neighborhood'], 'sale_price': ['Sale Price', 'Price', 'SALE_PRICE', 'SALE PRICE', 'PRICE'], 'building_class': ['Building Class', 'IMPROVEMENT_CLASS', 'BLDG_CLASS', 'GRADE', 'Bldg Class'], 'adjusted_area': ['Total Adj Area', 'Adj Area', 'TOTAL_ADJUSTED_AREA', 'TAA', 'TOTAL ADJ AREA'], 'grantor': ['Grantor', 'GRANTOR', 'GONAME'], 'grantee': ['Grantee', 'GRANTEE', 'GENAME'], 'deed_book': ['Deed Book', 'DEED_BOOK', 'D_BOOK', 'RECD BOOK', 'Book'], 'deed_page': ['Deed Page', 'DEED_BOOK_PAGE', 'D_BPAGE', 'RECD PAGE', 'Page'], 'sale_date': ['Sale Date', 'D_DATE', 'SALE DATE'], 'land_acreage': ['Land Acreage', 'Acre', 'LAND_ACRES', 'C_ACRE'], 'living_area': ['Living Area', 'Base Area', 'LIVING_AREA', 'LIVINGAREA', 'TOTAL LIVING AREA'], 'land_value': ['Land Value', 'LAND_VALUE', 'LANDVAL', 'LAND VALUE'], 'improvement_value': ['Improvement Value', 'Bld Value', 'IMPROVEMENT_VALUE', 'IMPRVAL', 'BLDG VALUE'], 'misc_value': ['Miscellaneous Improvement Value', 'Oby Value', 'MISC_VALUE', 'MISC VALUE', 'Misc Improvements Value'], 'total_value': ['Total Value of Property', 'Total Value', 'APPRAISED_VALUE', 'APPRVAL', 'APPR VALUE'], 'qualification': ['Qualified Sale', 'Included In Study', 'QUALIFIED_SALE', 'USABLE_SALE', 'RATIOABLE'], 'sales_ratio': ['Individual Ratio', 'Sale Ratio', 'RATIO', 'SALE RATIO'], 'comments': ['Sale Notes', 'Notes', 'COMMENTS', 'SALE COMMENT 1', 'Comment'], 'type': ['Sale Type', 'SALE_TYPE', 'SALE TYPE'], 'sale_date_year': ['SALE_DATE_YEAR'], 'sale_date_month': ['SALE_DATE_MONTH'], 'sale_date_day': ['SALE_DATE_DAY']}
CANONICAL_NAMES = {
'parcel_number':'Parcel Number','use_code':'Use Code','neighborhood':'Neighborhood','sale_price':'Sale Price','building_class':'Building Class','adjusted_area':'Total Adjusted Area','living_area':'Living Area','grantor':'Grantor','grantee':'Grantee','deed_book':'Deed Book','deed_page':'Deed Page','sale_date':'Sale Date','land_acreage':'Land Acreage','land_value':'Land Value','improvement_value':'Improvement Value','misc_value':'Miscellaneous Value','total_value':'Total Value','qualification':'Qualification','sales_ratio':'Sales Ratio','comments':'Comments','type':'Type'}

def normalize_header(v:Any)->str:
 s='' if v is None else str(v).strip().upper()
 for ch in (' ','_','-','/','.',"'",'#'): s=s.replace(ch,'')
 return s

def _blank(v):
 if v is None:return True
 try:
  if pd.isna(v):return True
 except Exception:pass
 return not str(v).strip()

def _header_positions(columns):
 d={}
 for i,c in enumerate(columns): d.setdefault(normalize_header(c),[]).append(i)
 return d

def detect_vendor(columns):
 pos=_header_positions(columns); scores=[]
 for vendor,profile in VENDOR_PROFILES.items():
  wanted={normalize_header(x) for x in profile.values()}
  hit=sum(1 for x in wanted if x in pos)
  scores.append((hit/max(len(wanted),1),hit,vendor))
 scores.sort(reverse=True)
 best=scores[0]
 return {'vendor':best[2] if best[0]>=0.55 else 'Unknown','confidence':round(best[0],3),'matched_fields':best[1],'ranking':scores}

def _find_columns(df, names):
 norms={normalize_header(n) for n in names}; return [c for c in df.columns if normalize_header(c) in norms]

def _pop(s): return int((~s.map(_blank)).sum())
def _valid(s,field):
 if field in {'sale_price','total_value','land_value','improvement_value','misc_value','sales_ratio','adjusted_area','living_area','land_acreage'}: return int(pd.to_numeric(s,errors='coerce').notna().sum())
 if field=='sale_date': return int(pd.to_datetime(s,errors='coerce').notna().sum())
 if field=='use_code':
  return int(s.map(lambda v: bool(re.fullmatch(r'\d{3,4}(?:\.0+)?',str(v).strip())) if not _blank(v) else False).sum())
 return _pop(s)

def choose(df,field,cols):
 if not cols:return None
 return max(cols,key=lambda c:(_valid(df[c],field),_pop(df[c]),-list(df.columns).index(c)))

DERIVED_RATIO_OPTION = "Calculate: Total Value / Sale Price"
DO_NOT_MAP_OPTION = "Do Not Map"


def mapping_options(df, field, report_row=None):
    """Return user-selectable source columns for a canonical field."""
    candidates = []
    if report_row is not None:
        raw = str(report_row.get('Candidate Headers', '') or '')
        candidates.extend([x.strip() for x in raw.split('|') if x.strip()])
        detected = str(report_row.get('Detected Header', '') or '').strip()
        if detected and not detected.startswith('DERIVED:'):
            candidates.insert(0, detected)
    # Let the analyst override to any raw source field, not just recognized aliases.
    candidates.extend([str(c) for c in df.columns])
    seen=[]
    for c in candidates:
        if c not in seen: seen.append(c)
    if field == 'sales_ratio':
        seen.append(DERIVED_RATIO_OPTION)
    seen.append(DO_NOT_MAP_OPTION)
    return seen


def map_report(df):
 det=detect_vendor(df.columns); vendor=det['vendor']; profile=VENDOR_PROFILES.get(vendor,{})
 rows=[]
 fields=set(FIELD_ALIASES)|set(profile)
 for field in sorted(fields):
  exact=[]
  if field in profile: exact=_find_columns(df,[profile[field]])
  cand=exact or _find_columns(df,FIELD_ALIASES.get(field,[]))
  sel=choose(df,field,cand); p=_pop(df[sel]) if sel else 0
  rows.append({'Vendor':vendor,'System Field':field,'Canonical Header':CANONICAL_NAMES.get(field,field),'Detected Header':sel or '', 'Status':'Mapped' if sel and p else ('Empty' if sel else 'Missing'),'Candidate Headers':' | '.join(map(str,cand)),'Populated Rows':p,'Method':'Exact vendor profile' if exact else ('VBA/general alias fallback' if cand else 'No match')})
 return pd.DataFrame(rows),det


def canonicalize(df, overrides=None):
 report,det=map_report(df); out=df.copy(); vendor=det['vendor']
 overrides = overrides or {}

 # Apply explicit user choices first; otherwise use the mapper's recommendation.
 for i,r in report.iterrows():
  f=r['System Field']; target=CANONICAL_NAMES.get(f)
  if not target: continue
  chosen = overrides.get(f, r['Detected Header'])
  if chosen == DO_NOT_MAP_OPTION:
   report.loc[i, ['Detected Header','Status','Method']] = [DO_NOT_MAP_OPTION, 'Not Mapped', 'User selection']
   if target in out.columns and target not in df.columns:
    out.drop(columns=[target], inplace=True)
   continue
  if f == 'sales_ratio' and chosen == DERIVED_RATIO_OPTION:
   report.loc[i, ['Detected Header','Status','Method']] = ['DERIVED: Total Value / Sale Price','Pending','User selection']
   continue
  if chosen in df.columns:
   out[target] = df[chosen]
   report.loc[i, 'Detected Header'] = chosen
   report.loc[i, 'Status'] = 'Mapped' if _pop(df[chosen]) else 'Empty'
   if f in overrides: report.loc[i, 'Method'] = 'User selection'

 # Delta default remains USE CODE when the analyst has not explicitly overridden it.
 if vendor == 'Delta' and 'use_code' not in overrides:
  delta_use = next((c for c in df.columns if normalize_header(c) == normalize_header('USE CODE')), None)
  if delta_use is not None:
   out['Use Code'] = df[delta_use]
   mask=report['System Field'].eq('use_code')
   report.loc[mask,'Detected Header']=delta_use
   report.loc[mask,'Status']='Mapped'
   report.loc[mask,'Method']='Exact Delta USE CODE rule'

 # Derive ratio when explicitly selected OR when no usable mapped ratio exists.
 ratio_choice = overrides.get('sales_ratio')
 need_ratio = ratio_choice == DERIVED_RATIO_OPTION or ('Sales Ratio' not in out.columns or pd.to_numeric(out['Sales Ratio'], errors='coerce').notna().sum() == 0)
 if ratio_choice == DO_NOT_MAP_OPTION:
  need_ratio = False
 if need_ratio and 'Total Value' in out.columns and 'Sale Price' in out.columns:
  tv = pd.to_numeric(out['Total Value'], errors='coerce')
  sp = pd.to_numeric(out['Sale Price'], errors='coerce')
  out['Sales Ratio'] = (tv / sp.where(sp != 0)).replace([float('inf'), float('-inf')], pd.NA)
  mask = report['System Field'].eq('sales_ratio')
  if mask.any():
   report.loc[mask, 'Detected Header'] = 'DERIVED: Total Value / Sale Price'
   report.loc[mask, 'Status'] = 'Mapped'
   report.loc[mask, 'Method'] = 'Calculated' if 'sales_ratio' not in overrides else 'User selection: calculated'

 # exact VBA date priority unless the user explicitly selected a date source.
 if 'sale_date' not in overrides:
  pos=_header_positions(df.columns)
  def col(n):
   xs=pos.get(normalize_header(n),[]); return df.columns[xs[0]] if xs else None
  d=col('D_DATE'); sd=col('Sale Date')
  if d is not None: out['Sale Date']=pd.to_datetime(df[d],errors='coerce')
  elif sd is not None: out['Sale Date']=pd.to_datetime(df[sd],errors='coerce')
  else:
   y=col('SALE_DATE_YEAR');m=col('SALE_DATE_MONTH');day=col('SALE_DATE_DAY')
   if y and m and day:
    out['Sale Date']=pd.to_datetime(pd.DataFrame({'year':pd.to_numeric(df[y],errors='coerce'),'month':pd.to_numeric(df[m],errors='coerce'),'day':pd.to_numeric(df[day],errors='coerce')}),errors='coerce')
 return out,report,det
