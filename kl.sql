SELECT 
to_char(VAS_DATUM,'YYYYMM')                                as EV_HO,
to_char(VAS_DATUM,'YYYY')||'HET'|| to_char(VAS_DATUM,'WW') as EV_HET,
  V_PART_ID,
  V_NAME,
  V_MT_ID,
  IGAZGATOSAG,
  UZLETAG,
  AGAZAT,
  UGYF_CSOP_KOD,
  UGYF_CSOP_NEV,
  BSS_VEVO_ID,
  ACCO_ID,
  VAS_DATUM,
  TIME_KEY,
  EVENT_DATE_TIME,
  KESZ_NEV,
  KESZ_KOD,
  DB,
  AR,
  KEDV,
  FIZ_NETTO_AR,
  KOLTS 
from kecskemetil.bbu_kesz_vas_2011
where to_char(VAS_DATUM,'YYYY') =  '2012' 
