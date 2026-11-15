from configparser import ConfigParser, ExtendedInterpolation
import os
import sys

CONFIG_FILE = "config.ini"
if not os.path.exists(CONFIG_FILE):
    print("File config tidak ditemukan.")
    sys.exit()

config = ConfigParser(interpolation=ExtendedInterpolation())
config.read(CONFIG_FILE)


def get_config(section, option, type, default=None):
    if config.has_section(section):
        if config.has_option(section, option):
            if type == "bool":
                return config.getboolean(section, option)
            elif type == "int":
                return config.getint(section, option)
            elif type == "str":
                return config.get(section, option)
            elif type == "float":
                return config.getfloat(section, option)
        else:
            return default
    else:
        return default


STN_CODE = config.getint("STATION", "STN_CODE")
STN_ID = config.get("STATION", "STN_ID")

ME48_WS_NAME = config.get("EXCEL", "ME48_WS_NAME")
PERCELL_DIFF = config.getint("EXCEL", "PERCELL_DIFF")
PERPAGE_DIFF = config.getint("EXCEL", "PERPAGE_DIFF")
UTC_DIFF = config.getint("EXCEL", "UTC_DIFF")
PATH_TO_SYNOP = config.get("EXCEL", "PATH_TO_SYNOP")

DB_NAME = config.get("DATABASE", "DB_NAME")
ME48_TABLE_NAME = config.get("DATABASE", "ME48_TABLE_NAME")

OBSERVER = dict(config.items("OBSERVER"))

BMKGSABANG_TOKEN = config.get("BMKGSATU", "BMKGSABANG_TOKEN")
SINOPTIK_TOKEN = config.get("BMKGSATU", "SINOPTIK_TOKEN")
DEBUG = get_config("OTHER", "DEBUG", "bool", False)
SYNOP_HOURS = [int(i.strip()) for i in get_config("EXCEL", "SYNOP_HOURS", "str").split(',')]
SYNOP_HEADING_UTAMA = config.get("BMKGSATU", "SYNOP_HEADING_UTAMA")
SYNOP_HEADING_ANTARA = config.get("BMKGSATU", "SYNOP_HEADING_ANTARA")
SYNOP_HEADING_OTHER = config.get("BMKGSATU", "SYNOP_HEADING_OTHER")

BMKGSATU_URL = config.get("BMKGSATU", "BMKGSATU_URL")
BMKGSABANG_URL = config.get("BMKGSATU", "BMKGSABANG_URL")

MONTHS = {
    1: ['JAN', 'JANUARI'],
    2: ['FEB', 'FEBRUARI'],
    3: ['MAR', 'MARET'],
    4: ['APR', 'APRIL'],
    5: ['MEI', 'MEI'],
    6: ['JUN', 'JUNI'],
    7: ['JUL', 'JULI'],
    8: ['AGS', 'AGUSTUS'],
    9: ['SEP', 'SEPTEMBER'],
    10: ['OKT', 'OKTOBER'],
    11: ['NOV', 'NOVEMBER'],
    12: ['DES', 'DESEMBER'],
}

CANNOT_EMPTY_LIST = [
    "pressure_qff_mb_derived",
    "pressure_qfe_mb_derived",
    "temp_drybulb_c_tttttt",
    "temp_wetbulb_c"
]

ME48_LIST = [
    "wind_indicator_iw",
    "wind_dir_deg_dd",
    "wind_speed_ff",
    "visibility_vv",
    "weather_indicator_ix",
    "present_weather_ww",
    "past_weather_w1",
    "past_weather_w2",
    "pressure_temp_c",
    "pressure_reading_mb",
    "pressure_tend_3h_mb_ppp",
    "pressure_qff_mb_derived",
    "pressure_qfe_mb_derived",
    "pressure_tend_24h_mb_ppp",
    "temp_drybulb_c_tttttt",
    "temp_wetbulb_c",
    "temp_dewpoint_c_tdtdtd",
    "relative_humidity_pc",
    "temp_max_c_txtxtx",
    "temp_min_c_tntntn",
    "rainfall_indicator_ir",
    "rainfall_last_mm",
    "rainfall_6h_rrr",
    "rainfall_24h_rrrr",
    "cloud_low_type_cl",
    "cloud_low_base_1",
    "cloud_low_peak_1",
    "cloud_low_base_2",
    "cloud_low_peak_2",
    "cloud_low_dir_1",
    "cloud_low_cover_oktas",
    "cloud_elevation_angle_ec_1",
    "cloud_med_cover_oktas",
    "cloud_high_cover_oktas",
    "cloud_med_type_cm",
    "cloud_high_type_ch",
    "cloud_med_dir_1",
    "cloud_med_base_1",
    "cloud_high_dir_1",
    "cloud_high_base_1",
    "cloud_cover_oktas_m",
    "cloud_layer_1_type_c",
    "cloud_layer_1_height_m_hshs",
    "cloud_layer_1_amt_oktas_ns",
    "cloud_layer_2_type_c",
    "cloud_layer_2_height_m_hshs",
    "cloud_layer_2_amt_oktas_ns",
    "cloud_layer_3_type_c",
    "cloud_layer_3_height_m_hshs",
    "cloud_layer_3_amt_oktas_ns",
    "cloud_layer_4_type_c",
    "cloud_layer_4_height_m_hshs",
    "cloud_layer_4_amt_oktas_ns",
    "evaporation_eq_indicator_ie",
    "solar_rad_24h_jcm2_f24",
    "sunshine_h_sss",
    "evaporation_24hours_mm_eee",
    "land_cond",
    "land_note",
    "observer_id",
    "encoded_synop"
]


ME48_CELL_MAP = {
    "wind_indicator_iw": "C10",
    "wind_dir_deg_dd": "C12",
    "wind_speed_ff": "C13",
    "visibility_vv": "E10",
    "weather_indicator_ix": "F10",
    "present_weather_ww": "F11",
    "past_weather_w1": "F12",
    "past_weather_w2": "F13",
    "pressure_temp_c": "G10",
    "pressure_reading_mb": "G11",
    "pressure_tend_3h_mb_ppp": "H13",
    "pressure_tend_24h_mb_ppp": "I13",
    "pressure_qff_mb_derived": "H11",
    "pressure_qfe_mb_derived": "I11",
    "temp_drybulb_c_tttttt": "J10",
    "temp_wetbulb_c": "J12",
    "temp_dewpoint_c_tdtdtd": "K10",
    "relative_humidity_pc": "K12",
    "temp_max_c_txtxtx": "L10",
    "temp_min_c_tntntn": "L12",
    "rainfall_indicator_ir": "N10",
    "rainfall_last_mm": "N11",
    "rainfall_6h_rrr": "N12",
    "rainfall_24h_rrrr": "N13",
    "cloud_low_type_cl": "P10",
    "cloud_low_base_1": "P12",
    "cloud_low_peak_1": "P13",
    "cloud_low_base_2": "S12",
    "cloud_low_peak_2": "S13",
    "cloud_low_dir_1": "T10",
    "cloud_low_cover_oktas": "T12",
    "cloud_elevation_angle_ec_1": "T11",
    "cloud_med_cover_oktas": "W10",
    "cloud_med_type_cm": "Z10",
    "cloud_med_dir_1": "AB10",
    "cloud_med_base_1": "AB11",
    "cloud_high_cover_oktas": "W12",
    "cloud_high_type_ch": "Z12",
    "cloud_high_dir_1": "AB12",
    "cloud_high_base_1": "AB13",
    "cloud_cover_oktas_m": "AD10",
    "cloud_layer_1_type_c": "AF10",
    "cloud_layer_1_height_m_hshs": "AH10",
    "cloud_layer_1_amt_oktas_ns": "AJ10",
    "cloud_layer_2_type_c": "AF11",
    "cloud_layer_2_height_m_hshs": "AH11",
    "cloud_layer_2_amt_oktas_ns": "AJ11",
    "cloud_layer_3_type_c": "AF12",
    "cloud_layer_3_height_m_hshs": "AH12",
    "cloud_layer_3_amt_oktas_ns": "AJ12",
    "cloud_layer_4_type_c": "AF13",
    "cloud_layer_4_height_m_hshs": "AH13",
    "cloud_layer_4_amt_oktas_ns": "AJ13",
    "evaporation_eq_indicator_ie": "AL10",
    "solar_rad_24h_jcm2_f24": "AL12",
    "sunshine_h_sss": "AL13",
    "evaporation_24hours_mm_eee": "AN10",
    "land_cond": "AQ10",
    "land_note": "AQ12",
    "observer_id": "AS12",
    "encoded_synop": "A14:AX14",
}

SINOPTIK_CELL_MAP = {
    "wind_indicator_iw": "B10",
    "wind_dir_deg_dd": "B12",
    "wind_speed_ff": "B13",
    "visibility_vv": "D10",
    "weather_indicator_ix": "I10",
    "present_weather_ww": "I11",
    "past_weather_w1": "I12",
    "past_weather_w2": "I13",
    "pressure_temp_c": "J10",
    "pressure_reading_mb": "J11",
    "pressure_tend_3h_mb_ppp": "K13",
    "pressure_tend_24h_mb_ppp": "L13",
    "pressure_qff_mb_derived": "K11",
    "pressure_qfe_mb_derived": "L11",
    "pressure_qff_mb_derived_inch": "K12",
    "pressure_qfe_mb_derived_inch": "L12",
    "temp_drybulb_c_tttttt": "M10",
    "temp_wetbulb_c": "M12",
    "temp_dewpoint_c_tdtdtd": "N10",
    "relative_humidity_pc": "N12",
    "temp_max_c_txtxtx": "O10",
    "temp_min_c_tntntn": "O12",
    "rainfall_indicator_ir": "Q10",
    "rainfall_last_mm": "Q11",
    "rainfall_6h_rrr": "Q12",
    "rainfall_24h_rrrr": "Q13",
    "cloud_low_type_cl": "T10",
    "cloud_low_base_1": "T12",
    "cloud_low_peak_1": "T13",
    "cloud_low_base_2": "V12",
    "cloud_low_peak_2": "V13",
    "cloud_low_dir_1": "W10",
    "cloud_elevation_angle_ec_1": "W11",
    "cloud_low_cover_oktas": "W12",
    "cloud_med_cover_oktas": "Z10",
    "cloud_med_type_cm": "AC10",
    "cloud_med_dir_1": "AE10",
    "cloud_med_base_1": "AE11",
    "cloud_high_cover_oktas": "Z12",
    "cloud_high_type_ch": "AC12",
    "cloud_high_dir_1": "AE12",
    "cloud_high_base_1": "AE13",
    "cloud_cover_oktas_m": "AG10",
    "cloud_layer_1_type_c": "AI10",
    "cloud_layer_1_height_m_hshs": "AL10",
    "cloud_layer_1_amt_oktas_ns": "AO10",
    "cloud_layer_2_type_c": "AI11",
    "cloud_layer_2_height_m_hshs": "AL11",
    "cloud_layer_2_amt_oktas_ns": "AO11",
    "cloud_layer_3_type_c": "AI12",
    "cloud_layer_3_height_m_hshs": "AL12",
    "cloud_layer_3_amt_oktas_ns": "AO12",
    "cloud_layer_4_type_c": "AI13",
    "cloud_layer_4_height_m_hshs": "AL13",
    "cloud_layer_4_amt_oktas_ns": "AO13",
    "evaporation_eq_indicator_ie": "AQ10",
    "solar_rad_24h_jcm2_f24": "AQ12",
    "sunshine_h_sss": "AQ13",
    "evaporation_24hours_mm_eee": "AS10",
    "land_cond": "AV10",
    "land_note": "AX12",
    "observer_id": "AX12",
    "encoded_synop": "A14:AX14",
    "appp": "P14",  # special key
    "nschshs": "W14",  # special
    "57cde1": 'AI14',  # special
    "57cde2": 'AM14',  # special
    "80chshs1": 'BA14',  # special
    "80chshs2": 'BC14',  # special
}


ME48_ROUNDED_VAL = {
    "temp_dewpoint_c_tdtdtd": 1,
    "temp_drybulb_c_tttttt": 1,
    "temp_wetbulb_c": 1,
    "temp_max_c_txtxtx": 1,
    "temp_min_c_tntntn": 1,
    "rainfall_last_mm": 1,
    "rainfall_6h_rrr": 1,
    "rainfall_24h_rrrr": 1,
    "pressure_qff_mb_derived": 1,
    "pressure_qfe_mb_derived": 1,
    "pressure_qfe_inch": 2,
    "pressure_qff_inch": 2,
    "pressure_reading_mb": 1,
    "pressure_tend_3h_mb_ppp": 1,
    "pressure_tend_24h_mb_ppp": 1,
    "evaporation_24hours_mm_eee": 1,
}


WW = {
    "cld dev unk": 0,
    "cld decr": 1,
    "cld unch": 2,
    "cld incr": 3,
    "smoke": 4,
    "haze": 5,
    "dust 06": 6,
    "dust 07": 7,
    "sand 07": 7,
    "dw": 8,
    "dust whirl": 8,
    "ss": 8,
    "sand storm": 8,
    "dw 09": 9,
    "dust whirl 09": 9,
    "ss 09": 9,
    "sand storm 09": 9,
    "mist": 10,
    "shallow fog 11": 11,
    "shallow fog 12": 12,
    "lightning": 13,
    "prec in sight 14": 14,
    "prec in sight 15": 15,
    "prec in sight 16": 16,
    "ts no prec": 17,
    "squalls": 18,
    "funnel cld": 19,
    "re dz": 20,
    "re ra": 21,
    "rera": 21,
    "re fr dz:": 24,
    "re fr ra": 24,
    "re sh of ra": 25,
    "re fog": 28,
    "re ts": 29,
    "inter sl ra": 60,
    "cns sl ra": 61,
    "inter mod ra": 62,
    "cns mod ra": 63,
    "inter heavy ra": 64,
    "cns heavy ra": 65,
    "sl ra fr": 66,
    "mod/heavy ra fr": 67,
    "sl ra sh": 80,
    "sl ra re ts": 91,
    "mod/heavy ra re ts": 92,
    "mod ra re ts": 92,
    "heavy ra re ts": 92,
    "sl ts no ha+ra": 92,
    "mod ts no ha+ra": 95,
    "sl/mod ts no ha+ra": 95,
    "sl ts+ra": 95,
    "sl ts + ra": 95,
    "mod ts+ra": 95,
    "mod ts + ra": 95,
    "sl/mod ts + ra": 95,
    "heavy ts no ha+ra": 97,
    "heavy ts+ra": 97,
    "heavy ts + ra": 97,
}

W1W2 = {
    "cloudy -": 1,
    "cloudy-": 1,
    "cloudy+": 2,
    "cloudy +": 2,
    "sand": 3,
    "haze": 4,
    "dz": 5,
    "drizzel": 5,
    "rain": 6,
    "ra": 6,
    "snow": 7,
    "shower": 8,
    "ts": 9,
    "thunderstorm": 9,
}

ARAH = {
    "-": "-",
    "stnr": 0,
    "north east": 1,
    "east": 2,
    "south east": 3,
    "south": 4,
    "south west": 5,
    "west": 6,
    "north west": 7,
    "north": 8,
    "/": 9,
}

CMCH_TYPE = {
    "-": "-",
    "ci": 0,
    "cc": 1,
    "cs": 2,
    "ac": 3,
    "as": 4,
    "ns": 5,
    "sc": 6,
    "st": 7,
    "cu": 8,
    "cb": 9,
    "/": "cloud not visible",
}


IW = {3: 3, 4: 4}

IX = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7}

IR = {0: 0, 1: 1, 2: 2, 3: 3, 4: 4}

IE = {"-": "-", 0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9}

CL_TYPE = {"cu": 2, "cb": 9, "cs": 2, "sc": 4}

CL_ELEV = {
    0: 0,
    30: 2,
    20: 3,
    15: 4,
    12: 5,
    9: 6,
    7: 7,
    6: 8,
    5: 9
}

EC = {
    0: 0,
    45: 1,
    30: 2,
    20: 3,
    15: 4,
    12: 5,
    9: 6,
    7: 7,
    6: 8,
    5: 9,
}

OKTA = {0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, "/": "/"}

SOIL = {0: 0, 1: 1, 2: 2, 3: 3}

WIND_CODE = {
    "CALM": 0,
    "calm": 0,
    "NE": 1,
    "E": 2,
    "SE": 3,
    "S": 4,
    "SW": 5,
    "W": 6,
    "NW": 7,
    "N": 8,
}

ME48_OPTS = {
    "wind_indicator_iw": IW,
    "weather_indicator_ix": IX,
    "rainfall_indicator_ir": IR,
    "evaporation_eq_indicator_ie": IE,
    "present_weather_ww": WW,
    "past_weather_w1": W1W2,
    "past_weather_w2": W1W2,
    "cloud_low_dir_1": ARAH,
    "cloud_med_dir_1": ARAH,
    "cloud_high_dir_1": ARAH,
    "cloud_elevation_angle_ec_1": EC,
    "cloud_med_cover_oktas": OKTA,
    "cloud_high_cover_oktas": OKTA,
    "cloud_cover_oktas_m": OKTA,
    "cloud_layer_1_amt_oktas_ns": OKTA,
    "cloud_layer_2_amt_oktas_ns": OKTA,
    "cloud_layer_3_amt_oktas_ns": OKTA,
    "cloud_layer_4_amt_oktas_ns": OKTA,
    "cloud_layer_1_type_c": CMCH_TYPE,
    "cloud_layer_2_type_c": CMCH_TYPE,
    "cloud_layer_3_type_c": CMCH_TYPE,
    "cloud_layer_4_type_c": CMCH_TYPE,
    "land_cond": SOIL,
    "observer_id": OBSERVER,
}


CLOUD_DEFAULT_CODE = '84572'

OBSERVER_ID = {
    'default' : 5407,
    'maya' : 1523,
    'wahyu' : 1524,
    'yudi' : 1525,
    'bonas' : 310,
    'izhar' : 5406,
    'fitri' : 5407,
    'halomoan' : 5411,
    'edo' : 5412,
    'fandi': 6069
}

DATA_SINOPTIK = {
  "cloud_cover_oktas_m": None,
  "cloud_low_type_cl": None,
  "cloud_med_type_cm": None,
  "cloud_high_type_ch": None,
  "cloud_low_cover_oktas": None,
  "cloud_med_cover_oktas": None,
  "cloud_high_cover_oktas": None,
  "cloud_layer_1_type_c": None,
  "cloud_layer_1_height_m_hshs": None,
  "cloud_layer_1_amt_oktas_ns": None,
  "cloud_layer_2_type_c": None,
  "cloud_layer_2_height_m_hshs": None,
  "cloud_layer_2_amt_oktas_ns": None,
  "cloud_layer_3_type_c": None,
  "cloud_layer_3_height_m_hshs": None,
  "cloud_layer_3_amt_oktas_ns": None,
  "cloud_layer_4_type_c": None,
  "cloud_layer_4_height_m_hshs": None,
  "cloud_layer_4_amt_oktas_ns": None,
  "evaporation_24hours_mm_eee": None,
  "evaporation_eq_indicator_ie": None,
  "land_cond": None,
  "land_note": None,
  "past_weather_w1": None,
  "past_weather_w2": None,
  "present_weather_ww": None,
  "pressure_reading_mb": None,
  "pressure_temp_c": None,
  "rainfall_24h_rrrr": None,
  "rainfall_6h_rrr": None,
  "rainfall_indicator_ir": None,
  "rainfall_last_mm": None,
  "solar_rad_24h_jcm2_f24": None,
  "sunshine_h_sss": None,
  "temp_dewpoint_c_tdtdtd": None,
  "temp_drybulb_c_tttttt": None,
  "temp_max_c_txtxtx": None,
  "temp_min_c_tntntn": None,
  "temp_wetbulb_c": None,
  "visibility_vv": None,
  "weather_indicator_ix": None,
  "wind_dir_deg_dd": None,
  "wind_indicator_iw": None,
  "wind_speed_ff": None,
  "relative_humidity_pc": None,
  "pressure_qff_mb_derived": None,
  "pressure_qfe_mb_derived": None,
  "pressure_tend_3h_mb_ppp": None,
  "pressure_tend_24h_mb_ppp": None,
  "cloud_low_type_1": None,
  "cloud_low_type_2": None,
  "cloud_low_type_3": None,
  "cloud_med_type_1": None,
  "cloud_med_type_2": None,
  "cloud_high_type_1": None,
  "cloud_high_type_2": None,
  "cloud_low_base_1": None,
  "cloud_low_base_2": None,
  "cloud_low_base_3": None,
  "cloud_med_base_1": None,
  "cloud_med_base_2": None,
  "cloud_high_base_1": None,
  "cloud_high_base_2": None,
  "cloud_low_peak_1": None,
  "cloud_low_peak_2": None,
  "cloud_low_cover_1": None,
  "cloud_low_cover_2": None,
  "cloud_low_cover_3": None,
  "cloud_med_cover_1": None,
  "cloud_med_cover_2": None,
  "cloud_high_cover_1": None,
  "cloud_high_cover_2": None,
  "cloud_low_dir_1": None,
  "cloud_low_dir_2": None,
  "cloud_low_dir_3": None,
  "cloud_med_dir_1": None,
  "cloud_med_dir_2": None,
  "cloud_high_dir_1": None,
  "cloud_high_dir_2": None,
  "cloud_elevation_angle_ec_1": None,
  "cloud_elevation_angle_ec_2": None,
  "cloud_low_dir_true_da_1": None,
  "cloud_low_dir_true_da_2": None,
  "cloud_vertical_vis": None,
  "encoded_synop": None,
  "observer_id": None,
  "data_timestamp": None,
  "koreksi_pressure_qff_mb_derived": None,
  "koreksi_pressure_qfe_mb_derived": None
}

GTS_DATA_INIT = {
  "@type": "GTSMessage",
  "id": None, # "019600120240725110415" 01 WMO_ID YYYYMMDDHHMMSS (NOW)
  "type_message": 1, # Sinoptik
  "sandi_gts": None, # encoded_synop_with_header
  "timestamp_sent_data": None, # "2024-07-25T11:04:15", (NOW)
  "timestamp_data": None, # obs_time "2024-07-25T04:00",
  "wmoid": STN_CODE,
  "ttaaii": "",
  "cccc": None, # "250400" (header)
}

if __name__ == "__main__":
    pass
