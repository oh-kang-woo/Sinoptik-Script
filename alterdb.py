import sqlite3

def rename_columns(database_path, table_name, column_mapping):
    conn = sqlite3.connect(database_path)
    cursor = conn.cursor()

    for old_column, new_column in column_mapping.items():
        query = f"ALTER TABLE {table_name} RENAME COLUMN {old_column} TO {new_column};"
        print(f"Executing: {query}")
        try:
            cursor.execute(query)
        except:
            pass

    conn.commit()
    conn.close()

# Parameter
database_path = "database_96001.db"
table_name = "me48"
column_mapping = {
    "utc": "data_timestamp",
    "iw": "wind_indicator_iw",
    "dd": "wind_dir_deg_dd",
    "ff": "wind_speed_ff",
    "vv": "visibility_vv",
    "ix": "weather_indicator_ix",
    "ww": "present_weather_ww",
    "w1": "past_weather_w1",
    "w2": "past_weather_w2",
    "tp": "pressure_temp_c",
    "p": "pressure_reading_mb",
    "p3": "pressure_tend_3h_mb_ppp",
    "qff": "pressure_qff_mb_derived",
    "qfe": "pressure_qfe_mb_derived",
    "p24": "pressure_tend_24h_mb_ppp",
    "tbk": "temp_drybulb_c_tttttt",
    "tbb": "temp_wetbulb_c",
    "td": "temp_dewpoint_c_tdtdtd",
    "rh": "relative_humidity_pc",
    "tx": "temp_max_c_txtxtx",
    "tn": "temp_min_c_tntntn",
    "ir": "rainfall_indicator_ir",
    "rrr": "rainfall_last_mm",
    "r6": "rainfall_6h_rrr",
    "r24": "rainfall_24h_rrrr",
    "cl": "cloud_low_type_cl",
    "cl1_h": "cloud_low_base_1",
    "cl1_p": "cloud_low_peak_1",
    "cl2_h": "cloud_low_base_2",
    "cl2_p": "cloud_low_peak_2",
    "dl": "cloud_low_dir_1",
    "nh": "cloud_low_cover_oktas",
    "ec": "cloud_elevation_angle_ec_1",
    "cm_n": "cloud_med_cover_oktas",
    "ch_n": "cloud_high_cover_oktas",
    "cm": "cloud_med_type_cm",
    "ch": "cloud_high_type_ch",
    "dm": "cloud_med_dir_1",
    "cm_h": "cloud_med_base_1",
    "dh": "cloud_high_dir_1",
    "ch_h": "cloud_high_base_1",
    "n": "cloud_cover_oktas_m",
    "c1": "cloud_layer_1_type_c",
    "c1_h": "cloud_layer_1_height_m_hshs",
    "c1_n": "cloud_layer_1_amt_oktas_ns",
    "c2": "cloud_layer_2_type_c",
    "c2_h": "cloud_layer_2_height_m_hshs",
    "c2_n": "cloud_layer_2_amt_oktas_ns",
    "c3": "cloud_layer_3_type_c",
    "c3_h": "cloud_layer_3_height_m_hshs",
    "c3_n": "cloud_layer_3_amt_oktas_ns",
    "c4": "cloud_layer_4_type_c",
    "c4_h": "cloud_layer_4_height_m_hshs",
    "c4_n": "cloud_layer_4_amt_oktas_ns",
    "ie": "evaporation_eq_indicator_ie",
    "f24": "solar_rad_24h_jcm2_f24",
    "sss": "sunshine_h_sss",
    "eee": "evaporation_24hours_mm_eee",
    "e": "land_cond",
    "note": "land_note",
    "observer": "observer_id",
    "synop" : "encoded_synop"
}


rename_columns(database_path, table_name, column_mapping)
