from fastapi import BackgroundTasks, FastAPI, Depends, HTTPException
from typing import Optional, Union, Callable
from fastapi.security import OAuth2PasswordBearer
from pydantic import BaseModel
from utils import *
import uvicorn
import asyncio


app = FastAPI()
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")


class SinoptikInputData(BaseModel):
    data_timestamp: str
    created_by: Optional[str] = None
    wind_dir_deg_dd: Optional[str] = None
    wind_speed_ff: Optional[float] = None
    pressure_qfe_mb_derived: Optional[float] = None
    temp_drybulb_c_tttttt: Optional[float] = None
    temp_dewpoint_c_tdtdtd: Optional[float] = None
    relative_humidity_pc: Optional[int] = None
    rainfall_last_mm: Optional[float] = None
    rainfall_6h_rrr: Optional[float] = None
    rainfall_24h_rrrr: Optional[float] = None
    pressure_qff_mb_derived: Optional[float] = None
    temp_min_c_tntntn: Optional[Union[float, str]] = None
    temp_max_c_txtxtx: Optional[Union[float, str]] = None
    weather_indicator_ix: Optional[int] = None
    rainfall_indicator_ir: Optional[int] = None
    present_weather_ww: Optional[str] = None
    past_weather_w1: Optional[str] = None
    past_weather_w2: Optional[str] = None


def verify_token(token: str = Depends(oauth2_scheme)):
    if token != SINOPTIK_TOKEN:
        raise HTTPException(status_code=401, detail="Invalid authentication credentials")
    return token


def read_excel(timestamp):
    "Read data 3 hours before, date based on timestamp"
    synop_file = get_synop_path(PATH_TO_SYNOP, timestamp.year, timestamp.month, timestamp.day)
    sheet_name = get_sheet_name(timestamp.hour)
    cell_map = generate_new_sinoptik_cell_map(timestamp.hour, SINOPTIK_CELL_MAP)
    # ganti nilai cell perawanan
    cell_map['nschshs'] = get_cloud_cell(timestamp.hour)
    cell_map['57cde1'] = get_convective_cloud_cell(timestamp.hour)[0]
    cell_map['57cde2'] = get_convective_cloud_cell(timestamp.hour)[1]
    try:
        wb = open_workbook(synop_file)
        data = extract_values(wb, sheet_name, cell_map)
    except Exception as e:
        print_with_timestamp(e)
    
    return data
    

def write_excel(data: dict, callback: Callable = None):
    print_with_timestamp(f'Mengisi data sinoptik {data["data_timestamp"]} ke excel...')
    timestamp = arrow.get(data["data_timestamp"])
    data_before = {}  # data dari 3 jam sebelumnya
    
    try:
        year = timestamp.year
        month = timestamp.month
        day = timestamp.day
        hour = timestamp.hour        
        synop_file = get_synop_path(PATH_TO_SYNOP, year, month, day)
        sheet_name = get_sheet_name(hour)
        cell_map = generate_new_sinoptik_cell_map(hour, SINOPTIK_CELL_MAP)

        # ganti nilai cell perawanan
        cell_map["nschshs"] = get_cloud_cell(hour)
        cell_map['57cde1'] = get_convective_cloud_cell(hour)[0]
        cell_map['57cde2'] = get_convective_cloud_cell(hour)[1]
        cell_map['80chshs1'] = get_cl_peak_cell(hour)[0]
        cell_map['80chshs2'] = get_cl_peak_cell(hour)[1]

        # data 3 hours before
        for h in range(3, 0, -1):
            dt_before = timestamp.shift(hours=-h)
            prev_data = read_excel(dt_before)
            data_before[h] = prev_data  # prev_data[1] = data 1 jam sebelumnya

        # data 24 hours before, just extract qfe
        if hour in (0, 12):
            dt24_before = timestamp.shift(hours=-24)
            prev24_data = read_excel(dt24_before)
    except Exception as e:
        print_with_timestamp(e)

    try:

        wb = open_workbook(synop_file)
        ws = wb.sheets[sheet_name]
        try:
            if hour in (0, 6, 12, 18):
                ws.select()
        except Exception as e:
            print_with_timestamp(e)

        for key in ('wind_dir_deg_dd', 'wind_speed_ff', 'pressure_qfe_mb_derived', 'pressure_qff_mb_derived', 
                    'temp_drybulb_c_tttttt', 'temp_wetbulb_c'):
            # don't override
            if not ws.range(cell_map[key]).value:
                ws.range(cell_map['wind_dir_deg_dd']).value = data["wind_dir_deg_dd"]
                ws.range(cell_map['wind_speed_ff']).value = data["wind_speed_ff"]
                ws.range(cell_map['pressure_qfe_mb_derived']).value = data["pressure_qfe_mb_derived"]
                ws.range(cell_map['pressure_qff_mb_derived']).value = data["pressure_qff_mb_derived"]
                ws.range(cell_map['temp_drybulb_c_tttttt']).value = data["temp_drybulb_c_tttttt"]
                ws.range(cell_map['temp_wetbulb_c']).value = data["temp_wetbulb_c"]

        try:
            if hour in (0, 12):
                if not ws.range(cell_map['pressure_tend_24h_mb_ppp']).value and prev24_data['pressure_qfe_mb_derived']:
                    ws.range(cell_map['pressure_tend_24h_mb_ppp']).value = round(data['pressure_qfe_mb_derived'] - prev24_data['pressure_qfe_mb_derived'], 1)
                if hour == 12:
                    if not ws.range(cell_map['temp_max_c_txtxtx']).value:
                        # bug garing
                        if data['tx']:
                            if '/' not in str(data['temp_max_c_txtxtx']):
                                ws.range(cell_map['temp_max_c_txtxtx']).value = data["temp_max_c_txtxtx"]
                elif hour == 0:
                    if not ws.range(cell_map['temp_min_c_tntntn']).value:
                        if data['temp_min_c_tntntn']:
                            if '/' not in str(data['temp_min_c_tntntn']):
                                ws.range(cell_map['temp_min_c_tntntn']).value = data["temp_min_c_tntntn"]
        except Exception as e:
            print_with_timestamp(e)

        try:
            if data_before[1]:
                for key in ('visibility_vv', 'weather_indicator_ix', 'present_weather_ww', 'past_weather_w1', 
                            'past_weather_w2', 'rainfall_indicator_ir', 'cloud_low_type_cl', 'cloud_low_base_1', 
                            'cloud_low_peak_1', 'cloud_low_dir_1', 'cloud_elevation_angle_ec_1', 'cloud_low_cover_oktas', 
                            'cloud_med_cover_oktas', 'cloud_high_cover_oktas', 'cloud_med_type_cm', 'cloud_high_type_ch', 
                            'cloud_med_dir_1', 'cloud_med_base_1', 'cloud_high_dir_1', 'cloud_high_base_1',
                            'cloud_cover_oktas_m', 'cloud_layer_1_type_c', 'cloud_layer_1_height_m_hshs', 'cloud_layer_1_amt_oktas_ns',
                            'cloud_layer_2_type_c', 'cloud_layer_2_height_m_hshs', 'cloud_layer_2_amt_oktas_ns', 'land_cond', 'encoded_synop'):
                    if not ws.range(cell_map[key]).value:
                        if key == 'weather_indicator_ix':
                            ws.range(cell_map[key]).value = 2
                        elif key == 'present_weather_ww':
                            ws.range(cell_map[key]).value = 'Cld Unch'
                        elif key in ('past_weather_w1', 'past_weather_w2'):
                            ws.range(cell_map[key]).value = 'Cloudy +'
                        elif key == 'rainfall_indicator_ir':
                            ws.range(cell_map[key]).value = 3
                        else:
                            ws.range(cell_map[key]).value = data_before[1][key]

                    # awan konvektif
                    try:
                        if key == 'cloud_low_base_1':
                            if ws.range(cell_map['cloud_low_type_cl']).value:
                                cl_val = ws.range(cell_map['cloud_low_type_cl']).value.lower()
                                dl = str(ws.range(cell_map['cloud_low_dir_1']).value).lower()
                                elev = ws.range(cell_map['cloud_elevation_angle_ec_1']).value
                                if cl_val in ('cu', 'cu/sc'):
                                    if dl and elev:
                                        ws.range(cell_map['cloud_low_base_1']).value = '8' + str(ARAH[dl]) + str(CL_ELEV[int(elev)])
                                elif cl_val in ('cb', 'cb/sc'):
                                    if dl and elev:
                                        ws.range(cell_map['cloud_low_base_1']).value = '9' + str(ARAH[dl]) + str(CL_ELEV[int(elev)])
                                elif cl_val in ('cb/cu'):
                                    if dl and elev:
                                        ws.range(cell_map['cloud_low_base_1']).value = '9' + str(ARAH[dl]) + str(CL_ELEV[int(elev)])
                                        ws.range(cell_map['cloud_low_peak_1']).value = '8' + str(ARAH[dl]) + str(CL_ELEV[int(elev)])

                        elif key == 'cloud_high_base_1':
                            if ws.range(cell_map['cloud_low_type_cl']).value:
                                cl_val = ws.range(cell_map['cloud_low_type_cl']).value.lower()
                                if ws.range(cell_map['cloud_low_peak_1']).value:
                                    hshs = int(ws.range(cell_map['cloud_low_peak_1']).value)
                                    cl_peak = calculate_hshs(hshs)
                                if cl_val == 'cu':
                                    if cl_peak:
                                        ws.range(cell_map['cloud_high_base_1']).value = '8' + str(cl_peak)
                                elif cl_val == 'cb':
                                    if cl_peak:
                                        ws.range(cell_map['cloud_high_base_1']).value = '9' + str(cl_peak)
                                elif cl_val in ('cb/cu'):
                                    hshs2 = int(ws.range(cell_map['cloud_layer_2_amt_oktas_ns']).value)
                                    cl_peak2 = calculate_hshs(hshs2)
                                    if cl_peak:
                                        ws.range(cell_map['cloud_high_base_1']).value = '8' + str(cl_peak)
                                    if cl_peak2:
                                        ws.range(cell_map['cloud_high_base_1']).value = '9' + str(cl_peak2)

                    except Exception as e:
                        print_with_timestamp(e)

            # isi data tendency per 3 jam
            if hour % 3 == 0:
                plist = [data_before[i]['pressure_qff_mb_derived'] for i in range(3, 0, -1)] + [data['pressure_qff_mb_derived']]
                tendency = calculate_tendency(plist)
                ws.range(cell_map['appp']).value = tendency
                if hour == 0:
                    if not ws.range(cell_map['appp']).value:
                        ws.range(cell_map['appp']).value = round(plist[3] - plist[0], 1)

        except Exception as e:
            print_with_timestamp(e)
                
        wb.save()

        # Call the callback on complete
        if callback:
            callback()

    except Exception as e:
        print_with_timestamp(e)

    
@app.get("/sinoptik/fetch")
async def fetch_sinoptik_data(timestamp: str, bg_task: BackgroundTasks, token: str = Depends(verify_token), user_id: Optional[str] = None):
    try:
        # Menjalankan subprocess dengan stdout dan stderr yang bisa dibaca
        process = await asyncio.create_subprocess_exec(
            'python', 'autosinoptik.py', timestamp,
            stdout=asyncio.subprocess.PIPE,  # Tangkap output stdout
            stderr=asyncio.subprocess.PIPE   # Tangkap output stderr
        )

        # Membaca output secara real-time
        stdout, stderr = await process.communicate()

        # Cetak hasil output ke console
        if stdout:
            print_with_timestamp(f"\n{stdout.decode()}")
        if stderr:
            print_with_timestamp(f"\n{stderr.decode()}")

        # Periksa status proses
        if process.returncode == 0:
            return {"status": "success", "message": f"Data sinoptik {timestamp} berhasil diambil."}
        else:
            raise HTTPException(status_code=500, detail=f"Terjadi kesalahan saat mengambil data sinoptik.")
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{str(e)}")


@app.post("/sinoptik/input")
def post_data(data: SinoptikInputData, bg_task: BackgroundTasks, token: str = Depends(verify_token)):
    try:
        # Mengonversi Pydantic model menjadi dictionary
        data_dict = data.model_dump()

        # Menghitung tbb menggunakan fungsi calculate_tbb  
        tbb = calculate_tbb(data_dict['temp_drybulb_c_tttttt'], data_dict['relative_humidity_pc'], data_dict['temp_dewpoint_c_tdtdtd'])  
        data_dict.update({'temp_wetbulb_c': tbb})  

        # Memeriksa apakah arah angin mengandung simbol '/'
        if "/" in data_dict['wind_dir_deg_dd']:  
            data_dict['wind_dir_deg_dd'] = get_previous_wind_direction(data_dict['data_timestamp'])
            
        else:
            data_dict['wind_dir_deg_dd'] = int(data_dict['wind_dir_deg_dd'])  

        # Menormalisasi data di data_dict  
        normalized_data = normalize_data(data_dict)

        def on_write_complete():
            url = BMKGSABANG_URL + "/client/sinoptik/sinoptik-data-written-to-excel"
            params = {
                'token': SINOPTIK_TOKEN,
                'data_timestamp': data_dict['data_timestamp'],
                'user_id': data_dict['created_by']
            }
            try:
                requests.get(url, params=params, timeout=5, verify=False)
            except:
                pass
            print_with_timestamp(f"Data sinoptik {data_dict['data_timestamp']} berhasil diinput ke excel.")


        # Menambahkan tugas latar belakang untuk menulis ke Excel  
        bg_task.add_task(write_excel, normalized_data, callback=on_write_complete)
      
        return {"status": "success", "message": "Data sinoptik berhasil diinput ke excel."}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
