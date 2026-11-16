from fastapi import BackgroundTasks, FastAPI, Depends, HTTPException
from fastapi.responses import StreamingResponse
from typing import Optional, Union
from fastapi.security import OAuth2PasswordBearer
from pydantic import BaseModel
from utils import *
import asyncio
import uvicorn
from uvicorn.config import LOGGING_CONFIG


app = FastAPI()
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")


class SinoptikInputData(BaseModel):
    data_timestamp: str
    created_by: Optional[str] = None
    wind_dir_deg_dd: Optional[Union[int, str]] = None
    wind_speed_ff: Optional[float] = None
    pressure_qfe_mb_derived: Optional[float] = None
    temp_drybulb_c_tttttt: Optional[float] = None
    temp_dewpoint_c_tdtdtd: Optional[float] = None
    temp_wetbulb_c: Optional[float] = None
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

        # Ambil output stdout & stderr
        stdout_text = stdout.decode().strip() if stdout else ""
        stderr_text = stderr.decode().strip() if stderr else ""

        # Cetak output ke console
        if stdout_text:
            print_with_timestamp(f"\n{stdout_text}")

        if process.returncode != 0:
            last_line = stderr_text.splitlines()[-1] if stderr_text else "Terjadi kesalahan saat mengambil data sinoptik."
            raise HTTPException(status_code=500, detail=last_line)

        return {"status": "success", "message": f"Data sinoptik {timestamp} berhasil diambil."}
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{str(e)}")


@app.post("/sinoptik/input")
def post_data(data: SinoptikInputData, bg_task: BackgroundTasks, token: str = Depends(verify_token)):
    try:
        # Mengonversi Pydantic model menjadi dictionary
        data_dict = data.model_dump()

        # Sebelum 1 Juni 2025 perhitungan angin dan suhu bola basah masih menggunakan fungsi
        # normalize_data dan calculate_tbb di utils.py
        # Sekarang, perhitungan angin dan suhu bola basah dilakukan di server bmkgsabang

        # Jika data angin AWS dan AWOS tidak ada
        if data_dict['wind_dir_deg_dd'] == -9999:
            data_dict['wind_dir_deg_dd'] = get_previous_wind_direction(data_dict['data_timestamp'])

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
        bg_task.add_task(write_excel, data_dict, callback=on_write_complete)
      
        return {"status": "success", "message": f"Data sinoptik {data_dict['data_timestamp']} berhasil diinput ke excel."}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    

@app.get("/sinoptik/screenshot")
async def capture_and_upload_screenshot(token: str = Depends(verify_token), bg_task: BackgroundTasks = None):
    """Menjalankan proses pengambilan screenshot dan upload dalam background."""
    try:
        take_screenshot()
        bg_task.add_task(upload_image)

        return {"status": "success", "message": "Screenshot sedang diproses dan diunggah."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/sinoptik/restart-excel")
def restart_excel(token: str = Depends(verify_token)):
    result = kill_excel_processes_and_reload()
    return {
        "status": "success" if result == 0 else "error",
        "message": "Program excel sinoptik sudah direstart."
    }


@app.get("/app/cctv")
def webcam_stream():
    return StreamingResponse(
        generate_video(),
        media_type='multipart/x-mixed-replace; boundary=frame',
        headers={
            "Cache-Control": "no-store",
            "Pragma": "no-cache",
            "Expires": "0"
        }
    )


@app.get("/app/cctv-audio")
def stream_audio():
    return StreamingResponse(generate_audio(), media_type="audio/wav")


def run_server():
    LOGGING_CONFIG["formatters"]["default"]["fmt"] = "%(asctime)s\t%(message)s"
    LOGGING_CONFIG["formatters"]["default"]["datefmt"] = "%Y-%m-%d %H:%M:%S"

    LOGGING_CONFIG["formatters"]["access"]["fmt"] = '%(asctime)s\t%(client_addr)s - "%(request_line)s" %(status_code)s'
    LOGGING_CONFIG["formatters"]["access"]["datefmt"] = "%Y-%m-%d %H:%M:%S"

    uvicorn.run(app, host="0.0.0.0", port=8080, log_config=LOGGING_CONFIG, access_log=DEBUG)


if __name__ == "__main__":
    try:
        start_caddy()
        run_server()
    finally:
        stop_caddy()
