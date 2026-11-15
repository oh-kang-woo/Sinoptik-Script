import threading
import logging
import os
import sys
import requests
import mss
import urllib3
import signal
import argparse
import subprocess
import platform
from time import sleep
from datetime import datetime, timedelta

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = 'https://bmkgsabang.test/api/peralatan'

# token user dari bmkgsabang
MONITORING_TOKEN = '27|T6SD1DvLuRQPNpM0V7HUFSoChrAWpmW3pamqWyCz076b510a'

# id alat dari bmkgsabang
ID_PERALATAN = '9e33138f-27db-4ed9-b2b7-5acc37dfbb28'
PING_INTERVAL = 1  # minute
IMAGE_INTERVAL = 1  # minute
UTC_DIFF = 7  # hour
SCREENSHOT_FILENAME = 'D:\\autosinoptik\\image.jpg'

# use screenshot and time from external api
USE_EXTERNAL_SCREENSHOT = False
USE_EXTERNAL_TIME = False
TIME_DATA_KEY = 'waktu'
LATEST_DATA_URL = 'http://202.90.199.132/aws-new/data/station/latest/3000000038'
EXTERNAL_SCREENSHOT_URL = 'http://202.90.199.132/aws-new/monitoring/3000000038'

if platform.system() == 'Windows':
    chrome_path = '"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"'
elif platform.system() == 'Linux':
    chrome_path = 'google-chrome-stable'
else:
    raise EnvironmentError("Unsupported OS: " + platform.system())

EXTERNAL_SCREENSHOT_CMD = f'{chrome_path} --headless ' \
                          '--disable-gpu --window-size=1280,768 ' \
                          '--screenshot={} ' \
                          '--disable-logging --log-level=3 ' \
                          '--wait-for-js-execution ' \
                          '--virtual-time-budget=10000 ' \
                          '--run-all-queries ' \
                          '--disable-features=IsolateOrigins,site-per-process ' \
                          '{} > NUL 2>&1'.format(SCREENSHOT_FILENAME, EXTERNAL_SCREENSHOT_URL)


# exit flag (interrupt script with Ctrl+C)
exit_flag = False

# Logger setup
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)


def signal_handler(signum, frame):
    """Handle Ctrl+C"""
    global exit_flag
    logger.info('Interrupt received, exiting...')
    exit_flag = True

def get_new_dt(dt_string, hour_diff):
    """Menghitung waktu baru berdasarkan perbedaan UTC"""
    dt = datetime.strptime(dt_string, '%Y-%m-%d %H:%M:%S')
    dt = dt + timedelta(hours=hour_diff)
    return dt.strftime('%Y-%m-%d %H:%M:%S')

def fetch_time_from_external_api(key='waktu'):
    """Ambil data waktu dari API eksternal"""
    try:
        res = requests.get(LATEST_DATA_URL, verify=False)
        return res.json()[key]
    except:
        return None

def take_screenshot():
    """Ambil screenshot, baik menggunakan perintah eksternal atau module mss"""
    try:
        logger.debug('Taking screenshot')
        if USE_EXTERNAL_SCREENSHOT:
            # Simpan waktu modifikasi file sebelum diambil
            if os.path.exists(SCREENSHOT_FILENAME):
                last_modified_time_before = os.path.getmtime(SCREENSHOT_FILENAME)
            else:
                last_modified_time_before = 0  # Jika file belum ada sebelumnya

            subprocess.call(EXTERNAL_SCREENSHOT_CMD, shell=True)

            # Periksa waktu modifikasi file setelah screenshot diambil
            if os.path.exists(SCREENSHOT_FILENAME):
                last_modified_time_after = os.path.getmtime(SCREENSHOT_FILENAME)

                # Bandingkan waktu modifikasi sebelum dan sesudah
                if last_modified_time_after > last_modified_time_before:
                    logger.debug('Screenshot successfully taken')

        else:
            with mss.mss(display=':0') as sct:
                sct.shot(output=SCREENSHOT_FILENAME)
            logger.debug('Screenshot successfully taken')

    except Exception as e:
        logger.error('Failed to take screenshot: %s', e)


def send_ping_signal():
    """Kirim ping dengan retry maksimal 3 kali"""
    headers = {'Authorization': f'Bearer {MONITORING_TOKEN}', 'Accept': 'application/json'}
    url = f'{BASE_URL}/{ID_PERALATAN}'
    data = {}
    
    for attempt in range(3):  # Retry maksimal 3 kali
        try:
            logger.debug(f'Sending ping signal')
            if USE_EXTERNAL_TIME:
                data_waktu = fetch_time_from_external_api(TIME_DATA_KEY)
                if data_waktu:
                    data_waktu = get_new_dt(data_waktu, UTC_DIFF)
                data = {'last_update': data_waktu}
            response = requests.put(url, headers=headers, verify=False, data=data, timeout=10)
            response.raise_for_status()
            logger.debug('Ping signal successfully sent')
            return
        except requests.RequestException as e:
            logger.warning(f'Failed to send ping signal: {e}')
            sleep(2)  # Delay sebelum mencoba lagi

    logger.error('Ping signal failed after 3 attempts')


def upload_image():
    """Upload screenshot dengan retry maksimal 3 kali"""
    headers = {'Authorization': f'Bearer {MONITORING_TOKEN}'}
    url = f'{BASE_URL}/{ID_PERALATAN}?_method=PUT'

    if not os.path.exists(SCREENSHOT_FILENAME):
        return

    for attempt in range(3):  # Retry maksimal 3 kali
        try:
            logger.debug(f'Uploading image')
            with open(SCREENSHOT_FILENAME, 'rb') as f:
                files = {'image': f}
                response = requests.post(url, headers=headers, files=files, verify=False, timeout=15)
                response.raise_for_status()
                logger.debug('Image successfully uploaded')
                return
        except requests.RequestException as e:
            logger.warning(f'Failed to upload image: {e}')
            sleep(2)

    logger.error('Image upload failed after 3 attempts')


def take_screenshot_thread():
    """Thread untuk screenshot"""
    global exit_flag
    while not exit_flag:
        take_screenshot()
        for _ in range(55):
            if exit_flag:
                return
            sleep(1)


def ping_thread():
    """Thread untuk mengirim ping"""
    global exit_flag
    while not exit_flag:
        send_ping_signal()
        for _ in range(PING_INTERVAL * 60):
            if exit_flag:
                return
            sleep(1)


def upload_image_thread():
    """Thread untuk upload gambar"""
    global exit_flag
    while not exit_flag:
        upload_image()
        for _ in range(IMAGE_INTERVAL * 60):
            if exit_flag:
                return
            sleep(1)


def start_monitoring_task():
    """Jalankan semua thread"""
    global exit_flag
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    logger.info('Starting monitoring (Press Ctrl+C to stop)')

    threads = [
        threading.Thread(target=take_screenshot_thread, daemon=True),
        threading.Thread(target=ping_thread, daemon=True),
        threading.Thread(target=upload_image_thread, daemon=True)
    ]

    for thread in threads:
        thread.start()

    try:
        while not exit_flag:
            sleep(1)
    except KeyboardInterrupt:
        logger.info('KeyboardInterrupt received')
        exit_flag = True

    for thread in threads:
        thread.join()

    logger.info('Script stopped successfully')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Monitoring Script')
    parser.add_argument('--external-sshot', action='store_true', help='Enable external screenshot command')
    parser.add_argument('--external-time', action='store_true', help='Use external time data for last update')

    args = parser.parse_args()
    USE_EXTERNAL_SCREENSHOT = args.external_sshot
    USE_EXTERNAL_TIME = args.external_time

    if MONITORING_TOKEN is None:
        print('TOKEN tidak ditemukan')
        sys.exit()

    start_monitoring_task()
