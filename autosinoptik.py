import argparse
from utils import *


def process_sinoptik(timestamp, bmkgsatu_only=False):
    """Proses sinoptik data
    timestamp: bisa tunggal YYYYMMDDHH atau range dipisah tanda -
    bmkgsatu_only: hanya kirim ke bmkgsabang untuk disimpan, jangan teruskan ke bmkgsatu"""

    print_with_timestamp(f'Memproses data sinoptik {timestamp}...')

    data_dict = save_sinoptik_data(timestamp, bmkgsatu_only)

    # Loop jika timestamp yang diberikan adalah range
    for ts, data in data_dict.items():
        sinoptik_data = mutate_sinoptik_data_before_sent(data, bmkgsatu_only)

        print_with_timestamp(f'Memproses data sinoptik {ts} ...')
        if (bmkgsatu_only):
            break_here = wait_for_input('', timeout=5, default='Y')
            if break_here in ('C', 'c'):
                print_with_timestamp(f'Data sinoptik {ts} telah disimpan ke database.')
                continue  # Lewati pengiriman dan lanjutkan ke timestamp berikutnya

        send_sinoptik_data(sinoptik_data, bmkgsatu_only)


def send_sinoptik_data(sinoptik_data, bmkgsatu_only=False):
    """Send sinoptik data to bmkgsabang and/or bmkgsatu
    sinoptik_data: dictionary data sinoptik
    bmkgsatu_only: jika True, langsung mengirim ke bmkgsatu tanpa melewati bmkgsabang
    """

    if bmkgsatu_only:
        sinoptik_post_url = BMKGSATU_URL + '/db/bmkgsatu//1/1/20/1/meteorologi/sinoptik'
        sinoptik_gts_post_url = BMKGSATU_URL + '/db/bmkgsatu//1/1/20/1/gts'
    else:
        sinoptik_post_url = BMKGSABANG_URL + '/api/data-sinoptik'
    
    hour = arrow.get(sinoptik_data['data_timestamp']).hour
    if hour in SYNOP_HOURS:
        try:
            if bmkgsatu_only:
                send_confirmation(sinoptik_data['encoded_synop_with_header'])
                response = send_to_bmkgsatu(url=sinoptik_post_url, data=sinoptik_data)

                if response.status_code not in [requests.codes.ok, requests.codes.created]:
                    print_with_timestamp(f'Gagal mengirim data sinoptik jam {hour:02} ke bmkgsatu. Status: {response.status_code}, Response: {response.text}')

                print_with_timestamp(f'Berhasil mengirim data sinoptik jam {hour:02} ke bmkgsatu')

                gts_data = create_gts_data(sinoptik_data)
                gts_response = send_to_bmkgsatu(url=sinoptik_gts_post_url, data=gts_data)

                if gts_response.status_code not in [requests.codes.ok, requests.codes.created]:
                    print_with_timestamp(f'Gagal mengirim data sinoptik jam {hour:02} ke bmkgswitching. Status: {gts_response.status_code}, Response: {gts_response.text}')

                print_with_timestamp(f'Berhasil mengirim data sinoptik jam {hour:02} ke bmkgswitching')
                
                return {'status': 'success', 'message': f'Data sinoptik jam {hour:02} berhasil dikirim ke BMKGsatu dan BMKGswitching'}
            
            else:
                response = send_to_bmkgsabang(url=sinoptik_post_url, data=sinoptik_data)

                if response.status_code not in [requests.codes.ok, requests.codes.created]:
                    print_with_timestamp(f'Gagal mengirim data sinoptik jam {hour:02} ke bmkgsabang. Status: {response.status_code}, Response: {response.text}')

                print_with_timestamp(f'Berhasil mengirim data sinoptik jam {hour:02} ke bmkgsabang')

                return {'status': 'success', 'message': f'Data sinoptik jam {hour:02} berhasil dikirim ke BMKGsabang'}

        except Exception as e:
            return {'status': 'error', 'message': 'Error: Terjadi kesalahan di server sinoptik.'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('timestamp', type=str, nargs='?', help='Timestamp dalam format YYYYMMDDHH atau range YYYYMMDDHH - YYYYMMDDHH')
    parser.add_argument('--bmkgsatu', action='store_true', help='Langsung mengirim ke bmkgsatu tanpa mengirim ke bmkgsabang')
    
    args = parser.parse_args()
    
    # Jika timestamp tidak diberikan, minta input tahun, bulan, hari, jam secara terpisah
    if not args.timestamp:
        default_timestamp = get_nearest_sent_hour()
        default_year, default_month, default_day, default_hour = (
            default_timestamp[:4],   # YYYY
            default_timestamp[4:6],  # MM
            default_timestamp[6:8],  # DD
            default_timestamp[8:10]  # HH
        )

        # timestamp harus dalam format YYYYMMDDHH
        print("Masukkan waktu pengamatan (kosongkan untuk nilai default):")
        year = input(f"Tahun ({default_year}): ") or default_year
        month = (input(f"Bulan ({default_month}): ") or default_month).zfill(2)
        day = (input(f"Hari ({default_day}): ") or default_day).zfill(2)
        hour = (input(f"Jam ({default_hour}): ") or default_hour).zfill(2)


        args.timestamp = f"{year}{month}{day}{hour}"
        # args.timestamp = f"2025021011-2025021023"

    process_sinoptik(args.timestamp, bmkgsatu_only=args.bmkgsatu)
