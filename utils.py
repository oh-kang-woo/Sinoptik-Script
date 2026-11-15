import os
import sys
import arrow
import math
import msvcrt
import time
import keyboard
import re
import requests
import urllib3
import datetime
import sqlite3
import xlwings
from fuzzywuzzy import process
from constants import *


# disable InsecureRequestWarning
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

token = None  
token_expiration = None 

SP = ' '
CR = '\r'
LF = '\n'
CRLF = CR + LF


def echo(string):
    sys.stdout.write(string)
    sys.stdout.flush()


def wait_for_input(prompt='', timeout=10, default=None):
    echo(prompt)
    begin = time.monotonic()
    end = begin + timeout
    line = ''

    while time.monotonic() < end:
        if msvcrt.kbhit():
            c = msvcrt.getwche()
            if c in (CR, LF):
                echo(CRLF)
                return line
            if c == '\003':
                raise KeyboardInterrupt
            if c == '\b':
                line = line[:-1]
                cover = SP * len(prompt + line + SP)
                echo(''.join([CR, cover, CR, prompt, line]))
            else:
                line += c
        time.sleep(0.05)

    echo(CRLF)

    return default


def input_with_default(prompt_='', default_=''):
    keyboard.write(default_)
    return input(prompt_)


def print_with_timestamp(message):
    """Cetak pesan dengan timestamp di depannya"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{timestamp}\t{message}")


def round_half_nearest_odd(n):
    """Round .5 to nearest odd number"""
    x, y = str(n).split(".")
    if y[0] == "5":
        if int(x) % 2 == 0:
            return math.ceil(n)
        else:
            return math.floor(n)
    else:
        return round(n)
    

def parse_timestamp(timestamp):
    """
    Parse a timestamp string into a datetime object.
    
    :param timestamp: The timestamp string to parse in 'YYYYMMDDHH' format
    :return: A datetime object representing the parsed timestamp or None if parsing fails.
    """
    
    try:
        # Attempt to parse the timestamp using the current format
        dt = arrow.get(timestamp, 'YYYYMMDDHH')
        
        # Convert the Arrow object to a datetime object
        return dt.datetime  # Return the datetime object directly
    except arrow.parser.ParserError:
        # Continue to the next format if parsing fails
        print_with_timestamp('Kesalahan format timestamp')
    
    # If none of the formats matched, handle the error
    print_with_timestamp(f"Error: Invalid timestamp format. Received: {timestamp}")
    return None  # Return None if parsing fails


# Excel Stuff
#--------------
def get_synop_path(path_template, year, month, day):
    # get synop file path
    path = path_template.replace('{YEAR}', str(year))\
        .replace('{MONTH_NUM}', str(month))\
        .replace('{MONTH_NUM_ZERO}', f"{month:02d}")\
        .replace('{MONTH_NAME_FULL}', MONTHS[month][1])\
        .replace('{MONTH_NAME_SHORT}', MONTHS[month][0])\
        .replace('{FILENAME}', str(day))
    
    if not os.path.exists(path):
        # Coba cek variasi nama folder tanpa spasi
        alternate_path = path.replace('. ', '.')
        if os.path.exists(alternate_path):
            return alternate_path
            
        print_with_timestamp('File tidak ditemukan.')
        print_with_timestamp('Periksa kembali nama-nama folder dan file')
        raise FileNotFoundError('File tidak ditemukan.')


    return path


def get_sheet_name(hour):
    if hour in (0, 1, 2, 3, 4, 5):
        sheet_name = "jam 00 - 05 UTC"
    elif hour in (6, 7, 8, 9, 10, 11):
        sheet_name = "jam 06 - 11 UTC"
    elif hour in (12, 13, 14, 15, 16, 17):
        sheet_name = "jam 12 - 17 UTC"
    else:
        sheet_name = "jam 18 - 23 UTC"

    return sheet_name


def extract_cell_pos(cell, range=0):
    """return tuple of cell position, single cell or in range"""
    try:
        col = re.findall(r'[a-zA-Z]+', cell)
        row = re.findall(r'\d+', cell)
        if ':' in cell:
            return (col[0], row[0], col[1], row[1])
        else:
            return (col[0], row[0])

    except:
        return None


def generate_new_me48_cell_map(hour, old_cell_map):
    """return new me48_dict with new cells position, jam UTC"""
    new_cell_map = {}
    for key, val in old_cell_map.items():
        if key == 'data_timestamp': # fix value
            new_cell_map['data_timestamp'] = old_cell_map['data_timestamp']
            continue
        else:
            if hour in [0, 1, 2, 3, 4, 5]:
                diff = hour * PERCELL_DIFF
            elif hour in [6, 7, 8, 9, 10, 11]:
                diff = hour * PERCELL_DIFF + (PERPAGE_DIFF*1)
            elif hour in [12, 13, 14, 15, 16, 17]:
                diff = hour * PERCELL_DIFF + (PERPAGE_DIFF*2)
            elif hour in [18, 19, 20, 21, 22, 23]:
                diff = hour * PERCELL_DIFF + (PERPAGE_DIFF*3)

        cell = extract_cell_pos(val)
        if len(cell) == 2:
            new_cell_map[key] = cell[0]+str(int(cell[1])+diff)
        elif len(cell) == 4:
            new_cell_map[key] = cell[0]+str(int(cell[1])+diff)+':'+cell[2]+str(int(cell[3])+diff)
        
    return new_cell_map


def generate_new_sinoptik_cell_map(hour, old_cell_map):
    """return new cell positions of sinoptik data in excel, not using me48 worksheet as generate_new_me48_cell_map"""
    PERCELL_DIFF = 5
    new_cell_map = {}
    for key, val in old_cell_map.items():
        if key == 'data_timestamp': # fix value
            new_cell_map['data_timestamp'] = old_cell_map['data_timestamp']
            continue
        else:
            if hour in [0, 6, 12, 18]:
                diff = PERCELL_DIFF * 0
            elif hour in [1, 7, 13, 19]:
                diff = PERCELL_DIFF * 1
            elif hour in [2, 8, 14, 20]:
                diff = PERCELL_DIFF * 2
            elif hour in [3, 9, 15, 21]:
                diff = PERCELL_DIFF * 3
            elif hour in [4, 10, 16, 22]:
                diff = PERCELL_DIFF * 4
            elif hour in [5, 11, 17, 23]:
                diff = PERCELL_DIFF * 5 

        cell = extract_cell_pos(val)
        if len(cell) == 2:
            new_cell_map[key] = cell[0]+str(int(cell[1])+diff)
        elif len(cell) == 4:
            new_cell_map[key] = cell[0]+str(int(cell[1])+diff)+':'+cell[2]+str(int(cell[3])+diff)
        
    return new_cell_map


def clean_value(val, decimal=None):
    """Modify value to accepted value
    :val: str
    :decimal: digit belakang koma"""
    if val != None:
        val = str(val)
        if '.' in val:
            try:
                if decimal:
                    return round(float(val), decimal)
                else:
                    return round(float(val))
            except:
                pass
        
        if val == '' or val.isspace():
            return None

        else:
            return val


def open_workbook(fname, visible=True):
    """Open excel synop file"""
    try:
        wb = xlwings.Book(fname, update_links=False, read_only=False, ignore_read_only_recommended=True)

    except FileNotFoundError:
        print_with_timestamp(f'File {fname} tidak ditemukan.')
        raise FileNotFoundError(f"File {fname} tidak ditemukan. Periksa kembali path dan nama file.")
    
    except OSError:
        wb = xlwings.Book(fname)

    except ValueError:
        import os
        wb_name = fname.split(os.sep)[-1]
        wb = xlwings.Book(wb_name)
        if wb.fullname != fname:
            print_with_timestamp(f'Error. Tidak bisa membuka 2 file {wb_name} bersamaan. Close file terlebih dahulu dan jalankan ulang program.')
            time.sleep(10)
            raise RuntimeError(f"Error membuka 2 file {wb_name} bersamaan.")

    if not visible:
        wb.app.visible = False
    
    wb.app.display_alerts = False

    return wb


def extract_values(workbook, sheet_name, cell_map, autosave=False):
    """Extract Sinoptik data values from synop file"""
    wb = workbook
    if autosave:
        wb.save()
    ws = wb.sheets[sheet_name]
    data = DATA_SINOPTIK
    for key, val in cell_map.items():
        s = None
        if ':' in val: # join val in cell range
            s = [clean_value(i) for i in ws[val].value if i != None]
            if s:
                s = ' '.join([str(i) for i in s if i != None])
                if s == '' or s.isspace():
                    s = None
            else:
                s = None

        else:
            if key in ME48_ROUNDED_VAL:
                s = clean_value(ws[val].value, ME48_ROUNDED_VAL[key])
            else:
                s = clean_value(ws[val].value)

        data[key] = s

    # beberapa penyesuaian data
    data['wind_dir_deg_dd'] = 0 if data['wind_dir_deg_dd'] in ('CALM', 'calm', None, '') else data['wind_dir_deg_dd']
    data['wind_speed_ff'] = 0 if data['wind_dir_deg_dd'] == 0 else data['wind_speed_ff']  
    data['rainfall_last_mm'] = 0 if data['rainfall_last_mm'] in (None, '') else data['rainfall_last_mm']  
    data['rainfall_last_mm'] = 8888 if data['rainfall_last_mm'] in ('TTU', 'ttu') else data['rainfall_last_mm']  
    data['rainfall_6h_rrr'] = 8888 if data['rainfall_6h_rrr'] in ('TTU', 'ttu') else data['rainfall_6h_rrr']  
    data['rainfall_24h_rrrr'] = 8888 if data['rainfall_24h_rrrr'] in ('TTU', 'ttu') else data['rainfall_24h_rrrr']  
    data['rainfall_last_mm'] = 9999 if data['rainfall_last_mm'] in ('////', '///') else data['rainfall_last_mm']  
    data['rainfall_6h_rrr'] = 9999 if data['rainfall_6h_rrr'] in ('////', '///') else data['rainfall_6h_rrr']  
    data['rainfall_24h_rrrr'] = 9999 if data['rainfall_24h_rrrr'] in ('////', '///') else data['rainfall_24h_rrrr']  
    data['pressure_reading_mb'] = 9999 if data['pressure_reading_mb'] in ('////', '///') else data['pressure_reading_mb']  
    data['pressure_tend_3h_mb_ppp'] = 9999 if data['pressure_tend_3h_mb_ppp'] in ('////', '///') else data['pressure_tend_3h_mb_ppp']  
    data['pressure_tend_24h_mb_ppp'] = 9999 if data['pressure_tend_24h_mb_ppp'] in ('////', '///') else data['pressure_tend_24h_mb_ppp']  
    data['temp_max_c_txtxtx'] = 9999 if data['temp_max_c_txtxtx'] in ('////', '///') else data['temp_max_c_txtxtx']  
    data['temp_min_c_tntntn'] = 9999 if data['temp_min_c_tntntn'] in ('////', '///') else data['temp_min_c_tntntn']  
    data['evaporation_eq_indicator_ie'] = 0 if data['evaporation_eq_indicator_ie'] in (None, '') else data['evaporation_eq_indicator_ie']  

    # Pertahankan nilai Kondisi Tanah dari Excel (jangan ambil dari evaporation_eq_indicator_ie)
    raw_land_cond = data['land_cond']
    if raw_land_cond in (None, ''):
        data['land_cond'] = 0
    else:
        land_cond = int(float(raw_land_cond))
        if land_cond not in SOIL:
            raise ValueError(f"Kode land_cond tidak valid: {raw_land_cond}")
        data['land_cond'] = land_cond
      
    data['encoded_synop'] = data['encoded_synop'].strip().replace('  ', ' ')  
    temp = []
    # excel menghapus 0 dari angka 2004, harusnya 02004, kembalikan
    for i in data['encoded_synop'].split():  
        if i.isdecimal() and len(i) == 4:  
            temp.append("0"+i)  
        else:  
            temp.append(i)  
      
    data['encoded_synop'] = ' '.join(temp)  

    return data


# khusus perawanan
def get_cloud_cell(hour):
    """return cell location for cloud 8xxxx"""
    cell = None
    if hour in (0, 6, 12, 18):
        cell = 'W14'
    elif hour in (1, 7, 13, 19):
        cell = 'R19'
    elif hour in (2, 8, 14, 20):
        cell = 'R24'
    elif hour in (4, 10, 16, 22):
        cell = 'R34'
    elif hour in (5, 11, 17, 23):
        cell = 'R39'
    elif hour in (3, 9, 15, 21):
        cell = 'T29'

    return cell


# khusus perawanan konvektif
def get_convective_cloud_cell(hour):
    """return cell location for convective cloud, first and second 57xxx sections"""
    cell = None
    if hour == 0:
        cell = ('AI14','AM14')
    elif hour == 12:
        cell = ('AE14', 'AG14') 
    elif hour in  (6, 18):
        cell = ('AC14','AF14')
    elif hour in (1, 7, 13, 19):
        cell = ('Y19', 'AB19')
    elif hour in (2, 8, 14, 20):
        cell = ('Y24', 'AB24')
    elif hour in (3, 9, 15, 21):
        cell = ('AB29', 'AE29')
    elif hour in (4, 10, 16, 22):
        cell = ('Y34', 'AB34')
    elif hour in (5, 11, 17, 23):
        cell = ('Y39', 'AB39')
    
    return cell


# tinggi awan konvektif
def get_cl_peak_cell(hour):
    """return cell location for convective cloud, first and second 80xxx sections"""
    cell = None
    if hour == 0:
        cell = ('BA14','BC14') 
    elif hour == 12:
        cell = ('AW14','AZ14')
    elif hour in (6, 18):
        cell = ('AX14','BA14')
    elif hour in (1, 7, 13, 19):
        cell = ('AP19', 'AS19')
    elif hour in (2, 8, 14, 20):
        cell = ('AP24', 'AS24')
    elif hour in (3, 9, 15, 21):
        cell = ('AW29', 'AZ29')
    elif hour in (4, 10, 16, 22):
        cell = ('AP34', 'AS34')
    elif hour in (5, 11, 17, 23):
        cell = ('AP39', 'AS39')
    
    return cell


def normalize_data(data):
    # currently only wind
    if data:
        data['wind_speed_ff'] = data['wind_speed_ff'] * 1.94384  # m/s to knot
        if data['wind_speed_ff'] < 4.0:
            data['wind_speed_ff'] = 0
            data['wind_dir_deg_dd'] = 0
        else:
            data['wind_speed_ff'] = round(data['wind_speed_ff'])

        data['wind_dir_deg_dd'] = round(data['wind_dir_deg_dd']/10) * 10

    return data


def extract_cloud(synop_code):
    cl_pattern = r'(8[/\d]{4})(?:\s+333)'
    return re.search(cl_pattern, synop_code).group(1)


def get_wind_direction(degree):
    if degree % 5 != 0: 
        degree = 5 * round(degree / 5)
        return get_wind_direction(degree)
    
    if str(degree) in ("CALM", "calm", "0", 'None'):
        return "CALM"
    
    else:
        degree = int(degree)
        if 25 <= degree <= 65:
            return "NE"
        elif 70 <= degree <= 110:
            return "E"
        elif 115 <= degree <= 155:
            return "SE"
        elif 160 <= degree <= 200:
            return "S"
        elif 205 <= degree <= 245:
            return "SW"
        elif 250 <= degree <= 290:
            return "W"
        elif 295 <= degree <= 335:
            return "NW"
        elif 340 <= degree <= 360 or 0 <= degree <= 20:
            return "N"


def get_previous_wind_direction(timestamp):
    """In case of failure of getting wind direction from AWS"""
    previous_timestamp = arrow.get(timestamp).shift(hours=-1).format('YYYY-MM-DD HH:mm:ss')
    previous_wind = get_data_from_db_based_on_timestamp(previous_timestamp, column='wind_dir_deg_dd')

    return previous_wind


def calculate_tbb(t, rh, td):
    result = 0
    tbb = td
    while result <= rh:
        result = (((6.11*10**(7.5*tbb/(237.3+tbb)))-(0.7947*(t-tbb)))/(6.11*10**(7.5*t/(237.3+t))))*100
        if result < rh:
            tbb += 0.1
    return round(tbb, 1)


def calculate_hshs(cl_peak):
    # kode tinggi puncak awan rendah
    if cl_peak <= 1500:
        return cl_peak // 30
    elif cl_peak >= 1800:
        return cl_peak // 300 + 50
    

def calculate_tendency(plist):
    tendency = None
    pdiff = round(plist[3] - plist[0], 1)
    # kondisi 1 atau 2, 6 atau 7, atau 4
    if plist[0] < plist[1] < plist[2] < plist[3]:
        # naik konstan
        if round(plist[1] - plist[0],1) == round(plist[2] - plist[1],1) == round(plist[3]-plist[2],1):
            tendency = 2
        # naik tidak konstan
        elif round(plist[1] - plist[0],1) < round(plist[2] - plist[1],1) < round(plist[3] - plist[2],1):
            tendency = 2
        else:
            tendency = 1
    elif plist[0] > plist[1] > plist[2] > plist[3]:
        # turun konstan
        if round(plist[1] - plist[0],1) == round(plist[2] - plist[1],1) == round(plist[3]-plist[2],1):
            tendency = 7
        # turun tidak konstan
        elif round(plist[0] - plist[1],1) < round(plist[1] - plist[2],1) < round(plist[2] - plist[3],1):
            tendency = 7
        else:
            tendency = 6
    elif plist[0] == plist[1] == plist[2] == plist[3]:
        tendency = 4
    else:
        # kondisi 0, 1, atau 3
        if plist[0] <= plist[3]:
            # kondisi pasti 0
            if plist[0] == plist[3]:
                tendency = 0
            else:
                # kondisi pasti 3
                if plist[0] >= plist[1]:
                    tendency = 3
                # kondisi 0 atau 1
                else:
                    if plist[1] <= plist[2]:
                        tendency = 1
                    else:
                        tendency = 0

        elif plist[0] >= plist[3]:
            # kondisi 5, 6, atau 8
            if plist[0] == plist[3]:
                tendency = 5
            else:
                # kondisi pasti 8
                if plist[0] <= plist[1]:
                    tendency = 8
                # kondisi 5 atau 6
                else:
                    if plist[1] >= plist[2]:
                        tendency = 6
                    else:
                        tendency = 5

    tendency = f'{tendency}{round(abs(pdiff/10),2):.2f}'.replace('.','').replace(',','')

    return tendency


# Local Database Stuff
#---------------------
def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except sqlite3.Error as e:
        print_with_timestamp(e)

    return conn


def read_data(table, condition=None, column='all'):  
    """Read data record. Return list of dictionaries with column names as keys."""  
    conn = create_connection(DB_NAME)  
    cursor = conn.cursor()  
      
    if condition:  
        if column == 'all':  
            sql = "SELECT * FROM %s WHERE %s" % (table, condition)  
        else:  
            sql = "SELECT %s FROM %s WHERE %s" % (column, table, condition)  
    else:  
        if column == 'all':  
            sql = "SELECT * FROM %s" % (table,)  
        else:  
            sql = "SELECT %s FROM %s" % (column, table,)  
      
    cursor.execute(sql)  
    results = cursor.fetchall()  
      
    # Get column names from cursor description  
    column_names = [description[0] for description in cursor.description]  
      
    if results:  
        # Convert each row to a dictionary with column names as keys  
        result_dicts = [dict(zip(column_names, row)) for row in results]  
        return result_dicts  
    else:  
        return None  


def check_data_exist(table, condition):
    conn = create_connection(DB_NAME)
    cursor = conn.cursor()
    sql = "SELECT * FROM %s WHERE %s"%(table, condition)
    cursor.execute(sql)
    results = cursor.fetchall()
    if results:
        return True
    else:
        return False

def insert_data(conn, table, data):
    """Insert data, masih perlu optimasi, tapi untuk lokal masih bisalah"""
    cursor = conn.cursor()
    
    # Get existing table columns
    cursor.execute(f"PRAGMA table_info({table})")
    table_columns = [column[1] for column in cursor.fetchall()]
    
    # Filter data to only include existing columns
    filtered_data = {k: v for k, v in data.items() if k in table_columns}
    
    if not filtered_data:
        print_with_timestamp('No valid columns found in data.')
        return
    
    placeholders = ', '.join(['?'] * len(filtered_data))
    columns = ', '.join(filtered_data.keys())
    sql = "INSERT INTO %s ( %s ) VALUES ( %s )" % (table, columns, placeholders)
    
    try:
        cursor.execute(sql, list(filtered_data.values()))
        conn.commit()
    except sqlite3.IntegrityError as err:
        print_with_timestamp('Data di jam yang sama sudah ada.')


def update_data(table, data, condition=None):
    conn = create_connection(DB_NAME)
    cursor = conn.cursor()
    placeholders = ', '.join([i+'=?' for i in data.keys()])
    sql = "UPDATE %s SET %s WHERE %s" % (table, placeholders, condition)
    try:
        cursor.execute(sql, list(data.values()))
    except sqlite3.OperationalError:
        pass
    conn.commit()


def delete_data(table, condition=None):
    conn = create_connection(DB_NAME)
    cursor = conn.cursor()
    sql = "DELETE FROM %s WHERE %s" % (table, condition)
    cursor.execute(sql)
    print_with_timestamp('Berhasil menghapus data.')
    conn.commit()


def delete_all(conn, table):
    sql = 'DELETE FROM %s'%(table,)
    cur = conn.cursor()
    cur.execute(sql)
    conn.commit()


def get_data_from_db_based_on_timestamp(timestamp, table=ME48_TABLE_NAME, column='all'):
    """Get data from db based on timestamp, string can be in YYYYMMDDHH or YYYY-MM-DD HH:mm:ss"""
    if "-" not in timestamp:
        formatted_timestamp = arrow.get(timestamp, 'YYYYMMDDHH').format('YYYY-MM-DD HH:mm:ss')
    else:
        formatted_timestamp = timestamp
    condition = f"data_timestamp=\"{formatted_timestamp}\""
    data = read_data(table, condition, column)
    if (data):
        if column == 'all':
            data = data[0]
        else:
            data = data[0][column]

    return data


def save_data_to_db(data_timestamp, db_name, table, data, update='N'):
    conn = create_connection(db_name)
    data['data_timestamp'] = arrow.get(data_timestamp).format('YYYY-MM-DD HH:00:00')
    cannot_empty_data = [data.get(key) for key in CANNOT_EMPTY_LIST]
    if all(v != None for v in cannot_empty_data) and all(v != '' for v in cannot_empty_data):
        if update == 'Y':
            condition = f"data_timestamp=\"{arrow.get(data_timestamp).format('YYYY-MM-DD HH:00:00')}\""
            update_data(table, data, condition)
            print_with_timestamp(f"Data {data['data_timestamp']} berhasil diupdate.")
        else:
            insert_data(conn, table, data)
            print_with_timestamp(f"Data {data['data_timestamp']} berhasil disimpan ke database.")

    conn.close()
 

def get_yesterday_data(timestamp, column='all'):
    yesterday = arrow.get(f'{timestamp}').shift(days=-1)
    condition = f"date(data_timestamp)='{yesterday.year:04}-{yesterday.month:02}-{yesterday.day:02}'"
    data = read_data(ME48_TABLE_NAME, condition, column)

    return data
    

def save_sinoptik_data(timestamp, bmkgsatu_only=False):
    """
    Save data to the database. The timestamp can be a single timestamp in 'YYYYMMDDHH' format, or a range of timestamps
    specified with a hyphen (e.g., "2024100500 - 2023100523").
    :param timestamp: A single timestamp string or a range of timestamps as a string.
    """
    # Split the input string by hyphen to determine if it's a range
    if '-' in timestamp:
        start_ts, end_ts = map(str.strip, timestamp.strip().split('-'))
        start_dt = parse_timestamp(start_ts)
        end_dt = parse_timestamp(end_ts)
    else:
        current_dt = parse_timestamp(timestamp)
        start_dt = current_dt
        end_dt = current_dt

    data_dict = {}
    current_dt = start_dt

    while current_dt <= end_dt:
        try:
            synop_file = get_synop_path(PATH_TO_SYNOP, current_dt.year, current_dt.month, current_dt.day)
            wb = open_workbook(synop_file)

            cell_map = generate_new_me48_cell_map(current_dt.hour, ME48_CELL_MAP)
            if (arrow.utcnow() - current_dt).days == 0:
                data = extract_values(wb, ME48_WS_NAME, cell_map, autosave=True)
            else:
                data = extract_values(wb, ME48_WS_NAME, cell_map)

            condition = f"data_timestamp=\"{arrow.get(current_dt).format('YYYY-MM-DD HH:00:00')}\""
            # QC Data list that not pass
            errors = simple_qc(current_dt, data)
            if errors:
                raise ValueError(f"{str(errors)}")

            # Jika data sudah ada, update, jika tidak, insert baru
            update_flag = 'Y' if check_data_exist(ME48_TABLE_NAME, condition) else 'N'
            save_data_to_db(current_dt, DB_NAME, ME48_TABLE_NAME, data, update=update_flag)

            db_data = read_data(ME48_TABLE_NAME, condition)
            if db_data:
                db_data = db_data[0]
                # update jadi semua properties data sinoptik
                data.update(DATA_SINOPTIK)
                data_dict[arrow.get(current_dt).format('YYYYMMDDHH')] = db_data
            else:
                print_with_timestamp(f"Data sinoptik {arrow.get(current_dt).format('YYYYMMDDHH')} tidak ditemukan.")

        except (FileNotFoundError, ValueError, Exception) as e:
            print_with_timestamp(f"Error pada {arrow.get(current_dt).format('YYYYMMDDHH')}: {str(e)}")
            raise Exception(f"Terjadi error pada {arrow.get(current_dt).format('YYYYMMDDHH')}: {str(e)}")
 

        # Move to the next hour
        current_dt = arrow.get(current_dt).shift(hours=1)
    
        # Close excel if no longer needed
        if (arrow.utcnow() - arrow.get(current_dt)).days > 1:
            wb.app.quit()
    # return dictionary of data_dict with respective YYYYMMDDHH key
    return data_dict


# QC Stuff
#------------
def get_last_tbk(timestamp):
    """Get last 24 hour tbk records and last 12 hour tbb records"""
    timestamp = arrow.get(timestamp)
    year = timestamp.year
    month = timestamp.month
    day = timestamp.day
    hour = timestamp.hour
    tbk_min = None
    tbk_maks = None
    if hour == 0:
        yesterday = arrow.get(f'{year:04}-{month:02}-{day:02}').shift(days=-1)
        condition = f"date(data_timestamp)='{yesterday.year:04}-{yesterday.month:02}-{yesterday.day:02}'"
        tbk_min_list = read_data(ME48_TABLE_NAME, condition, column='data_timestamp, temp_drybulb_c_tttttt') 
        if tbk_min_list:
            tbk_min = min([i for i in tbk_min_list[1:]], key=lambda t: t['temp_drybulb_c_tttttt']) # exclude last 00 UTC
        return tbk_min
    elif hour == 12:
        condition = f"date(data_timestamp)='{year:04}-{month:02}-{day:02}'"
        tbk_maks_list = read_data(ME48_TABLE_NAME, condition, column='data_timestamp, temp_drybulb_c_tttttt')[:12] # 12 jam terakhir
        if tbk_maks_list:
            tbk_maks = max([i for i in tbk_maks_list], key=lambda t: t['temp_drybulb_c_tttttt'])
        return tbk_maks


def simple_qc(timestamp, data):
    """Perform some QC.
    data: dict
    timestamp: string of datetime
    return list of QC errors (empty list if no errors).
    """
    timestamp = arrow.get(timestamp)
    hour = timestamp.hour
    qc_list = []

    # Pastikan field yang tidak boleh kosong
    for k in CANNOT_EMPTY_LIST:
        if not data.get(k):  # Cek jika None, '', atau 0
            qc_list.append(f'{k} tidak boleh kosong!')

    # Suhu bola kering vs bola basah
    try:
        dry_bulb = data.get('temp_drybulb_c_tttttt')
        wet_bulb = data.get('temp_wetbulb_c')

        if dry_bulb is not None and wet_bulb is not None:
            if dry_bulb < wet_bulb:
                qc_list.append(f'Suhu bola kering ({dry_bulb} C) < suhu bola basah ({wet_bulb} C) jam {hour:02d}.')
    except Exception as e:
        qc_list.append(f'Error dalam pengecekan suhu bola kering & basah: {e}')

    # Suhu minimum dan maksimum (hanya pada jam tertentu)
    try:
        if hour == 0:  # Suhu minimum diperiksa saat jam 00:00
            tbk_min = get_last_tbk(timestamp)
            if tbk_min:
                utc_hour = arrow.get(tbk_min['data_timestamp']).format('HH:00 UTC')
                temp_min = data.get('temp_min_c_tntntn')

                if temp_min is None or not isinstance(temp_min, (float, int)):
                    qc_list.append(f"Suhu minimum tidak diisi.")
                elif temp_min > tbk_min['temp_drybulb_c_tttttt']:
                    qc_list.append(
                        f"Suhu minimum {temp_min} C > suhu bola kering 24 jam terakhir ({tbk_min['temp_drybulb_c_tttttt']} C) jam {utc_hour}."
                    )

        elif hour == 12:  # Suhu maksimum diperiksa saat jam 12:00
            tbk_max = get_last_tbk(timestamp)
            if tbk_max:
                utc_hour = arrow.get(tbk_max['data_timestamp']).format('HH:00 UTC')
                temp_max = data.get('temp_max_c_txtxtx')

                if temp_max is None or not isinstance(temp_max, (float, int)):
                    qc_list.append(f"Suhu maksimum tidak diisi.")
                elif temp_max < tbk_max['temp_drybulb_c_tttttt']:
                    qc_list.append(
                        f"Suhu maksimum {temp_max} C < suhu bola kering 12 jam terakhir ({tbk_max['temp_drybulb_c_tttttt']} C) jam {utc_hour}."
                    )
    except Exception as e:
        qc_list.append(f'Error dalam pengecekan suhu min/max: {e}')

    # Perawanan
    try:
        extract_cloud(data['encoded_synop'])  # Jika gagal, berarti kode salah
    except Exception:
        qc_list.append(f'Kode perawanan jam {hour:02d} salah atau tidak diisi.')

    return qc_list  # Kembalikan daftar error


# BMKGSATU Stuff
#------------------
def get_nearest_sent_hour():
    """Menghasilkan timestamp dalam format 'YYYYMMDDHH' berdasarkan aturan menit."""
    now = arrow.utcnow()
    minute = now.minute

    if 10 <= minute <= 50:
        nearest_hour = now.floor('hour')
    else:
        nearest_hour = now.floor('hour') if minute < 10 else now.shift(hours=1).floor('hour')

    return nearest_hour.format("YYYYMMDDHH")


def get_obs_id(name):
    obs_name_list = list(OBSERVER_ID)
    if name:
        name = process.extractOne(name, obs_name_list)[0]
        if (name):
            obs_id = OBSERVER_ID[name]
            return obs_id
    return OBSERVER_ID['default']


def create_sinoptik_data_timestamp(data_timestamp=None):
    """Create id for data sinoptik from %Y-%m-%d %H:%M:%S to %Y-%m-%d T%H"""
    if data_timestamp is None:
        data_timestamp = arrow.now()
    
    dt = arrow.get(data_timestamp)
    formatted_timestamp = dt.format('YYYY-MM-DD[T]HH')
    # id is 1_sinoptik_ + formatted_timestamp
    return formatted_timestamp


def create_sinoptik_data_id(data_timestamp=None):
    """Create id for data sinoptik from %Y-%m-%d %H:%M:%S to %Y-%m-%d T%H"""
    if data_timestamp is None:
        data_timestamp = arrow.now()
    
    formatted_timestamp = create_sinoptik_data_timestamp(data_timestamp)

    # id is 1_sinoptik_ + formatted_timestamp
    return '1_sinoptik_' + formatted_timestamp


def create_encoded_synop(synop_string):
    """Create encoded synop message using synop string from excel"""
    encoded_synop_str = synop_string.replace(' Z ', ' \n').replace(' 333 ', ' \n  333 ')
    return encoded_synop_str


def create_encoded_synop_with_header(synop_string, data_timestamp):
    """Create encoded synop message with header"""
    encoded_synop_str = synop_string.replace(' Z ', ' \n').replace(' 333 ', ' \n  333 ')
    dt = arrow.get(data_timestamp)
    dt_formatted = dt.format('DDHH') + '00'
    if dt.hour in (0, 6, 12, 18):
        heading = SYNOP_HEADING_UTAMA + ' ' + STN_ID + ' ' + dt_formatted
    elif dt.hour in (3, 9, 15, 21):
        heading = SYNOP_HEADING_ANTARA + ' ' + STN_ID + ' ' + dt_formatted
    else:
        heading = SYNOP_HEADING_OTHER + ' ' + STN_ID + ' ' + dt_formatted

    return heading + '\n' + encoded_synop_str


def create_gts_data_id_and_timestamp(sinoptik_data_timestamp):
    """Create id for gts_data and return tuple of three"""
    now = arrow.now()
    gts_id = '01' + str(STN_CODE) + now.format('YYYYMMDDHHmmss')
    timestamp_sent_data = now.format('YYYY-MM-DDTHH:mm:ss')
    timestamp_data = sinoptik_data_timestamp+':00'

    return (gts_id, timestamp_sent_data, timestamp_data)


def create_gts_data(sinoptik_data):
    gts_data = GTS_DATA_INIT
    gts_id, timestamp_sent_data, timestamp_data = create_gts_data_id_and_timestamp(sinoptik_data['data_timestamp'])
    gts_data.update({
        'id': gts_id,
        'timestamp_sent_data': timestamp_sent_data,
        'timestamp_data': timestamp_data,
        'sandi_gts': sinoptik_data['encoded_synop_with_header'],
        'cccc': sinoptik_data['encoded_synop_with_header'].split('\n', 1)[0].strip()
    })

    return gts_data


def mutate_sinoptik_data_before_sent(data=None, bmkgsatu_only=False):
    """Create data in sinoptik format ready to save and send
    data: dict sinoptik data to be sent
    bmkgsatu_only: determine if sent data to bmkgsatu via bmkgsabang or not
    if false, then some properties will be modified at bmkgsabang server"""

    sinoptik_data = data.copy()

    # common
    sinoptik_data["observer_id"] = get_obs_id(sinoptik_data["observer_id"])
    cloud_code = extract_cloud(sinoptik_data['encoded_synop'])
    if not cloud_code:  
        cloud_code = CLOUD_DEFAULT_CODE
    else:
        sinoptik_data['cloud_cover_oktas_m'] = cloud_code[1]  
        sinoptik_data['cloud_low_type_cl'] = cloud_code[2]  
        sinoptik_data['cloud_med_type_cm'] = cloud_code[3]  
        sinoptik_data['cloud_high_type_ch'] = cloud_code[4]
    
    for key in ME48_LIST:  
        val = sinoptik_data[key]  
          
        # kode hujan TTU dan lainnya  
        if val in (0.0, 8888.0, 9999.0):  
            if val == 0.0:  
                val = 0  
            elif val == 8888.0:  
                val = 8888  
            elif val == 9999.0:  
                val = 9999  
            sinoptik_data[key] = val  
  
        if isinstance(val, str) and key not in ('encoded_synop',):  
            val = val.lower()  
            sinoptik_data[key] = val
  
        # convert all string value except cl, nh, cm, ch, observer  
        if key in ME48_OPTS and key not in ('cloud_low_type_cl', 'cloud_cover_oktas_m', 'cloud_med_type_cm', 'cloud_high_type_ch', 'observer_id') and isinstance(val, str):  
            val = ME48_OPTS[key][val]  
            sinoptik_data[key] = val

    keys = [
        "cloud_low_type_cl", "cloud_low_cover_oktas", "cloud_med_cover_oktas",
        "cloud_high_cover_oktas", "cloud_med_type_cm", "cloud_high_type_ch",
        "cloud_cover_oktas_m", "cloud_layer_1_type_c", "cloud_layer_2_type_c",
        "cloud_layer_3_type_c","cloud_low_type_1", "cloud_med_type_1", "cloud_high_type_1", 
        "cloud_low_cover_1", "cloud_med_cover_1", "cloud_high_cover_1", "rainfall_24h_rrrr"
        ]
    
    for key in keys:  
        sinoptik_data[key] = str(sinoptik_data.get(key)) if sinoptik_data.get(key) is not None else None

    # only when direct to bmkgsatu
    if bmkgsatu_only:
        sinoptik_data['encoded_synop'] = create_encoded_synop(sinoptik_data['encoded_synop'])
        sinoptik_data['encoded_synop_with_header'] = create_encoded_synop_with_header(sinoptik_data['encoded_synop'], sinoptik_data['data_timestamp'])
        sinoptik_data["data_timestamp"] = create_sinoptik_data_timestamp(sinoptik_data["data_timestamp"])
        sinoptik_data["id"] = create_sinoptik_data_id(sinoptik_data["data_timestamp"])
        sinoptik_data["entry_user"] = 1
        sinoptik_data["sandi_tr"] = 7
        sinoptik_data["qc_flag"] = 0
        sinoptik_data["qc_stepcheck_flag"] = 0
        sinoptik_data["qc_histories"] = "{\"before\":null,\"after\":null}"
        sinoptik_data["qc_stepcheck_histories"] = "{\"before\":null,\"after\":null}"
        sinoptik_data["@type"]= "Sinoptik"

    # send via bmkgsabang
    else:
        sinoptik_data.pop('id', None)

    return sinoptik_data


def send_to_bmkgsatu(url=None, data=None):
    login_url = "https://bmkgsatu.bmkg.go.id/db/bmkgsatu//@login"
    auth_data = {"username": "96001","password": "opr96001"}
    token = None
    login_response = requests.post(f'{login_url}', json=auth_data, verify=False, timeout=30)
    if (login_response):
        token = login_response.json()['token']
    if (token):
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        response = requests.request("POST", url, headers=headers, json=data, verify=False)

    return response


def send_to_bmkgsabang(url=None, data=None):
    """Send data to bmkgsabang"""
    token = BMKGSABANG_TOKEN
    if (token):
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        response = requests.request("POST", url, headers=headers, json=data, verify=False)

        if response.status_code == 409:  # Jika data sudah ada
            print_with_timestamp("Data sudah ada, mengupdate dengan data terbaru...")
            # Ambil ID dari respons server
            existing_data = response.json()
            resource_id = existing_data.get('data', {}).get('id')
            
            if resource_id:
                # Update URL dengan ID resource
                update_url = f"{url}/{resource_id}"
                # Kirim permintaan PUT untuk memperbarui resource
                response = requests.request("PUT", update_url, headers=headers, json=data, verify=False)

    return response


def send_confirmation(message, confirm_timeout=10):
    print('\nBerikut data yang akan dikirimkan:\n')
    print(message)
    print()
    response = wait_for_input('Konfirmasi pengiriman? Yes/No/Reload? (Y/N/R): ', timeout=confirm_timeout, default='Y')
    
    return response

if __name__ == '__main__':
    pass