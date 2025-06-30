import pandas as pd
import random
import sys
from datetime import datetime, time, timedelta
import re

# ===========================================================================
# == BAGIAN 1: PENGATURAN UTAMA (ATUR DI SINI) ==
# ===========================================================================

# --- 1. NAMA FILE ---
# Nama file Excel yang berisi daftar mata kuliah lengkap.
EXCEL_INPUT_FILE = "jadwal.xlsx"

# Nama file Excel yang berisi ketersediaan dosen.
# Jalankan script 'pembuat_template_ketersediaan.py' jika Anda belum memiliki file ini.
LECTURER_AVAILABILITY_FILE = "ketersediaan_dosen.xlsx"

# Nama file hasil akhir yang akan dibuat.
EXCEL_OUTPUT_FILE = "hasil_jadwal_final.xlsx"


# --- 2. ATURAN JADWAL UMUM ---
# Hari apa saja yang digunakan untuk penjadwalan secara umum.
DAYS_TO_PROCESS = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']

# Daftar ruangan yang tersedia.
ROOM_DATA = [
    # Lantai 3
    {'name': 'B3G', 'floor': 3},{'name': 'B3A', 'floor': 3},
    {'name': 'B3B', 'floor': 3},{'name': 'B3H', 'floor': 3},
    # Lantai 4
    {'name': 'B4A', 'floor': 4}, {'name': 'B4B', 'floor': 4},
    {'name': 'B4C', 'floor': 4}, {'name': 'B4D', 'floor': 4},
    {'name': 'B4E', 'floor': 4}, {'name': 'B4F', 'floor': 4},
    {'name': 'B4G', 'floor': 4}, {'name': 'B4H', 'floor': 4},
    {'name': 'A4A', 'floor': 4}, {'name': 'A4B', 'floor': 4}, 
    {'name': 'A4C', 'floor': 4}, {'name': 'A4D', 'floor': 4},
    # Lantai 5
    {'name': 'B5A', 'floor': 5}, {'name': 'B5B', 'floor': 5},
    {'name': 'B5C', 'floor': 5}, {'name': 'B5D', 'floor': 5},
    {'name': 'B5E', 'floor': 5}, {'name': 'B5F', 'floor': 5},
]

# Aturan jam istirahat.
BREAK_PERIODS = [(time(12, 0), time(13, 0)), (time(18, 0), time(19, 0))]


# --- 3. OPSI TAMBAHAN ---
# Pisahkan jadwal dosen berikut ke dalam sheet mereka sendiri di hasil akhir.
LECTURERS_TO_SEPARATE_SHEETS = [
    'Alun Sujjada, ST., M.Kom',
    'Nugraha, M.Kom',
    'Ivana Lucia Kharisma, M.Kom'
]

# ===========================================================================
# == AKHIR DARI BAGIAN PENGATURAN ==
# ===========================================================================


# --- BAGIAN 2: FUNGSI-FUNGSI UTAMA ---

def load_course_data(filename):
    """
    Memuat daftar mata kuliah yang perlu dijadwalkan dari Excel.
    Fungsi ini sekarang secara eksplisit mengabaikan kolom HARI dan JAM dari file sumber.
    """
    try:
        df = pd.read_excel(filename, sheet_name=0, header=2)
        print(f"✅ Berhasil membaca file mata kuliah: '{filename}'.")
    except Exception as e:
        sys.exit(f"❌ ERROR saat membaca file mata kuliah: {e}")

    df = df.rename(columns=lambda c: str(c).strip().upper())
    df = df.dropna(how='all')

    required = ['MATA KULIAH', 'SKS', 'DOSEN', 'KELAS', 'SEMESTER']
    for col in required:
        if col not in df.columns:
            sys.exit(f"❌ ERROR: Kolom '{col}' tidak ditemukan di file '{filename}'.")

    # <<< PERUBAHAN: Secara eksplisit mengabaikan kolom hari dan jam dari file sumber >>>
    ignored_cols = ['HARI', 'JAM']
    existing_ignored = [col for col in ignored_cols if col in df.columns]
    if existing_ignored:
        print(f"ℹ️ Info: Kolom {existing_ignored} dari file sumber akan diabaikan karena ini adalah generator jadwal sejati.")

    cols_to_fill = ['DOSEN', 'MATA KULIAH', 'SKS', 'SEMESTER']
    for col in cols_to_fill:
        df[col] = df[col].ffill()

    df = df.dropna(subset=['KELAS'])
    df['KELAS'] = df['KELAS'].astype(str)
    df_ti = df[df['KELAS'].str.upper().str.startswith('TI')].copy()
    
    # Hanya mengambil kolom yang benar-benar dibutuhkan untuk penjadwalan
    final_cols_to_use = ['MATA KULIAH', 'SKS', 'DOSEN', 'KELAS', 'SEMESTER']
    df_ti = df_ti[final_cols_to_use]
    
    print(f"ℹ️ Ditemukan {len(df_ti)} mata kuliah TI untuk dijadwalkan.")
    return df_ti.to_dict('records')


def load_lecturer_availability(filename):
    """
    Memuat dan memproses file ketersediaan dosen, termasuk batas SKS harian.
    """
    try:
        df = pd.read_excel(filename)
        print(f"✅ Berhasil membaca file ketersediaan dosen: '{filename}'.")
    except Exception:
        print(f"⚠️ PERINGATAN: File ketersediaan dosen '{filename}' tidak ditemukan. Semua dosen dianggap selalu tersedia.")
        return {}
    
    df = df.rename(columns=lambda c: str(c).strip().title())
    
    availability_rules = {}
    for _, row in df.iterrows():
        lecturer_name = row.get('Name')
        if not lecturer_name or pd.isna(lecturer_name):
            continue

        rules = {}
        # Proses hari ketersediaan
        available_days = row.get('Available Day', 'All')
        if isinstance(available_days, str) and available_days.lower() != 'all':
            rules['days'] = [day.strip().title() for day in available_days.split(',')]

        # Proses jam ketersediaan
        available_times = row.get('Available Times', 'All')
        if isinstance(available_times, str) and available_times.lower() != 'all':
            time_ranges_found = re.findall(r'(\d{2}:\d{2})-(\d{2}:\d{2})', available_times)
            if time_ranges_found:
                rules['time_ranges'] = time_ranges_found

        # Membaca batas SKS harian dari file Excel
        max_sks_daily = row.get('Max Sks Harian')
        if pd.notna(max_sks_daily):
            try:
                rules['max_sks_daily'] = int(max_sks_daily)
            except (ValueError, TypeError):
                pass

        if rules:
            availability_rules[lecturer_name] = rules
            
    print(f"ℹ️ Ditemukan {len(availability_rules)} aturan ketersediaan khusus untuk dosen.")
    return availability_rules


def generate_schedule(course_data, availability_rules):
    """
    Fungsi utama untuk menjalankan algoritma penjadwalan.
    """
    start_time, end_time = time(7, 0), time(22, 0)
    
    time_slots = []
    current_time = datetime.combine(datetime.today(), start_time)
    while current_time.time() < end_time:
        time_slots.append(current_time.time())
        current_time += timedelta(minutes=10)

    room_schedule = {day: {r['name']: {t: None for t in time_slots} for r in ROOM_DATA} for day in DAYS_TO_PROCESS}
    lecturer_schedule = {day: {t: set() for t in time_slots} for day in DAYS_TO_PROCESS}
    
    # Inisialisasi pelacak SKS harian
    lecturers_in_courses = set(c['DOSEN'] for c in course_data if pd.notna(c['DOSEN']))
    lecturer_daily_sks = {day: {lecturer: 0 for lecturer in lecturers_in_courses} for day in DAYS_TO_PROCESS}
    
    random.shuffle(course_data)
    scheduled_list, unscheduled_courses = [], []

    for course in course_data:
        try:
            sks = int(float(course['SKS']))
            duration_minutes = sks * 50
            if duration_minutes <= 0: continue
        except (ValueError, TypeError, AttributeError):
            unscheduled_courses.append(course)
            continue

        placed = False
        lecturer = course['DOSEN'] if pd.notna(course['DOSEN']) else 'DOSEN BELUM ADA'
        
        lecturer_constraints = availability_rules.get(lecturer, {})
        possible_days = lecturer_constraints.get('days', DAYS_TO_PROCESS)
        
        for day in [d for d in possible_days if d in DAYS_TO_PROCESS]:
            if placed: break

            # Validasi batas SKS harian
            daily_sks_limit = lecturer_constraints.get('max_sks_daily', 99)
            current_sks = lecturer_daily_sks.get(day, {}).get(lecturer, 0)
            if (current_sks + sks) > daily_sks_limit:
                continue

            random.shuffle(ROOM_DATA) 
            for room in ROOM_DATA:
                if placed: break
                for start_slot in time_slots:
                    start_datetime = datetime.combine(datetime.today(), start_slot)
                    end_datetime = start_datetime + timedelta(minutes=duration_minutes)
                    
                    slots_to_book = []
                    temp_time = start_datetime
                    while temp_time < end_datetime:
                        slots_to_book.append(temp_time.time())
                        temp_time += timedelta(minutes=10)

                    if not all(s in time_slots for s in slots_to_book): continue

                    is_within_time_range = True
                    if 'time_ranges' in lecturer_constraints:
                        is_in_any_range = False
                        for start_str, end_str in lecturer_constraints['time_ranges']:
                            start_range = datetime.strptime(start_str, "%H:%M").time()
                            end_range = datetime.strptime(end_str, "%H:%M").time()
                            if start_range <= start_datetime.time() and end_datetime.time() <= end_range:
                                is_in_any_range = True; break
                        if not is_in_any_range: is_within_time_range = False
                    if not is_within_time_range: continue

                    is_break_conflict = any(bs <= s < be for s in slots_to_book for bs, be in BREAK_PERIODS)
                    if is_break_conflict: continue
                    
                    is_lecturer_free = lecturer == 'DOSEN BELUM ADA' or all(lecturer not in lecturer_schedule[day][s] for s in slots_to_book)
                    is_room_free = all(room_schedule[day][room['name']][s] is None for s in slots_to_book)

                    if is_lecturer_free and is_room_free:
                        scheduled_list.append({
                            'HARI': day, 'MULAI': start_datetime.strftime("%H:%M"), 'SELESAI': end_datetime.strftime("%H:%M"),
                            'MATAKULIAH': course['MATA KULIAH'], 'SKS': course['SKS'], 'SEMESTER': int(course['SEMESTER']),
                            'KELAS': course['KELAS'], 'DOSEN': lecturer, 'RUANG': room['name'],
                        })
                        for s in slots_to_book:
                            room_schedule[day][room['name']][s] = True
                            if lecturer != 'DOSEN BELUM ADA':
                                lecturer_schedule[day][s].add(lecturer)
                        
                        lecturer_daily_sks[day][lecturer] = current_sks + sks
                        placed = True
                        break
        
        if not placed:
            unscheduled_courses.append(course)

    return scheduled_list, unscheduled_courses


def save_schedule_to_excel(schedule_list, unscheduled_list, availability_rules):
    """
    Fungsi ini menyimpan semua hasil ke file Excel.
    """
    if not schedule_list:
        print("\n⚠️ Tidak ada jadwal yang berhasil dibuat.")
        return

    df_final = pd.DataFrame(schedule_list)
    
    day_order = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
    df_final['HARI'] = pd.Categorical(df_final['HARI'], categories=[d for d in day_order if d in df_final['HARI'].unique()], ordered=True)
    df_final = df_final.sort_values(by=['HARI', 'MULAI'])
    
    cols_order = ['HARI', 'MATAKULIAH', 'KELAS', 'MULAI', 'SELESAI', 'SKS', 'SEMESTER', 'DOSEN', 'RUANG']
    df_final = df_final[cols_order]

    print("\n" + "="*130)
    print("HASIL GENERATOR JADWAL OTOMATIS (LENGKAP)".center(130))
    print("="*130)
    print(df_final.to_string(index=False))

    try:
        with pd.ExcelWriter(EXCEL_OUTPUT_FILE, engine='openpyxl') as writer:
            # Sheet 1: Jadwal Lengkap
            df_final.to_excel(writer, sheet_name='Jadwal Otomatis Lengkap', index=False)
            
            # Sheet 2: Ketersediaan Dosen yang Dibaca
            if availability_rules:
                print("\nMenyimpan sheet ketersediaan dosen...")
                avail_data = []
                for name, rules in availability_rules.items():
                    days = ', '.join(rules.get('days', ['All']))
                    times = ', '.join([f"{r[0]}-{r[1]}" for r in rules.get('time_ranges', [])]) or 'All'
                    max_sks = rules.get('max_sks_daily', '-')
                    avail_data.append({'Name': name, 'Available Day': days, 'Available Times': times, 'Max SKS Harian': max_sks})
                pd.DataFrame(avail_data).to_excel(writer, sheet_name='Ketersediaan Dosen', index=False)
                print("  -> Sheet 'Ketersediaan Dosen' berhasil dibuat.")

            # Sheet-sheet berikutnya: Jadwal per Dosen
            if LECTURERS_TO_SEPARATE_SHEETS:
                print("\nMembuat sheet terpisah untuk dosen...")
                for lecturer_name in LECTURERS_TO_SEPARATE_SHEETS:
                    df_lecturer = df_final[df_final['DOSEN'] == lecturer_name]
                    if not df_lecturer.empty:
                        safe_sheet_name = re.sub(r'[\\/*?:"<>|]', "", lecturer_name)[:30]
                        df_lecturer.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                        print(f"  -> Sheet untuk '{lecturer_name}' berhasil dibuat.")

            if unscheduled_list:
                df_failed = pd.DataFrame(unscheduled_list)[['MATA KULIAH', 'SKS', 'KELAS', 'DOSEN']]
                df_failed.to_excel(writer, sheet_name='Gagal Dijadwalkan', index=False)
        
        print(f"\n✅ Berhasil! Jadwal telah dibuat dan disimpan di file '{EXCEL_OUTPUT_FILE}'")
    except Exception as e:
        print(f"\n❌ Gagal menyimpan file Excel. Error: {e}")


# --- BAGIAN EKSEKUSI PROGRAM ---
if __name__ == "__main__":
    courses = load_course_data(EXCEL_INPUT_FILE)
    availability = load_lecturer_availability(LECTURER_AVAILABILITY_FILE)
    if courses:
        final_schedule, failed_courses = generate_schedule(courses, availability)
        save_schedule_to_excel(final_schedule, failed_courses, availability)

