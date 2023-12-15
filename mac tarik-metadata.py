
from selenium import webdriver
from tkinter import simpledialog
# from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook

from selenium.webdriver.support import  expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import time

print('jalankan')

driver = webdriver.Chrome()
driver.maximize_window()

driver.get("https://indah.bps.go.id/login")
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div/div/main/div/div/div/div[2]/div/div/div/div/div[2]/span/div[6]/button')))

driver.find_element(By.ID, '__BVID__29').send_keys('kominfospmitra')
driver.find_element(By.ID, '__BVID__33').send_keys('2h5Y7nM40Q')
time.sleep(7)

driver.find_element(By.XPATH, '/html/body/div/div/main/div/div/div/div[2]/div/div/div/div/div[2]/span/div[6]/button').click()

awal = simpledialog.askinteger("Masukkan awalan = ", "Masukkan awalan = ")
akhir = simpledialog.askinteger("Masukkan akhir = ", "Masukkan akhir = ")

# buka file excel
wb = load_workbook(filename="Metadata Statistik_14-12-2023 16-07-33.xlsx")
sheetRange = wb['Database']

def tarik_disetujui(current_row):
    # Detail Kegiatan
    judul_kegiatan      = driver.find_element(By.XPATH, '//*[@id="__BVID__69"]/div/div[1]/div/div[2]/div/div[1]/div[2]').text
    tahun               = driver.find_element(By.XPATH, '//*[@id="__BVID__69"]/div/div[1]/div/div[2]/div/div[2]/div[2]').text
    cara_pengumpulan    = driver.find_element(By.XPATH, '//*[@id="__BVID__69"]/div/div[1]/div/div[2]/div/div[3]/div[2]').text
    sektor_kegiatan     = driver.find_element(By.XPATH, '//*[@id="__BVID__69"]/div/div[1]/div/div[2]/div/div[4]/div[2]').text

    # I. PENYELENGGARA
    instansi    = driver.find_element(By.XPATH, '//*[@id="blok_i"]/div[2]/div[1]/div[2]').text
    telepon     = driver.find_element(By.XPATH, '//*[@id="blok_i"]/div[2]/div[2]/div[3]/div[1]/div/div[2]').text
    faksmile    = driver.find_element(By.XPATH, '//*[@id="blok_i"]/div[2]/div[2]/div[3]/div[2]/div/div[2]').text
    email       = driver.find_element(By.XPATH, '//*[@id="blok_i"]/div[2]/div[2]/div[3]/div[3]/div/div[2]').text

    # II. PENANGGUNG JAWAB
    # 2.1 Unit Eselon Penanggung Jawab
    eselon_1    = driver.find_element(By.XPATH, '//*[@id="blok_ii"]/div[2]/div[1]/div[2]/div[1]/div/div[2]').text
    eselon_2    = driver.find_element(By.XPATH, '//*[@id="blok_ii"]/div[2]/div[1]/div[2]/div[2]/div/div[2]').text

    # 2.2 Penanggung Jawab Teknis (setingkat Eselon 3)
    nama        = driver.find_element(By.XPATH, '//*[@id="blok_ii"]/div[2]/div[2]/div[2]/div[2]').text
    jabatan     = driver.find_element(By.XPATH, '//*[@id="blok_ii"]/div[2]/div[2]/div[2]/div[4]').text
    alamat      = driver.find_element(By.XPATH, '//*[@id="blok_ii"]/div[2]/div[2]/div[2]/div[6]').text
    telepon     = driver.find_element(By.XPATH, '//*[@id="blok_ii"]/div[2]/div[2]/div[2]/div[7]/div[1]/div[2]').text
    faksmile    = driver.find_element(By.XPATH, '//*[@id="blok_ii"]/div[2]/div[2]/div[2]/div[7]/div[2]/div[2]').text
    email       = driver.find_element(By.XPATH, '//*[@id="blok_ii"]/div[2]/div[2]/div[2]/div[7]/div[3]/div[2]').text

    # III. PERENCANAAN DAN PERSIAPAN
    latar_belakang  = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[1]/div[2]').text
    tujuan          = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[2]/div[2]').text
    # A. Perencanaan
    Perencanaan_Kegiatan_mulai      = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[2]/td[3]/div').text
    Perencanaan_Kegiatan_selesai    = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[2]/td[4]/div').text
    Desain_mulai                    = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[3]/td[3]/div').text
    Desain_selesai                  = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[3]/td[4]/div').text
    # B. Pengumpulan
    Pengumpulan_Data_mulai      = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[5]/td[3]/div').text
    Pengumpulan_Data_selesai    = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[5]/td[4]/div').text
    # C. Pemeriksaan
    Pengolahan_Data_mulai   = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[7]/td[3]/div').text
    Pengolahan_Data_selesai = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[7]/td[4]/div').text
    # D. Penyebarluasan
    Analisis_mulai              = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[9]/td[3]/div').text
    Analisis_selesai            = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[9]/td[4]/div').text
    Diseminasi_Hasil_mulai      = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[10]/td[3]/div').text
    Diseminasi_Hasil_selesai    = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[10]/td[4]').text
    Evaluasi_mulai              = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[11]/td[3]/div').text
    Evaluasi_selesai            = driver.find_element(By.XPATH, '//*[@id="blok_iii"]/div[2]/div[3]/div[2]/table/tbody/tr[11]/td[4]/div').text

    # IV. DESAIN KEGIATAN
    kegiatan_dilakukan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[1]/div[2]').text
    print('ini yang ke = ', current_row - 5)

    if (kegiatan_dilakukan == 'Berulang'):
        frekuensi_penyelenggaraan   = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[2]/div[2]').text
        tipe_pengumpulan            = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[3]/div[2]').text
        cakupan_wilayah             = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[4]/div[2]').text
        if (cakupan_wilayah == 'Sebagian Wilayah Indonesia'):
            metode_pengumpulan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[6]/div[2]/ul').text
            sarana_pengumpulan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[7]/div[2]/ul').text
            unit_pengumpulan            = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[8]/div[2]/ul').text
        elif (cakupan_wilayah == 'Seluruh Wilayah Indonesia'):
            metode_pengumpulan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[5]/div[2]/ul').text
            sarana_pengumpulan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[6]/div[2]/ul').text
            unit_pengumpulan            = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[7]/div[2]/ul').text

    elif (kegiatan_dilakukan == 'Hanya Sekali'):
        frekuensi_penyelenggaraan   = ''
        tipe_pengumpulan            = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[2]/div[2]').text
        cakupan_wilayah             = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[3]/div[2]').text
        if (cakupan_wilayah == 'Sebagian Wilayah Indonesia'):
            metode_pengumpulan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[5]/div[2]/ul').text
            sarana_pengumpulan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[6]/div[2]/ul').text
            unit_pengumpulan            = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[7]/div[2]/ul').text
        elif (cakupan_wilayah == 'Seluruh Wilayah Indonesia'):
            metode_pengumpulan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[4]/div[2]/ul').text
            sarana_pengumpulan          = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[5]/div[2]/ul').text
            unit_pengumpulan            = driver.find_element(By.XPATH, '//*[@id="blok_iv"]/div[2]/div[6]/div[2]/ul').text

    # VI. PENGUMPULAN DATA
    uji_coba_pilot_survey       = driver.find_element(By.XPATH, '//*[@id="blok_vi"]/div[2]/div[1]/div[2]').text
    metode_pemeriksaan_kualitas = driver.find_element(By.XPATH, '//*[@id="blok_vi"]/div[2]/div[2]/div[2]/ul').text
    penyesuaian_nonrespon       = driver.find_element(By.XPATH, '//*[@id="blok_vi"]/div[2]/div[3]/div[2]').text
    petugas                     = driver.find_element(By.XPATH, '//*[@id="blok_vi"]/div[2]/div[4]/div[2]').text
    pendidikan_terendah         = driver.find_element(By.XPATH, '//*[@id="blok_vi"]/div[2]/div[5]/div[2]').text
    # 6.6. Jumlah Petugas
    supervisor  = driver.find_element(By.XPATH, '//*[@id="blok_vi"]/div[2]/div[6]/div[2]/div[1]/div[2]').text
    pengumpul   = driver.find_element(By.XPATH, '//*[@id="blok_vi"]/div[2]/div[6]/div[2]/div[2]/div[2]').text
    # 6.7. Apakah Melakukan Pelatihan Petugas
    pelatihan   = driver.find_element(By.XPATH, '//*[@id="blok_vi"]/div[2]/div[7]/div[2]').text

    # VII. PENGOLAHAN DAN ANALISIS
    # 7.1. Tahapan Pengolahan Data
    editing     = driver.find_element(By.XPATH, '//*[@id="blok_vii"]/div[2]/div[1]/div[2]/div[1]/div[2]').text
    coding      = driver.find_element(By.XPATH, '//*[@id="blok_vii"]/div[2]/div[1]/div[2]/div[2]/div[2]').text
    data_entri  = driver.find_element(By.XPATH, '//*[@id="blok_vii"]/div[2]/div[1]/div[2]/div[3]/div[2]').text
    validasi    = driver.find_element(By.XPATH, '//*[@id="blok_vii"]/div[2]/div[1]/div[2]/div[4]/div[2]').text
    			
    metode_analisis     = driver.find_element(By.XPATH, '//*[@id="blok_vii"]/div[2]/div[2]/div[2]').text
    unit_analisis       = driver.find_element(By.XPATH, '//*[@id="blok_vii"]/div[2]/div[3]/div[2]/ul').text
    tingkat_penyajian   = driver.find_element(By.XPATH, '//*[@id="blok_vii"]/div[2]/div[4]/div[2]/ul').text

    # VIII. DISEMINASI HASIL
    # 8.1. Produk Kegiatan yang Tersedia untuk Umum
    produk_hardcopy     = driver.find_element(By.XPATH, '//*[@id="blok_viii"]/div[2]/div[1]/div[2]/div[1]/div[2]').text
    produk_softcopy     = driver.find_element(By.XPATH, '//*[@id="blok_viii"]/div[2]/div[1]/div[2]/div[2]/div[2]').text
    produk_data_mikro   = driver.find_element(By.XPATH, '//*[@id="blok_viii"]/div[2]/div[1]/div[2]/div[3]/div[2]').text

    # 8.2. Rencana Rilis Produk Kegiatan
    if ((produk_hardcopy == ': Tidak') and (produk_softcopy == ': Tidak') and (produk_data_mikro == ': Tidak')):
        rilis_hardcopy      = ''
        rilis_softcopy      = ''
        rilis_data_mikro    = ''
    else: 
        if (produk_hardcopy == ': Ya'):
            rilis_hardcopy      = driver.find_element(By.XPATH, '//*[@id="blok_viii"]/div[2]/div[2]/div[2]/div/table/tbody/tr[1]/td[2]').text
        else:
            rilis_hardcopy      = ''

        if (produk_softcopy == ': Ya'):
            rilis_softcopy      = driver.find_element(By.XPATH, '//*[@id="blok_viii"]/div[2]/div[2]/div[2]/div/table/tbody/tr[2]/td[2]').text
        else:
            rilis_softcopy      = ''

        if (produk_data_mikro == ': Ya'):
            rilis_data_mikro    = driver.find_element(By.XPATH, '//*[@id="blok_viii"]/div[2]/div[2]/div[2]/div/table/tbody/tr[3]/td[2]').text
        else:
            rilis_data_mikro    = ''        

    sheetRange['E'+str(current_row)].value = judul_kegiatan
    sheetRange['F'+str(current_row)].value = tahun
    sheetRange['G'+str(current_row)].value = cara_pengumpulan
    sheetRange['H'+str(current_row)].value = sektor_kegiatan
    sheetRange['I'+str(current_row)].value = instansi
    sheetRange['J'+str(current_row)].value = telepon
    sheetRange['K'+str(current_row)].value = faksmile
    sheetRange['L'+str(current_row)].value = email
    sheetRange['M'+str(current_row)].value = eselon_1
    sheetRange['N'+str(current_row)].value = eselon_2
    sheetRange['O'+str(current_row)].value = nama
    sheetRange['P'+str(current_row)].value = jabatan
    sheetRange['Q'+str(current_row)].value = alamat
    sheetRange['R'+str(current_row)].value = telepon
    sheetRange['S'+str(current_row)].value = faksmile
    sheetRange['T'+str(current_row)].value = email
    sheetRange['U'+str(current_row)].value = latar_belakang
    sheetRange['V'+str(current_row)].value = tujuan
    sheetRange['W'+str(current_row)].value = Perencanaan_Kegiatan_mulai
    sheetRange['X'+str(current_row)].value = Perencanaan_Kegiatan_selesai
    sheetRange['Y'+str(current_row)].value = Desain_mulai
    sheetRange['Z'+str(current_row)].value = Desain_selesai
    sheetRange['AA'+str(current_row)].value = Pengumpulan_Data_mulai
    sheetRange['AB'+str(current_row)].value = Pengumpulan_Data_selesai
    sheetRange['AC'+str(current_row)].value = Pengolahan_Data_mulai
    sheetRange['AD'+str(current_row)].value = Pengolahan_Data_selesai
    sheetRange['AE'+str(current_row)].value = Analisis_mulai
    sheetRange['AF'+str(current_row)].value = Analisis_selesai
    sheetRange['AG'+str(current_row)].value = Diseminasi_Hasil_mulai
    sheetRange['AH'+str(current_row)].value = Diseminasi_Hasil_selesai
    sheetRange['AI'+str(current_row)].value = Evaluasi_mulai
    sheetRange['AJ'+str(current_row)].value = Evaluasi_selesai
    sheetRange['AK'+str(current_row)].value = kegiatan_dilakukan
    sheetRange['AL'+str(current_row)].value = frekuensi_penyelenggaraan
    sheetRange['AM'+str(current_row)].value = tipe_pengumpulan
    sheetRange['AN'+str(current_row)].value = cakupan_wilayah
    sheetRange['AO'+str(current_row)].value = metode_pengumpulan
    sheetRange['AP'+str(current_row)].value = sarana_pengumpulan
    sheetRange['AQ'+str(current_row)].value = unit_pengumpulan
    sheetRange['AR'+str(current_row)].value = uji_coba_pilot_survey
    sheetRange['AS'+str(current_row)].value = metode_pemeriksaan_kualitas
    sheetRange['AT'+str(current_row)].value = penyesuaian_nonrespon
    sheetRange['AU'+str(current_row)].value = petugas
    sheetRange['AV'+str(current_row)].value = pendidikan_terendah
    sheetRange['AW'+str(current_row)].value = supervisor
    sheetRange['AX'+str(current_row)].value = pengumpul
    sheetRange['AY'+str(current_row)].value = pelatihan
    sheetRange['AZ'+str(current_row)].value = editing
    sheetRange['BA'+str(current_row)].value = coding
    sheetRange['BB'+str(current_row)].value = data_entri
    sheetRange['BC'+str(current_row)].value = validasi
    sheetRange['BD'+str(current_row)].value = metode_analisis
    sheetRange['BE'+str(current_row)].value = unit_analisis
    sheetRange['BF'+str(current_row)].value = tingkat_penyajian
    sheetRange['BG'+str(current_row)].value = produk_hardcopy
    sheetRange['BH'+str(current_row)].value = produk_softcopy
    sheetRange['BI'+str(current_row)].value = produk_data_mikro
    sheetRange['BJ'+str(current_row)].value = rilis_hardcopy
    sheetRange['BK'+str(current_row)].value = rilis_softcopy
    sheetRange['BL'+str(current_row)].value = rilis_data_mikro


# program utama
for i in range(awal, akhir + 1):
# while awal < akhir + 1:
    current_row = i + 4
    try:
        driver.get("https://indah.bps.go.id/metadata/view-kegiatan/" + str(i))
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="__BVID__69"]/div/div[1]/div/div[1]')))
    except:
        try:
            driver.get("https://indah.bps.go.id/metadata/view-kegiatan/" + str(i))
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="__BVID__69"]/div/div[1]/div/div[1]')))
        except:
            try:
                driver.get("https://indah.bps.go.id/metadata/view-kegiatan/" + str(i))
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="__BVID__69"]/div/div[1]/div/div[1]')))
            except:
                sheetRange['A'+str(current_row)].value = i
                sheetRange['B'+str(current_row)].value = 'Tidak ditemukan'
                sheetRange['C'+str(current_row)].value = 'Tidak ditemukan'
                sheetRange['D'+str(current_row)].value = 'Tidak ditemukan'
                continue
    
    status_approval = driver.find_element(By.XPATH, '//*[@id="__BVID__69"]/div/div[1]/div/div[1]').text
    detil           = driver.find_element(By.XPATH, '//*[@id="detail-pelaporan"]/div[2]/div[1]/div[1]/div/div[1]/div/span').text
    produsen_data   = driver.find_element(By.XPATH, '//*[@id="detail-pelaporan"]/div[2]/div[1]/div[1]/div/div[2]/div').text

    sheetRange['A'+str(current_row)].value = i
    sheetRange['B'+str(current_row)].value = status_approval
    sheetRange['C'+str(current_row)].value = detil
    sheetRange['D'+str(current_row)].value = produsen_data
    
    jumlah_variabel     = driver.find_element(By.XPATH, '//*[@id="detail-pelaporan"]/div[2]/div[2]/div/div[2]/span').text
    jumlah_indikator    = driver.find_element(By.XPATH, '//*[@id="detail-pelaporan"]/div[2]/div[2]/div/div[3]/span').text
    sheetRange['BM'+str(current_row)].value = jumlah_variabel
    sheetRange['BN'+str(current_row)].value = jumlah_indikator

    if(status_approval == 'Disetujui'):
        tarik_disetujui(current_row)

    print('berhasil = ', i)

from datetime import datetime
now = datetime.now()

dt_string = now.strftime("%d-%m-%Y %H-%M-%S")
wb.save("Metadata Statistik_" + dt_string + ".xlsx")