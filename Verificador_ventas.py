from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
import time
import psycopg2
import os
import shutil
import pandas as pd

# --- CONFIGURACIÓN DE RUTAS Y DB ---
DOWNLOAD_DIR = os.path.normpath(os.path.join(os.path.expanduser("~"), "Downloads"))
TARGET_DIR = r"Ruta de la carpeta"
DB_URL = "postgresql://postgres:XXXX@XXXXXX/XXXXXXXX"

if not os.path.exists(TARGET_DIR):
    os.makedirs(TARGET_DIR)

# --- FUNCIONES DE SOPORTE ---

def hacer_login(driver, wait):
    """Realiza el login de forma limpia."""
    print("🔑 Iniciando sesión...")
    driver.get("https://erp.com/")
    try:
        user_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[1]/input")))
        user_input.clear()
        user_input.send_keys("luis.cornieles@riomarket.com")
        
        pass_input = driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[2]/input")
        pass_input.clear()
        pass_input.send_keys("lkdrc291")
        
        driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[4]/button/span").click()
        wait.until(EC.url_changes("https://erp.com/login")) 
        time.sleep(2)
    except Exception as e:
        print(f"❌ Error durante el login: {e}")

def verificar_y_reloguear(driver, wait):
    """Si la URL actual es la de login, re-autentica."""
    if "login" in driver.current_url or len(driver.find_elements(By.XPATH, "//input[@type='password']")) > 0:
        print("⚠️ Sesión expirada o redirigido al login. Re-autenticando...")
        hacer_login(driver, wait)
        return True
    return False

def esperar_descarga_estricta(directory, timeout=120):
    """Espera a que no haya archivos temporales y el archivo final esté listo."""
    seconds = 0
    while seconds < timeout:
        files = os.listdir(directory)
        descargando = any(f.endswith(".crdownload") or f.endswith(".tmp") for f in files)
        xlsx_files = [os.path.join(directory, f) for f in files if f.endswith(".xlsx")]
        
        if not descargando and xlsx_files:
            ultimo_archivo = max(xlsx_files, key=os.path.getctime)
            try:
                time.sleep(2) 
                with open(ultimo_archivo, 'rb') as f:
                    pass
                return ultimo_archivo
            except IOError:
                pass 
                
        time.sleep(2)
        seconds += 2
    return None

# --- SOLICITUD DE FECHAS POR TECLADO ---
def input_fecha(prompt):
    while True:
        try:
            fecha_str = input(prompt)
            fecha = datetime.strptime(fecha_str, "%Y-%m-%d").date()
            return fecha
        except ValueError:
            print("❌ Formato incorrecto. Use AAAA-MM-DD (ejemplo: 2025-01-15)")

print("📅 Rango de fechas a procesar:")
fecha_inicio = input_fecha("Ingrese fecha de inicio (AAAA-MM-DD): ")
fecha_fin = input_fecha("Ingrese fecha de fin (AAAA-MM-DD): ")

if fecha_inicio > fecha_fin:
    print("❌ La fecha de inicio no puede ser mayor que la fecha de fin. Saliendo.")
    exit(1)

print(f"✅ Procesando desde {fecha_inicio} hasta {fecha_fin}")

# --- INICIO DEL DRIVER ---
chrome_options = Options()

chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--start-maximized")

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 30)

# --- LISTA PARA ALMACENAR LAS DIFERENCIAS ENCONTRADAS ---
diferencias_report = []

try:
    hacer_login(driver, wait)
    driver.get("https://erp.com/som/salesreport")
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[2]/span"))).click()
    time.sleep(3)
    sucursales_els = driver.find_elements(By.XPATH, "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[3]/div[2]/ul/p-dropdownitem")
    cantidad_sucursales = len(sucursales_els)
    print(f"Total de sucursales detectadas: {cantidad_sucursales}")

    fecha_consulta = fecha_inicio
    while fecha_consulta <= fecha_fin:
        f_consulta_str = fecha_consulta.strftime('%d-%m-%Y')
        print(f"\n🚀 PROCESANDO FECHA: {f_consulta_str}")

        j = 1
        while j <= cantidad_sucursales:
            try:
                verificar_y_reloguear(driver, wait)
                print(f"\n--- [{f_consulta_str}] Sucursal {j}/{cantidad_sucursales} ---")
                driver.get("https://erp.com/som/salesreport")
                
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[2]/span"))).click()
                xpath_sucursal = f"/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[3]/div[2]/ul/p-dropdownitem[{j}]/li/span[1]"
                sucursal_el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_sucursal)))
                nombre_sucursal = sucursal_el.text
                
                # Excluir sucursales no deseadas
                if "CENDI " in nombre_sucursal or "CORPORATIVO CCS" in nombre_sucursal:
                    print(f"⏭️ Sucursal '{nombre_sucursal}' excluida. Saltando...")
                    driver.find_element(By.TAG_NAME, 'body').click()
                    j += 1
                    continue
                
                sucursal_el.click()
                time.sleep(1)

                # Selección de fechas en los calendarios (igual que antes)
                for div_id in ["11", "12"]:
                    wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{div_id}]/p-calendar/span/button/span[1]"))).click()
                    time.sleep(0.5)
                    wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{div_id}]/p-calendar/span/div/div/div/div[1]/div/button[2]"))).click()
                    idx_anio = fecha_consulta.year - 2019
                    wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{div_id}]/p-calendar/span/div/div[2]/span[{idx_anio}]"))).click()
                    idx_mes = fecha_consulta.month
                    wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{div_id}]/p-calendar/span/div/div[2]/span[{idx_mes}]"))).click()
                    tbody_xpath = f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{div_id}]/p-calendar/span/div/div/div/div[2]/table/tbody"
                    dia_str = str(fecha_consulta.day)
                    wait.until(EC.element_to_be_clickable((By.XPATH, f"{tbody_xpath}//span[text()='{dia_str}'] | {tbody_xpath}//a[text()='{dia_str}']"))).click()

                driver.find_element(By.XPATH, "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[15]/button[1]/span[2]").click()

                time.sleep(50)
                spinner_xpath = "/html/body/app-root/app-layout/app-loading/div/div/div/p-progressspinner/div/svg"
                try:
                    time.sleep(5)
                    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, spinner_xpath)))
                except: pass
                WebDriverWait(driver, 180).until(EC.invisibility_of_element_located((By.XPATH, spinner_xpath)))

                tabla_xpath = "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[2]/p-table/div/div/table/tbody"
                tabla_presente = False
                try:
                    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, tabla_xpath)))
                    tabla_presente = True
                except: pass

                valor_web = 0.0
                if tabla_presente:
                    time.sleep(3)
                    valor_web_raw = driver.find_element(By.XPATH, "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[2]/div[2]/span[5]").text
                    valor_web = float(valor_web_raw.replace('$', '').replace('.', '').replace(',', '.').strip())

                # Consultar DB
                conn = psycopg2.connect(DB_URL)
                cur = conn.cursor()
                cur.execute("SELECT SUM(total_ventas_usd), COUNT(*) FROM public.ventas WHERE sucursal = %s AND fecha = %s;", (nombre_sucursal, fecha_consulta))
                res_db, conteo_filas = cur.fetchone()
                valor_db = float(res_db) if res_db else 0.0
                cur.close()
                conn.close()

                # --- LÓGICA DE DETECCIÓN DE DIFERENCIAS ---
                es_correcto = (round(valor_web, 2) == round(valor_db, 2) and valor_web > 0) or (not tabla_presente and valor_web == 0 and (conteo_filas or 0) == 0)
                
                if es_correcto:
                    print(f"✅ Correcto para {nombre_sucursal}.")
                    j += 1
                    continue

                # --- ES UNA DISCREPANCIA: REGISTRAR EN EL REPORTE ---
                diferencia = valor_web - valor_db
                diferencias_report.append({
                    "fecha": fecha_consulta,
                    "sucursal": nombre_sucursal,
                    "total_web": valor_web,
                    "total_db": valor_db,
                    "diferencia": diferencia
                })
                print(f"📝 Diferencia registrada: Web={valor_web}, DB={valor_db}, Dif={diferencia}")

                # --- PROCESO DE DESCARGA DEL EXCEL (igual que antes) ---
                if (valor_web != 0 and (conteo_filas or 0) == 0):
                    print(f"⚠️ Web reporta {valor_web} pero DB tiene 0 registros. Verificando con Excel...")
                else:
                    print(f"⚠️ Discrepancia: Web {valor_web} vs DB {valor_db}. Verificando con Excel...")
                
                # Limpiar carpeta de descargas
                for f in os.listdir(DOWNLOAD_DIR):
                    if f.endswith(".xlsx"):
                        try: os.remove(os.path.join(DOWNLOAD_DIR, f))
                        except: pass

                driver.find_element(By.XPATH, "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[15]/button[3]/span[2]").click()
                
                archivo = esperar_descarga_estricta(DOWNLOAD_DIR)
                if archivo:
                    nuevo_nombre = f"{nombre_sucursal}+{f_consulta_str}.xlsx".replace("/", "-")
                    ruta_nueva = os.path.join(DOWNLOAD_DIR, nuevo_nombre)
                    shutil.move(archivo, ruta_nueva)
                    
                    try:
                        df_excel = pd.read_excel(ruta_nueva)
                        sum_excel = df_excel["Total ventas USD"].sum()
                        
                        if round(sum_excel, 2) == round(valor_web, 2):
                            shutil.move(ruta_nueva, os.path.join(TARGET_DIR, nuevo_nombre))
                            print(f"📄 Archivo '{nuevo_nombre}' verificado y movido a Ventas.")
                            j += 1
                            
                        else:
                            print(f"❌ Error: Suma Excel ({sum_excel}) != Web ({valor_web}). Reintentando...")
                            if os.path.exists(ruta_nueva): os.remove(ruta_nueva)
                            driver.refresh()
                        
                        time.sleep(1)
                        driver.get("https://erp.com/som/salesreport") # ------ en tal caso quitar
                        
                    except Exception as e_pandas:
                        print(f"❌ Error leyendo Excel: {e_pandas}. Reintentando...")
                        driver.refresh()
                else:
                    print("❌ El archivo no se descargó. Reintentando sucursal...")
                    driver.refresh()

            except Exception as e:
                print(f"💥 Volviendo al reporte de Ventas")
                time.sleep(1)
                driver.refresh()

        fecha_consulta += timedelta(days=1)

    print("\n✅ PROCESO FINALIZADO.")

finally:
    driver.quit()

# --- GENERACIÓN DEL REPORTE EXCEL CON LAS DIFERENCIAS ---
if diferencias_report:
    df_reporte = pd.DataFrame(diferencias_report)
    
    # Convertir la columna 'fecha' a datetime (corrige el error)
    df_reporte["fecha"] = pd.to_datetime(df_reporte["fecha"])
    
    # Ordenar por fecha y luego por sucursal
    df_reporte = df_reporte.sort_values(by=["fecha", "sucursal"])
    # Formatear la columna fecha para que se vea legible
    df_reporte["fecha"] = df_reporte["fecha"].dt.strftime("%Y-%m-%d")
    
    # Carpeta para informes
    REPORT_DIR = r"Ruta de la carpeta de los informes"
    if not os.path.exists(REPORT_DIR):
        os.makedirs(REPORT_DIR)
    
    ruta_reporte = os.path.join(REPORT_DIR, "reporte_diferencias.xlsx")
    df_reporte.to_excel(ruta_reporte, index=False, sheet_name="Diferencias")
    print(f"\n📊 REPORTE GENERADO: {ruta_reporte}")
    print("Contenido del reporte:")
    print(df_reporte.to_string(index=False))
else:
    print("\n🎉 No se encontraron diferencias en ningún día/sucursal. No se genera reporte.")