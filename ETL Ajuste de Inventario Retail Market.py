import os
import time
import shutil
import pandas as pd
import re
import urllib3
from datetime import datetime, date, timedelta
from sqlalchemy import create_engine, text
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- SOLUCIÓN A LAS ADVERTENCIAS DE SSL ---
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
os.environ['WDM_SSL_VERIFY'] = '0'

# --- CONFIGURACIÓN ---
USER = "luis.cornieles@mail.com"
PASSWORD = "xxxxx"
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
DEST_DIR = r"C:\Users\Ajustes"
DB_URL = 'postgresql://xxx:xxxxx5@host:xxx/RetailMARKET'
TABLE_NAME = 'ajustes_inventario'

if not os.path.exists(DEST_DIR):
    os.makedirs(DEST_DIR)

XPATHS = {
    "user": "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[1]/input",
    "pass": "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[2]/input",
    "login_btn": "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[4]/button/span",
    "fecha_ini": "(//kendo-dateinput//input)[1]",
    "fecha_fin": "(//kendo-dateinput//input)[2]",
    "btn_buscar": "//button[contains(., 'Buscar')]",
    "paginacion": "/html/body/app-root/app-layout/div/div/div[1]/app-report-main/div/app-inventory-count-report/app-grid/kendo-grid/kendo-pager/kendo-pager-numeric-buttons/ul/li[1]/span",
    "btn_exportar": "/html/body/app-root/app-layout/div/div/div[1]/app-report-main/div/app-inventory-count-report/app-grid/kendo-grid/kendo-grid-toolbar/button[1]/span[2]",
    "no_data": "//td[contains(text(), 'No hay datos') or contains(text(), 'No records')]",
    "error_toast": "/p-toastitem/div/div/div"  # XPATH DE ERROR PROPORCIONADO
}

STEPS_COLUMNS = [
    "/html/body/app-root/app-layout/div/div/div[1]/app-report-main/div/app-inventory-count-report/app-grid/kendo-grid/div/div/div/table/thead/tr[2]/th[1]/span[1]/kendo-grid-column-menu/a/span",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[1]/span",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[5]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[12]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[13]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[14]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[15]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[16]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[17]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[18]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[19]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[20]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[21]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[1]/label[22]/input",
    "/html/body/app-root/kendo-popup/div/kendo-grid-columnmenu-container/kendo-grid-columnmenu-chooser/kendo-grid-columnmenu-item/div[2]/kendo-grid-columnlist/div[2]/button[2]" 
]

# --- FUNCIONES ---

def get_dates_range(mode="last_month"):
    today = date.today()
    if mode == "last_month":
        last_prev = today.replace(day=1) - timedelta(days=1)
        first_prev = last_prev.replace(day=1)
        return first_prev.strftime("%d%m%Y"), last_prev.strftime("%d%m%Y")
    return today.replace(day=1).strftime("%d%m%Y"), today.strftime("%d%m%Y")

def safe_fill_input(driver, xpath, value, send_tab=False):
    for i in range(3):
        try:
            element = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(1)
            element.click()
            element.send_keys(Keys.CONTROL + "a", Keys.DELETE)
            element.send_keys(value)
            if send_tab: element.send_keys(Keys.TAB)
            return
        except:
            time.sleep(2)
            if i == 2: raise

def run_report_flow(driver, f_ini, f_fin):
    # Usamos un bucle de reintento interno por si aparece el Toast de error
    max_retries = 5
    for attempt in range(max_retries):
        try:
            driver.refresh() # F5 al iniciar para limpiar estado previo
            WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, XPATHS["fecha_ini"])))
            
            safe_fill_input(driver, XPATHS["fecha_ini"], f_ini)
            safe_fill_input(driver, XPATHS["fecha_fin"], f_fin, send_tab=True)
            time.sleep(1)
            
            for step in STEPS_COLUMNS:
                btn = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, step)))
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(0.4)
            
            btn_buscar = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS["btn_buscar"])))
            driver.execute_script("arguments[0].click();", btn_buscar)

            # --- LÓGICA DE DETECCIÓN DE ERROR ---
            # Esperamos a que aparezca o la paginación, o "No hay datos", O el error Toast
            result = WebDriverWait(driver, 180).until(EC.any_of(
                EC.visibility_of_element_located((By.XPATH, XPATHS["paginacion"])),
                EC.visibility_of_element_located((By.XPATH, XPATHS["no_data"])),
                EC.visibility_of_element_located((By.XPATH, XPATHS["error_toast"]))
            ))

            # Si lo que se visualizó es el Toast de error, lanzamos excepción para reintentar
            if driver.find_elements(By.XPATH, XPATHS["error_toast"]):
                print(f"⚠️ Error detectado (Toast). Reintentando proceso ({attempt + 1}/{max_retries})...")
                continue # Esto vuelve al inicio del for y hace refresh (F5)

            if driver.find_elements(By.XPATH, XPATHS["no_data"]): 
                return None

            # Proceso de descarga normal
            files_before = set(os.listdir(DOWNLOAD_DIR))
            btn_export = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, XPATHS["btn_exportar"])))
            driver.execute_script("arguments[0].click();", btn_export)
            
            start_t = time.time()
            while time.time() - start_t < 120:
                diff = set(os.listdir(DOWNLOAD_DIR)) - files_before
                if diff:
                    fname = list(diff)[0]
                    if not fname.endswith(('.tmp', '.crdownload')): 
                        return os.path.join(DOWNLOAD_DIR, fname)
                time.sleep(2)
            
            return None # Si sale del while por tiempo

        except Exception as e:
            print(f"Aviso en run_report_flow: {e}. Reintentando...")
            time.sleep(2)
    
    return None

def process_and_move_file(downloaded_file):
    try:
        df = pd.read_excel(downloaded_file, header=1)
        col_name = "Fecha del ajuste"
        df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
        df_clean = df.dropna(subset=[col_name])
        if not df_clean.empty:
            f_min_str = df_clean[col_name].min().strftime("%d%m%Y")
            f_max_str = df_clean[col_name].max().strftime("%d%m%Y")
            f_final_format = df_clean[col_name].max().strftime("%d-%m-%Y")
            nuevo_nombre = f"Reporte-Ajuste-Inventario+{f_min_str}-{f_max_str}_{f_final_format}.xlsx"
            final_path = os.path.join(DEST_DIR, nuevo_nombre)
            if os.path.exists(final_path): os.remove(final_path)
            shutil.move(downloaded_file, final_path)
            print(f"Descargado y procesado: {nuevo_nombre}")
    except Exception as e: print(f"Error procesando archivo: {e}")

def main():
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        driver.get("https://erp.rm.com/")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, XPATHS["user"]))).send_keys(USER)
        driver.find_element(By.XPATH, XPATHS["pass"]).send_keys(PASSWORD)
        driver.find_element(By.XPATH, XPATHS["login_btn"]).click()
        time.sleep(5)
        
        for modo in ["last_month", "current_month"]:
            driver.get("https://erp.rm.com/report/inventory-adjustment")
            f_ini, f_fin = get_dates_range(modo)
            archivo = run_report_flow(driver, f_ini, f_fin)
            if archivo: 
                process_and_move_file(archivo)

        # ... (Resto del código de limpieza y carga a DB se mantiene igual)
        print("Iniciando limpieza de archivos temporales locales...")
        today = date.today()
        yesterday_str = (today - timedelta(days=1)).strftime("%d-%m-%Y")
        
        for f in os.listdir(DEST_DIR):
            if "_" in f and f.endswith(".xlsx"):
                try:
                    fecha_en_nombre = f.split("_")[1].split(".")[0]
                    if fecha_en_nombre == yesterday_str:
                        path_f = os.path.join(DEST_DIR, f)
                        df_c = pd.read_excel(path_f, header=1)
                        df_c["Fecha del ajuste"] = pd.to_datetime(df_c["Fecha del ajuste"], errors='coerce')
                        df_c = df_c.dropna(subset=["Fecha del ajuste"])
                        if not df_c.empty:
                            max_date_in_file = df_c["Fecha del ajuste"].max()
                            if max_date_in_file.day != 1 and max_date_in_file.month == today.month:
                                os.remove(path_f)
                                print(f"ELIMINADO: {f}")
                except Exception as e: 
                    print(f"Aviso: Error evaluando {f}: {e}")
    finally:
        driver.quit()

    # --- 3. CARGA A BASE DE DATOS ---
    engine = create_engine(DB_URL, pool_pre_ping=True)
    
    try:
        engine.dispose()
        print("Iniciando limpieza de tabla en PostgreSQL...")
        with engine.connect() as conn:
            with conn.begin(): 
                conn.execute(text(f"TRUNCATE TABLE {TABLE_NAME};"))
            print(f"✓ Tabla {TABLE_NAME} vaciada con éxito.")
            
    except Exception as e:
        print(f"❌ Error crítico en TRUNCATE: {e}")
        return

    archivos_finales = [f for f in os.listdir(DEST_DIR) if f.endswith('.xlsx')]
    lista_dfs = []
    
    mapping = {
        'Sucursal': 'sucursal', 'Área': 'area', 'Barra': 'barra', 'Nombre del producto': 'nombre_producto',
        'Categoría': 'categoria', 'Número de ajuste': 'numero_ajuste', 'Estatus': 'estatus',
        'Tipo de ajuste': 'tipo_ajuste', 'Motivo de ajuste': 'motivo_ajuste', 'Motivo de agrupación': 'motivo_agrupacion',
        'Operador': 'operador', 'Responsable': 'responsable', 'Fecha del ajuste': 'fecha_ajuste',
        'Existencia conteo': 'existencia_conteo', 'Conteo': 'conteo', 'Entradas': 'entradas',
        'Salidas': 'salidas', 'Existencia': 'existencia', 'Costo': 'costo', 'PVP': 'pvp',
        'Costo conversion': 'costo_conversion', 'PVP conversion': 'pvp_conversion'
    }

    print(f"Preparando {len(archivos_finales)} archivos para la carga...")
    for archivo in archivos_finales:
        try:
            path_file = os.path.join(DEST_DIR, archivo)
            match = re.search(r'_(\d{2}-\d{2}-\d{4})', archivo)
            f_arch_dt = pd.to_datetime(match.group(1), dayfirst=True).date() if match else None
            
            df_t = pd.read_excel(path_file, header=1, dtype={'Barra': str, 'Número de ajuste': str})
            df_t = df_t.rename(columns=mapping)
            
            for col in ['barra', 'numero_ajuste']:
                if col in df_t.columns:
                    df_t[col] = df_t[col].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', None)
            
            df_t['fecha_ajuste'] = pd.to_datetime(df_t['fecha_ajuste'], errors='coerce').dt.date
            df_t['fecha_archivo_dt'] = f_arch_dt
            df_t = df_t.dropna(subset=['fecha_ajuste'])
            lista_dfs.append(df_t)
        except Exception as e: 
            print(f"Aviso: Error procesando {archivo}: {e}")

    if lista_dfs:
        try:
            df_final = pd.concat(lista_dfs, ignore_index=True)
            with engine.connect() as conn:
                df_final.to_sql(TABLE_NAME, conn, if_exists='append', index=False)
                conn.commit()
            print(">>> ✅ PROCESO COMPLETADO EXITOSAMENTE. Datos actualizados en RIOMARKET.")
        except Exception as e:
            print(f"❌ Error durante la inserción de datos: {e}")
        finally:
            engine.dispose()

if __name__ == "__main__":
    main()