import os
import time
import logging
import sys
import pandas as pd
import numpy as np
import glob
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchWindowException
from sqlalchemy import create_engine, text
import psycopg2
from requests.exceptions import ConnectionError
from dotenv import load_dotenv

# --- CARGAR VARIABLES DE ENTORNO CON SOPORTE PARA MÚLTIPLES CODIFICACIONES ---
env_path = r"C:\Users\Python\Config\.env"

def load_env_with_encoding(env_path):
    """Intenta cargar el archivo .env con diferentes codificaciones."""
    for encoding in ['utf-8', 'latin-1', 'cp1252']:
        try:
            with open(env_path, 'r', encoding=encoding) as f:
                load_dotenv(stream=f)
            return True
        except UnicodeDecodeError:
            continue
    return False

if not load_env_with_encoding(env_path):
    print(f"No se pudo cargar el archivo .env en {env_path} con ninguna codificación.")
    sys.exit(1)

# --- CONFIGURACIÓN DE LOGGING ---
LOG_DIR = r"C:\Users\adminrpa\Documents\Python\Logs"
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

log_file = os.path.join(LOG_DIR, "automatizacion_ventas.log")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# --- VARIABLES DE ENTORNO OBLIGATORIAS ---
required_env_vars = [
    "DB_PASSWORD",
    "OWA_USER",
    "OWA_PASSWORD",
    "DESTINATARIOS",
    "DESTINATARIOS_ERROR",
    "ERP_USER",
    "ERP_PASSWORD"
]

missing_vars = []
for var in required_env_vars:
    if not os.getenv(var):
        missing_vars.append(var)

if missing_vars:
    logging.critical(f"Faltan variables de entorno obligatorias: {', '.join(missing_vars)}")
    sys.exit(1)

# --- CONFIGURACIÓN DE RUTAS Y NAVEGADOR ---
download_path = r"C:\Users\adminrpa\Documents\Reportes\Ventas"
if not os.path.exists(download_path):
    os.makedirs(download_path)

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
# === ELIMINA EL MENSAJE "CONTROLADO POR SOFTWARE DE PRUEBA AUTOMATIZADO" ===
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)
# ========================================================================
prefs = {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "directory_upgrade": True
}
chrome_options.add_experimental_option("prefs", prefs)

# --- CONFIGURACIÓN DE BASE DE DATOS (con valores por defecto para opcionales) ---
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_NAME = os.getenv("DB_NAME", "Retail_MARKET")
DB_USER = os.getenv("DB_USER", "XXXX")
DB_PASSWORD = os.getenv("DB_PASSWORD")  # Ya validado que no es None
DB_PORT = os.getenv("DB_PORT", "xxxx")

# Archivo de log para archivos ya procesados
LOG_PATH = r"C:\Users\archivos_cargados_Ventas.txt"

# Carpeta donde se guardará el reporte de control
report_folder = r"C:\Users\Reportes\Informes"
if not os.path.exists(report_folder):
    os.makedirs(report_folder)

# --- FUNCIONES DE APOYO ---

def esperar_descarga(folder, archivos_previos, timeout=120):
    """
    Espera de forma estricta a que aparezca un archivo nuevo 
    que no estaba en la lista de archivos_previos.
    """
    seconds = 0
    while seconds < timeout:
        archivos_actuales = set(os.listdir(folder))
        nuevos_archivos = archivos_actuales - set(archivos_previos)
        completados = [f for f in nuevos_archivos if f.endswith(('.xlsx', '.xls')) and not f.endswith('.crdownload')]
        if completados:
            return os.path.join(folder, completados[0])
        time.sleep(1)
        seconds += 1
    return None


def limpiar_numero(valor):
    """
    Convierte una cadena con formato numérico (español o inglés) a float.
    Si no puede convertir, devuelve NaN.
    """
    if pd.isna(valor):
        return np.nan
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if texto == '':
        return np.nan
    # Quitar símbolo $ si existe
    texto = texto.replace('$', '').strip()
    # Aplicar la misma lógica de parsear_valor_web
    if ',' in texto and '.' in texto:
        if texto.rfind(',') > texto.rfind('.'):
            texto = texto.replace('.', '').replace(',', '.')
        else:
            texto = texto.replace(',', '')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    # Si solo hay punto, ya está bien
    try:
        return float(texto)
    except ValueError:
        return np.nan


def cargar_archivos_a_db():
    """
    Recorre la carpeta de descargas, identifica archivos con formato
    'Sucursal+DD-MM-YYYY.xlsx', los carga en la tabla ventas de PostgreSQL
    con los tipos de datos correctos y actualiza el archivo de log.
    """
    # Crear conexión a la base de datos
    engine = create_engine(f'postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}')
    
    # Leer archivos ya procesados
    archivos_procesados = set()
    if os.path.exists(LOG_PATH):
        with open(LOG_PATH, 'r', encoding='utf-8') as f:
            archivos_procesados = set(line.strip() for line in f)
    
    # Mapeo de columnas del Excel a columnas de la tabla
    column_mapping = {
        'Fecha Desde': 'fecha_desde',
        'Fecha Hasta': 'fecha_hasta',
        'Referencia': 'referencia',
        'Barra': 'barra',
        'Producto': 'producto',
        'Marca': 'marca',
        'Categoría': 'categoria',
        'Proveedor': 'proveedor',
        'Procedencia': 'procedencia',
        'Empaque': 'empaque',
        'Costo neto': 'costo_neto',
        'Costo neto USD': 'costo_neto_usd',
        'Costo neto venta': 'costo_neto_venta',
        'Costo neto venta USD': 'costo_neto_venta_usd',
        'PVP': 'pvp',
        'PVP USD': 'pvp_usd',
        'Total costo': 'total_costo',
        'Total costo USD': 'total_costo_usd',
        'Total costo neto venta': 'total_costo_neto_venta',
        'Total costo neto venta USD': 'total_costo_neto_venta_usd',
        'Cantidad vendida': 'cantidad_vendida',
        'Base bruta': 'base_bruta',
        'Base bruta USD': 'base_bruta_usd',
        'Descuento': 'descuento',
        'Descuento USD': 'descuento_usd',
        'Base neta': 'base_neta',
        'Base neta USD': 'base_neta_usd',
        'IVA': 'iva',
        'IVA USD': 'iva_usd',
        'Total ventas': 'total_ventas',
        'Total ventas USD': 'total_ventas_usd',
        'MSV': 'msv'
    }

    # Lista de columnas que deben ser numéricas (double precision)
    columnas_float = [
        'costo_neto', 'costo_neto_usd', 'costo_neto_venta', 'costo_neto_venta_usd',
        'pvp', 'pvp_usd', 'total_costo', 'total_costo_usd', 'total_costo_neto_venta',
        'total_costo_neto_venta_usd', 'cantidad_vendida', 'base_bruta', 'base_bruta_usd',
        'base_neta', 'base_neta_usd', 'total_ventas', 'total_ventas_usd', 'msv',
        'descuento', 'descuento_usd', 'iva', 'iva_usd'
    ]

    # Recorrer archivos en carpeta de descargas
    for archivo in os.listdir(download_path):
        if archivo.endswith('.xlsx') and '+' in archivo and archivo.count('-') >= 2:
            if archivo in archivos_procesados:
                logging.info(f"Archivo ya procesado anteriormente: {archivo}")
                continue

            # Extraer sucursal y fecha del nombre
            nombre_sucursal = archivo.split('+')[0]
            fecha_str = archivo.split('+')[1].replace('.xlsx', '')
            try:
                fecha_consulta = datetime.strptime(fecha_str, "%d-%m-%Y").date()
            except ValueError:
                logging.error(f"Formato de fecha no válido en archivo: {archivo}")
                continue

            file_path = os.path.join(download_path, archivo)
            try:
                # Leer Excel como texto para evitar conversiones automáticas
                df = pd.read_excel(file_path, dtype=str, keep_default_na=False)

                # Renombrar columnas según mapeo
                df.rename(columns=column_mapping, inplace=True)

                # --- Conversión de tipos ---
                # 1. Columnas de texto: convertir a string y reemplazar cadenas vacías por None
                for col in df.columns:
                    if col not in columnas_float:
                        df[col] = df[col].astype(str).replace('', None)

                # 2. Columnas numéricas (float): aplicar limpiar_numero
                for col in columnas_float:
                    if col in df.columns:
                        df[col] = df[col].apply(limpiar_numero)
                    else:
                        logging.warning(f"Columna numérica {col} no encontrada en el archivo {archivo}")

                # 3. Agregar columnas adicionales que no vienen en el Excel
                df['sucursal'] = nombre_sucursal
                df['fecha'] = pd.to_datetime(fecha_consulta)   # timestamp
                df['clave_producto_final'] = None
                df['barra_match'] = None
                df['producto_match'] = None

                # 4. Insertar en la tabla
                df.to_sql('ventas', engine, if_exists='append', index=False, schema='public')

                # Registrar archivo en log
                with open(LOG_PATH, 'a', encoding='utf-8') as f:
                    f.write(archivo + '\n')
                logging.info(f"Archivo cargado exitosamente: {archivo}")

            except Exception as e:
                logging.error(f"Error al procesar archivo {archivo}: {e}")


def verificar_bloqueo(driver):
    """Hace click en el botón de cerrar si aparece el popup especificado."""
    try:
        # XPath del popup que mencionaste
        popup_xpath = "/html/body/div[12]/div/div[2]/div[4]/div/div/span/button/span"
        popup = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, popup_xpath)))
        popup.click()
        logging.info("Popup de OWA cerrado correctamente.")
    except TimeoutException:
        pass


def enviar_correo_error(asunto, cuerpo):
    """
    Envía un correo de error usando OWA. Crea su propio driver para no depender del principal.
    """
    logging.info("INICIANDO ENVÍO DE CORREO DE ERROR")
    error_driver = None
    try:
        # Credenciales de OWA desde variables de entorno (obligatorias)
        owa_user = os.getenv("OWA_USER")
        owa_password = os.getenv("OWA_PASSWORD")
        destinatarios_error = os.getenv("DESTINATARIOS_ERROR")

        error_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        wait_err = WebDriverWait(error_driver, 20)

        # Login en correo
        error_driver.get("https://mail.retailmarket.com")
        wait_err.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div/div[2]/div/div[3]/input"))).send_keys(owa_user)
        error_driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div/div[5]/input").send_keys(owa_password)
        error_driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div/div[9]/div/span").click()

        verificar_bloqueo(error_driver)
        time.sleep(10)

        # Nuevo correo
        wait_err.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div[3]/div[5]/div/div[1]/div/div[5]/div[1]/div/div[1]/div/div/div[1]/div/button[1]/span[2]"))).click()
        verificar_bloqueo(error_driver)
        time.sleep(5)

        # Destinatario
        wait_err.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[3]/div[5]/div/div[1]/div/div[5]/div[3]/div/div[5]/div[1]/div/div[3]/div[4]/div/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/div[1]/div/div/div/span/div[1]/form/input"))).send_keys(destinatarios_error)

        # Asunto y cuerpo con ActionChains
        actions = ActionChains(error_driver)
        for _ in range(3):
            actions.send_keys(Keys.TAB)
        actions.perform()
        actions.send_keys(asunto)
        actions.send_keys(Keys.TAB)
        actions.send_keys(cuerpo)
        actions.perform()

        # Enviar
        xpath_enviar = "/html/body/div[2]/div/div[3]/div[5]/div/div[1]/div/div[5]/div[1]/div/div[1]/div/div/div[1]/div/button/span[2]"
        btn_enviar = wait_err.until(EC.element_to_be_clickable((By.XPATH, xpath_enviar)))
        btn_enviar.click()
        logging.info("CORREO DE ERROR ENVIADO")

    except Exception as e:
        logging.error(f"Error al enviar correo de error: {e}")
    finally:
        if error_driver:
            error_driver.quit()


def reiniciar_driver(driver_actual):
    """Cierra el driver actual y crea uno nuevo."""
    logging.info("Reiniciando WebDriver debido a error de ventana...")
    try:
        if driver_actual:
            driver_actual.quit()
    except:
        pass
    # Reintentar hasta 3 veces
    for intento in range(3):
        try:
            service = Service(ChromeDriverManager().install())
            nuevo_driver = webdriver.Chrome(service=service, options=chrome_options)
            nuevo_wait = WebDriverWait(nuevo_driver, 20)
            logging.info("WebDriver reiniciado correctamente.")
            return nuevo_driver, nuevo_wait
        except Exception as e:
            logging.error(f"Error al reiniciar WebDriver (intento {intento+1}): {e}")
            time.sleep(5)
    raise Exception("No se pudo reiniciar el WebDriver después de 3 intentos.")


def parsear_valor_web(texto):
    """
    Convierte una cadena con formato de número (español o inglés) a float.
    Ejemplos:
        "1.234,56 $" -> 1234.56
        "1,234.56 $" -> 1234.56
        "1234.56 $"   -> 1234.56
        "1234,56 $"   -> 1234.56
    """
    # Eliminar el símbolo $ y espacios al inicio/final
    limpio = texto.replace('$', '').strip()
    # Detectar si hay coma y punto
    if ',' in limpio and '.' in limpio:
        # Determinar cuál es el separador decimal
        if limpio.rfind(',') > limpio.rfind('.'):
            # Formato español: el último separador es coma (decimal)
            # Ej: 1.234,56
            limpio = limpio.replace('.', '')   # quitar puntos (miles)
            limpio = limpio.replace(',', '.')  # coma a punto decimal
        else:
            # Formato inglés: el último separador es punto (decimal)
            # Ej: 1,234.56
            limpio = limpio.replace(',', '')   # quitar comas (miles)
    elif ',' in limpio:
        # Solo coma: puede ser español (decimal) o inglés (miles) pero asumimos español
        limpio = limpio.replace(',', '.')
    elif '.' in limpio:
        # Solo punto: asumimos inglés (decimal)
        pass
    # Si no hay separadores, es un número entero
    return float(limpio)


# ============================================================================
# === NUEVO BLOQUE: INICIALIZACIÓN DEL DRIVER Y LOGIN CON REINTENTOS (180s) ===
# ============================================================================
max_intentos = 4
intento = 0
driver = None
wait = None
login_exitoso = False
ERP_USER = os.getenv("ERP_USER")
ERP_PASSWORD = os.getenv("ERP_PASSWORD")

while intento < max_intentos and not login_exitoso:
    intento += 1
    logging.info(f"Intento {intento} de {max_intentos} para iniciar sesión en ERP...")
    # Cerrar driver anterior si existe
    if driver:
        try:
            driver.quit()
        except:
            pass
    try:
        # Crear nuevo driver con las opciones anti-detección ya definidas
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        # Timeout de 180 segundos exclusivamente para la espera del elemento de login
        wait_largo = WebDriverWait(driver, 180)
        
        # Navegar a la URL del ERP
        driver.get("https://erp.rm.com/")
        
        # Esperar hasta 180 segundos a que aparezca el campo de usuario
        input_usuario = wait_largo.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[1]/input")))
        input_usuario.send_keys(ERP_USER)
        driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[2]/input").send_keys(ERP_PASSWORD)
        driver.find_element(By.XPATH, "/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[4]/button/span").click()
        
        # Si llegamos aquí, login exitoso
        login_exitoso = True
        logging.info("Login exitoso en ERP.")
        
    except TimeoutException:
        logging.error(f"Intento {intento}: No se encontró el elemento de login después de 180 segundos. Reintentando...")
        # El driver se cerrará al inicio del siguiente intento
    except Exception as e:
        logging.error(f"Intento {intento}: Error inesperado durante el login: {e}. Reintentando...")
        # Continuar reintentando

if not login_exitoso:
    logging.critical("No se pudo realizar el login después de 4 intentos. Enviando correo de error y abortando.")
    fecha_actual = datetime.now()
    fecha_str = fecha_actual.strftime("%d-%m-%Y")
    asunto = f"Fallo en automatización de Ventas - {fecha_str}"
    cuerpo = "Falló el flujo de venta por timeout en la carga de la página de login del ERP después de 4 intentos."
    enviar_correo_error(asunto, cuerpo)
    sys.exit(1)

# Restaurar el wait a 20 segundos para el resto del flujo (como estaba originalmente)
wait = WebDriverWait(driver, 20)

# ============================================================================
# === FIN DEL BLOQUE MODIFICADO. EL RESTO DEL CÓDIGO PERMANECE IGUAL ===
# ============================================================================

# --- INICIO DEL PROCESO PRINCIPAL ---
report_rows = []  # Lista para el reporte de control

try:
    # 2. Navegar a Sales Report inicial
    time.sleep(5)
    driver.get("https://erp.rm.com/som/salesreport")

    # 3. Detectar cantidad de sucursales
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[2]/span"))).click()
    xpath_lista = "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[3]/div[2]/ul"
    wait.until(EC.presence_of_element_located((By.XPATH, f"{xpath_lista}/p-dropdownitem")))
    cantidad_sucursales = len(driver.find_elements(By.XPATH, f"{xpath_lista}/p-dropdownitem"))
    logging.info(f"Total sucursales a procesar: {cantidad_sucursales}")

    # 9. Determinar fechas
    fecha_actual = datetime.now()
    fecha_consulta = fecha_actual - timedelta(days=1)
    fecha_str_file = fecha_consulta.strftime("%d-%m-%Y")
    logging.info(f"Fecha consulta: {fecha_str_file}")

    j = 1
    while j <= cantidad_sucursales:
        exito_sucursal = False
        intentos_sucursal = 0
        descarga_exitosa = False
        
        while not exito_sucursal and intentos_sucursal < 4:
            try:
                # 24. Regresar a la URL
                driver.get("https://erp.rm.com/som/salesreport")
                intentos_sucursal += 1

                # 6. Seleccionar sucursal j
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[2]/span"))).click()
                xpath_sucursal_j = f"{xpath_lista}/p-dropdownitem[{j}]/li/span[1]"
                elem_sucursal = wait.until(EC.presence_of_element_located((By.XPATH, xpath_sucursal_j)))
                time.sleep(2)
                nombre_sucursal = elem_sucursal.text.strip()
                logging.info(f"--- Procesando ({j}/{cantidad_sucursales}): {nombre_sucursal} | Intento: {intentos_sucursal} ---")

                # --- CONDICIÓN PARA OMITIR SUCURSALES CENDI ---
                if "CENDI " in nombre_sucursal or "CORPORATIVO CCS" in nombre_sucursal:
                    logging.info(f"--- Sucursal omitida por contener 'CENDI ' o 'CORPORATIVO CCS' : {nombre_sucursal} ---")
                    # Registrar en el reporte como omitida
                    report_rows.append({
                        'Fecha': fecha_str_file,
                        'Estado': 'Omitida (CENDI)',
                        'Sucursal': nombre_sucursal,
                        'Archivo': '',
                        'Registros Cargados': 0,
                        'Total Ventas USD': 0.0,
                        'Intentos': 0
                    })
                    j += 1
                    # Salir del bucle de intentos para esta sucursal y pasar a la siguiente
                    break
                # ------------------------------------------------

                elem_sucursal.click()

                # 7-8. Filtrar en el reporte
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[1]/p-dropdown/div/div[2]/span"))).click()
                input_busqueda = wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[1]/p-dropdown/div/div[3]/div[1]/div/input")))
                input_busqueda.send_keys(nombre_sucursal)
                time.sleep(1)
                
                opciones_filtro = driver.find_elements(By.XPATH, "//p-dropdownitem")
                for op in opciones_filtro:
                    if op.text.strip() == nombre_sucursal:
                        op.click()
                        break

                # 10-19. Configurar Calendarios
                for cal_id in ["11", "12"]:
                    wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{cal_id}]/p-calendar/span/button/span[1]"))).click()
                    wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{cal_id}]/p-calendar/span/div/div/div/div[1]/div/button[2]"))).click()
                    
                    anios = driver.find_elements(By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{cal_id}]/p-calendar/span/div/div[2]/span")
                    for a in anios:
                        if a.text.strip() == str(fecha_consulta.year):
                            a.click()
                            break
                    
                    meses = driver.find_elements(By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{cal_id}]/p-calendar/span/div/div[2]/span")
                    meses[fecha_consulta.month - 1].click()
                    
                    dias = driver.find_elements(By.XPATH, f"/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[{cal_id}]/p-calendar/span/div/div/div/div[2]/table/tbody//span")
                    for d in dias:
                        if d.text.strip() == str(fecha_consulta.day):
                            d.click()
                            break
                    time.sleep(1)

                # 20. Buscar
                driver.find_element(By.XPATH, "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[15]/button[1]/span[2]").click()

                # Esperar a que desaparezca el spinner
                wait.until(EC.invisibility_of_element_located((By.XPATH, "//div/div/div/p-progressspinner/div/svg")))

                # Definir los XPath que se usarán
                xpath_tabla = "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[2]/p-table/div/div/table/tbody"
                xpath_total_ventas = "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[2]/div[2]/span[5]"

                # Monitorear el valor del total de ventas hasta que sea distinto de cero o pasen 90 segundos
                start_time = time.time()
                timeout = 90
                valor_anterior = None
                while time.time() - start_time < timeout:
                    try:
                        valor_actual = driver.find_element(By.XPATH, xpath_total_ventas).text.strip()
                        if valor_actual not in ["0,00 $", "0.00 $"]:
                            # Ya hay un valor distinto de cero, podemos continuar
                            logging.info(f"Valor web detectado antes de timeout: {valor_actual}")
                            break
                        if valor_actual != valor_anterior:
                            logging.info(f"Valor web actual (cero): {valor_actual}")
                            valor_anterior = valor_actual
                    except:
                        # Si por alguna razón no se puede obtener el elemento, esperamos
                        pass
                    time.sleep(1)
                else:
                    logging.info("Timeout de 90 segundos alcanzado, se procede con el valor actual.")

                # Ahora procedemos a verificar visibilidad de tabla y obtener valor
                try:
                    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, xpath_tabla)))
                    tabla_visible = True
                except TimeoutException:
                    tabla_visible = False

                valor_web_raw = driver.find_element(By.XPATH, xpath_total_ventas).text.strip()
                # Si la tabla no es visible y el valor es 0, asumimos que no hay datos y reintentamos (hasta 4 intentos)
                if not tabla_visible and valor_web_raw == "0.00 $":
                    logging.info(f"Sucursal sin datos visibles y valor 0.00 $. Intento {intentos_sucursal}/4")
                    if intentos_sucursal >= 4:
                        exito_sucursal = True
                        j += 1
                    continue

                # 21. Procesar valor capturado (ahora con detección de formato)
                valor_web = parsear_valor_web(valor_web_raw)
                logging.info(f"Valor Web: {valor_web}")

                # --- 22. DESCARGA ESTRICTA ---
                archivos_antes_de_clic = os.listdir(download_path)
                
                driver.find_element(By.XPATH, "/html/body/app-root/app-layout/div/div/div[1]/app-salesreport-list/div[1]/app-salesreport-filter/div/div[15]/button[3]/span[2]").click()
                
                path_descargado = esperar_descarga(download_path, archivos_antes_de_clic)

                if path_descargado:
                    df = pd.read_excel(path_descargado)
                    if df["Total ventas USD"].dtype == object:
                        df["Total ventas USD"] = df["Total ventas USD"].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
                    
                    suma_excel = df["Total ventas USD"].sum()
                    logging.info(f"Suma Excel: {round(suma_excel, 2)}")

                    if abs(suma_excel - valor_web) < 0.05:
                        nuevo_nombre = os.path.join(download_path, f"{nombre_sucursal}+{fecha_str_file}.xlsx")
                        if os.path.exists(nuevo_nombre):
                            os.remove(nuevo_nombre)
                        os.rename(path_descargado, nuevo_nombre)
                        logging.info("Resultado: COINCIDE. Archivo guardado.")
                        descarga_exitosa = True
                        exito_sucursal = True
                        j += 1

                        # Guardar información para el reporte
                        report_rows.append({
                            'Fecha': fecha_str_file,
                            'Estado': 'Cargado',
                            'Sucursal': nombre_sucursal,
                            'Archivo': os.path.basename(nuevo_nombre),
                            'Registros Cargados': df.shape[0],
                            'Total Ventas USD': suma_excel,
                            'Intentos': intentos_sucursal
                        })
                    else:
                        logging.warning("Resultado: NO COINCIDE. Reintentando sucursal...")
                        if os.path.exists(path_descargado):
                            os.remove(path_descargado)
                        time.sleep(2)
                else:
                    logging.error("Error: No se detectó un archivo nuevo después de descargar.")

            except Exception as e:
                error_msg = str(e)
                # Si el error es de ventana cerrada, intentar reiniciar el driver y reintentar la misma sucursal
                if "target window already closed" in error_msg or "no such window" in error_msg:
                    logging.error(f"Error de ventana cerrada detectado: {error_msg}. Reiniciando driver...")
                    try:
                        driver, wait = reiniciar_driver(driver)
                        # No incrementamos j ni intentos_sucursal, simplemente reintentamos la misma sucursal
                        # (el bucle while continuará con el mismo j)
                        continue
                    except Exception as reinicio_error:
                        logging.error(f"Error al reiniciar driver: {reinicio_error}. Abortando.")
                        raise
                else:
                    logging.error(f"Error en flujo de sucursal {j}: {e}")
                    time.sleep(3)

        # Registro de sucursales sin datos o fallidas (no aplica para las omitidas, ya registradas)
        if exito_sucursal and not descarga_exitosa:
            report_rows.append({
                'Fecha': fecha_str_file,
                'Estado': 'Sin datos',
                'Sucursal': nombre_sucursal,
                'Archivo': '',
                'Registros Cargados': 0,
                'Total Ventas USD': 0.0,
                'Intentos': intentos_sucursal
            })
        elif not exito_sucursal and "CENDI " not in nombre_sucursal and "CORPORATIVO CCS" not in nombre_sucursal:
            report_rows.append({
                'Fecha': fecha_str_file,
                'Estado': 'Fallido',
                'Sucursal': nombre_sucursal,
                'Archivo': '',
                'Registros Cargados': 0,
                'Total Ventas USD': 0.0,
                'Intentos': intentos_sucursal
            })

        # --- NUEVO: AVANZAR AL SIGUIENTE j SI NO SE PUDO PROCESAR LA SUCURSAL ---
        if not exito_sucursal and "CENDI " not in nombre_sucursal and "CORPORATIVO CCS" not in nombre_sucursal:
            j += 1

    logging.info("DESCARGA FINALIZADA. INICIANDO CARGA A BASE DE DATOS")
    cargar_archivos_a_db()
    logging.info("CARGA COMPLETADA")

    # --- CREACIÓN DEL REPORTE EXCEL DE CONTROL ---
    logging.info("GENERANDO REPORTE DE CONTROL")
    report_filename = f"Control_ventas+{fecha_str_file}.xlsx"
    report_path = os.path.join(report_folder, report_filename)

    df_reporte = pd.DataFrame(report_rows)
    df_reporte.to_excel(report_path, index=False)
    logging.info(f"Reporte guardado en: {report_path}")

    # --- ENVÍO DEL REPORTE POR CORREO (OUTLOOK WEB APP) ---
    logging.info("INICIANDO ENVÍO DE CORREO")

    # Credenciales de OWA para el envío normal (desde variables de entorno)
    owa_user = os.getenv("OWA_USER")
    owa_password = os.getenv("OWA_PASSWORD")
    destinatarios = os.getenv("DESTINATARIOS")

    # Login en OWA
    driver.get("https://mail.rm.com/")
    wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div/div[2]/div/div[3]/input"))).send_keys(owa_user)
    driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div/div[5]/input").send_keys(owa_password)
    driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div/div[9]/div/span").click()

    # Verificar si aparece popup después del login
    verificar_bloqueo(driver)

    # Esperar a que cargue el buzón
    time.sleep(10)

    # Hacer clic en "Nuevo correo"
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div[3]/div[5]/div/div[1]/div/div[5]/div[1]/div/div[1]/div/div/div[1]/div/button[1]/span[2]"))).click()

    # Verificar si aparece popup al abrir nuevo mensaje
    verificar_bloqueo(driver)

    # Esperar a que aparezca la ventana de redacción
    time.sleep(5)

    # Destinatario
    wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[3]/div[5]/div/div[1]/div/div[5]/div[3]/div/div[5]/div[1]/div/div[3]/div[4]/div/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/div[1]/div/div/div/span/div[1]/form/input"))).send_keys(destinatarios)
    time.sleep(4)

    # Usar ActionChains para navegar al asunto y cuerpo
    actions = ActionChains(driver)
    # Presionar 3 veces Tab para llegar al Asunto
    for _ in range(3):
        actions.send_keys(Keys.TAB)
    actions.perform()
    
    time.sleep(4)

    # Escribir Asunto
    actions.send_keys("Reporte de Descarga-Carga de Ventas a Base de Datos Retailmarket - Tabla Ventas")
    time.sleep(4)
    
    # Presionar 1 vez Tab para llegar al Cuerpo
    actions.send_keys(Keys.TAB)
    time.sleep(4)
    # Escribir Cuerpo del mensaje
    fecha_cuerpo = fecha_consulta.strftime("%d/%m/%Y")
    actions.send_keys(f"Estadística de la descarga-carga del reporte de ventas del día {fecha_cuerpo}")
    actions.perform()
    time.sleep(4)

    # Adjuntar archivo
    input_file = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']")))
    input_file.send_keys(report_path)
    time.sleep(5)  # Esperar que se complete la carga

    # Hacer clic en "Enviar"
    xpath_enviar = "/html/body/div[2]/div/div[3]/div[5]/div/div[1]/div/div[5]/div[1]/div/div[1]/div/div/div[1]/div/button/span[2]"
    btn_enviar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_enviar)))
    btn_enviar.click()
    logging.info("CORREO ENVIADO")

except Exception as e:
    logging.error(f"Error crítico: {e}")

finally:
    if driver:
        time.sleep(5)
        driver.quit()