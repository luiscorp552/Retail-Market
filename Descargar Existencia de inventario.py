import os
import time
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ================= CONFIGURACIÓN =================
USER = "luis.cornieles@riomarket.com"
PASSWORD = "lkdrc291"
LOGIN_URL = "https://erp.rt.com/"
TARGET_URL = "https://erp.rt.com/ims/inventory-existence-list"
DOWNLOAD_DIR = r"C:\Users\Ventas-Inventario_icompras_OTC-RX\inventario"

# Lista de sucursales objetivo
LISTA_SUCURSALES = [
    "Sucursal 1", "Sucursal 2", "Sucursal 3", "Sucursal 4", "Sucursal 5"
]

# Configuración de opciones de Chrome para descarga automática
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--disable-popup-blocking")
# Si se desea ejecutar sin ventana (headless), descomentar la siguiente línea
# options.add_argument("--headless")

driver = webdriver.Chrome(options=options)
driver.maximize_window()
wait = WebDriverWait(driver, 180)

# ================= FUNCIONES AUXILIARES =================
def esperar_y_click(xpath, timeout=180):
    """Espera a que el elemento sea clickeable y hace clic."""
    elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    elem.click()

def esperar_y_enviar_texto(xpath, texto, timeout=180):
    """Espera a que el elemento sea visible y envía texto."""
    elem = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))
    elem.clear()
    elem.send_keys(texto)

def obtener_texto_si_existe(xpath, timeout=10):
    """Intenta obtener el texto de un elemento, retorna None si no aparece."""
    try:
        elem = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))
        return elem.text
    except TimeoutException:
        return None

def refrescar_pagina():
    """Refresca la página actual y espera a que el body esté presente."""
    driver.refresh()
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

def asegurar_en_pagina_correcta():
    """
    Verifica si la URL actual es la página de inicio (home) y, en ese caso,
    regresa a la página de inventario. Espera a que la página de inventario
    esté completamente cargada (breadcrumb visible).
    """
    current = driver.current_url.rstrip('/')
    if current == "https://erp.rt.com/home":
        print("⚠️ Detectada redirección a home. Regresando a inventario...")
        driver.get(TARGET_URL)
        breadcrumb_xpath = "/html/body/app-root/app-layout/div/div/app-breadcrumb/div/div/div/p-breadcrumb/div/ul/li[5]/a/span"
        WebDriverWait(driver, 180).until(EC.presence_of_element_located((By.XPATH, breadcrumb_xpath)))
        print("✅ Página de inventario cargada nuevamente.")

def seleccionar_sucursal_por_nombre(nombre_sucursal):
    """
    Abre el dropdown de sucursales, busca la opción que coincida exactamente
    con el nombre dado y hace clic en ella.
    """
    asegurar_en_pagina_correcta()  # Verificar antes de interactuar
    
    # Click para desplegar el dropdown
    dropdown_trigger_xpath = "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[2]"
    esperar_y_click(dropdown_trigger_xpath)
    
    # Esperar a que aparezca el contenedor de opciones
    opciones_container_xpath = "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[3]/div[2]/ul"
    wait.until(EC.presence_of_element_located((By.XPATH, opciones_container_xpath)))
    
    # Obtener todos los items del dropdown (li con clase específica)
    items_xpath = f"{opciones_container_xpath}/p-dropdownitem"
    items = driver.find_elements(By.XPATH, items_xpath)
    
    encontrado = False
    for idx, item in enumerate(items, start=1):
        texto_elemento = item.text.strip()
        if texto_elemento == nombre_sucursal:
            item.click()
            encontrado = True
            break
        try:
            span_text = item.find_element(By.XPATH, ".//span").text.strip()
            if span_text == nombre_sucursal:
                item.click()
                encontrado = True
                break
        except:
            pass
    
    if not encontrado:
        raise Exception(f"No se encontró la sucursal '{nombre_sucursal}' en el dropdown.")
    
    time.sleep(1)

def seleccionar_fecha_actual():
    """
    Abre el calendario, selecciona el año actual, mes actual y día actual.
    Adaptado de la lógica robusta del código de ventas (Código 1).
    """
    asegurar_en_pagina_correcta()
    
    hoy = datetime.now()
    year = hoy.year
    month = hoy.month
    day = hoy.day
    
    # XPath base del componente p-calendar (sin el /span/button final)
    calendar_base = "/html/body/app-root/app-layout/div/div/div[1]/app-inventory-existence-list/div[1]/app-inventory-existence-filters-panel/div/div[5]/span/p-calendar"
    
    # 1. Hacer click en el botón del calendario para desplegarlo
    btn_cal = calendar_base + "/span/button"
    time.sleep(1)
    esperar_y_click(btn_cal)
    
    
    # 2. Hacer click en el botón para seleccionar el año (button[2])
    btn_year = calendar_base + "/span/div/div/div/div[1]/div/button[2]"
    time.sleep(1)    
    esperar_y_click(btn_year)

    
    # 3. Dentro del contenedor de años, buscar el span con el año actual
    year_container = calendar_base + "/span/div/div[2]"
    wait.until(EC.presence_of_element_located((By.XPATH, year_container)))
    year_spans = driver.find_elements(By.XPATH, year_container + "//span[contains(@class, 'p-yearpicker-year')]")
    for span in year_spans:
        if span.text.strip() == str(year):
            span.click()
            break
    time.sleep(1)
    
    # 4. Seleccionar el mes: el mes 'm' corresponde al span número 'm' (Enero=1, Diciembre=12)
    month_span_xpath = calendar_base + f"/span/div/div[2]/span[{month}]"
    esperar_y_click(month_span_xpath, timeout=10)
    time.sleep(1)
    
    # 5. Seleccionar el día dentro de la tabla de días
    tbody_xpath = calendar_base + "/span/div/div/div/div[2]/table/tbody"
    wait.until(EC.presence_of_element_located((By.XPATH, tbody_xpath)))
    days = driver.find_elements(By.XPATH, tbody_xpath + "//td[contains(@class, 'ng-star-inserted')]")
    for td in days:
        if td.text.strip() == str(day):
            td.click()
            break
    time.sleep(1)

def seleccionar_categorias():
    """
    Abre el panel de categorías y selecciona las dos checkboxes correspondientes
    a los elementos tr[26] y tr[27] dentro del treetable.
    """
    asegurar_en_pagina_correcta()
    
    btn_categorias_xpath = "/html/body/app-root/app-layout/div/div/div[1]/app-inventory-existence-list/div[1]/app-inventory-existence-filters-panel/div/div[6]/div/p-button/button"
    esperar_y_click(btn_categorias_xpath)
    
    time.sleep(1)
    chk1_xpath = "/html/body/div/div/p-treetable/div/div/table/tbody/tr[26]/td/p-treetablecheckbox/div/div[2]"
    chk2_xpath = "/html/body/div/div/p-treetable/div/div/table/tbody/tr[27]/td/p-treetablecheckbox/div/div[2]"
    
    esperar_y_click(chk1_xpath, timeout=10)
    esperar_y_click(chk2_xpath, timeout=10)
    
    esperar_y_click(btn_categorias_xpath)

def exportar_y_renombrar(nombre_sucursal, fecha_consulta):
    """
    Hace clic en el botón de exportar Excel, monitorea la carpeta de descargas,
    detecta el archivo nuevo y lo renombra según el formato:
    "{nombre_sucursal}_EI+{fecha_ddmmaa}.xlsx"
    """
    asegurar_en_pagina_correcta()
    
    exportar_btn_xpath = "/html/body/app-root/app-layout/div/div/div[1]/app-inventory-existence-list/div[1]/app-inventory-existence-filters-panel/div/div[14]/button[3]"
    esperar_y_click(exportar_btn_xpath)
    
    fecha_str = fecha_consulta.strftime("%d-%m-%Y")
    nombre_esperado = f"{nombre_sucursal}_EI+{fecha_str}.xlsx"
    ruta_destino = os.path.join(DOWNLOAD_DIR, nombre_esperado)
    
    archivos_antes = set(os.listdir(DOWNLOAD_DIR))
    timeout = 180
    inicio = time.time()
    archivo_nuevo = None
    while time.time() - inicio < timeout:
        time.sleep(1)
        archivos_despues = set(os.listdir(DOWNLOAD_DIR))
        nuevos = archivos_despues - archivos_antes
        xlsx_nuevos = [f for f in nuevos if f.endswith('.xlsx') and not f.endswith('.crdownload')]
        if xlsx_nuevos:
            archivo_nuevo = xlsx_nuevos[0]
            break
    
    if archivo_nuevo is None:
        raise Exception("No se detectó la descarga del archivo Excel en el tiempo esperado.")
    
    ruta_origen = os.path.join(DOWNLOAD_DIR, archivo_nuevo)
    if os.path.exists(ruta_destino):
        os.remove(ruta_destino)
    shutil.move(ruta_origen, ruta_destino)
    print(f"Archivo renombrado: {nombre_esperado}")

# ================= FLUJO PRINCIPAL =================
try:
    # 1. Login
    driver.get(LOGIN_URL)
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='email' or @placeholder='Usuario']")))
    esperar_y_enviar_texto("/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[1]/input", USER)
    esperar_y_enviar_texto("/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[2]/input", PASSWORD)
    esperar_y_click("/html/body/app-root/app-login/div/div/div[2]/div/div[2]/form/div[4]/button")
    home_page = "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[2]"
    WebDriverWait(driver, 180).until(EC.presence_of_element_located((By.XPATH, home_page)))
    
    # 2. Navegar a la página de inventario y esperar breadcrumb
    driver.get(TARGET_URL)
    asegurar_en_pagina_correcta()  # Verificar que no haya redirigido a home
    breadcrumb_xpath = "/html/body/app-root/app-layout/div/div/app-breadcrumb/div/div/div/p-breadcrumb/div/ul/li[5]/a/span"
    WebDriverWait(driver, 180).until(EC.presence_of_element_located((By.XPATH, breadcrumb_xpath)))
    
    # 3. Obtener número de sucursales y lista (opcional)
    dropdown_trigger = "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[2]"
    esperar_y_click(dropdown_trigger)
    container = "/html/body/app-root/app-layout/div/div/app-panel-topbar/div/div/div[1]/span/app-current-office-selector/div/div[2]/p-dropdown/div/div[3]/div[2]/ul"
    wait.until(EC.presence_of_element_located((By.XPATH, container)))
    items = driver.find_elements(By.XPATH, f"{container}/p-dropdownitem")
    total_sucursales = len(items)
    print(f"Total sucursales disponibles: {total_sucursales}")
    driver.find_element(By.TAG_NAME, "body").click()
    refrescar_pagina()
    asegurar_en_pagina_correcta()     
    wait.until(EC.presence_of_element_located((By.XPATH, breadcrumb_xpath)))
    
    # 4. Iterar sobre cada sucursal objetivo
    fecha_hoy = datetime.now()
    for sucursal in LISTA_SUCURSALES:
        print(f"\nProcesando sucursal: {sucursal}")
        
        asegurar_en_pagina_correcta()  # Verificar antes de cada iteración
        
        # Seleccionar sucursal
        seleccionar_sucursal_por_nombre(sucursal)
        
        # 5. Seleccionar fecha actual
        seleccionar_fecha_actual()
        
        # 6. Seleccionar categorías
        seleccionar_categorias()
        
        # 7. Hacer clic en buscar y esperar resultados
        asegurar_en_pagina_correcta()
        buscar_btn_xpath = "/html/body/app-root/app-layout/div/div/div[1]/app-inventory-existence-list/div[1]/app-inventory-existence-filters-panel/div/div[14]/button[1]"
        esperar_y_click(buscar_btn_xpath)
        resultado_xpath = "/html/body/app-root/app-layout/div/div/div[1]/app-inventory-existence-list/div[2]/p-table/div/div/table/tbody/tr[1]/td[1]/button[1]"
        wait.until(EC.presence_of_element_located((By.XPATH, resultado_xpath)))
        
        # 8. Exportar y renombrar archivo
        exportar_y_renombrar(sucursal, fecha_hoy)
        
        # 9. Refrescar página para la siguiente sucursal
        refrescar_pagina()
        asegurar_en_pagina_correcta()  # Después del refresh, verificar URL
        wait.until(EC.presence_of_element_located((By.XPATH, breadcrumb_xpath)))
        
        print(f"Finalizada sucursal: {sucursal}")
    
    print("\nProceso completado para todas las sucursales.")

except Exception as e:
    print(f"Error durante la ejecución: {e}")
    driver.save_screenshot("error.png")
finally:
    driver.quit()