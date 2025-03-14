import time
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import flet as ft

# Configuración del navegador
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")  # Ejecutar sin interfaz gráfica
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36"
)

# Preinstalar el driver
driver_path = ChromeDriverManager().install()

# URL del sistema RASFF con la tabla de incidencias
RASFF_URL = "https://webgate.ec.europa.eu/rasff-window/screen/search?searchQueries=eyJkYXRlIjp7InN0YXJ0UmFuZ2UiOiIiLCJlbmRSYW5nZSI6IiJ9LCJjb3VudHJpZXMiOnt9LCJ0eXBlIjp7fSwibm90aWZpY2F0aW9uU3RhdHVzIjp7fSwicHJvZHVjdCI6e30sInJpc2siOnt9LCJyZWZlcmVuY2UiOiIiLCJzdWJqZWN0IjoiIn0%3D"

def seleccionar_fecha_en_calendario(driver, fecha_seleccionada):
    """Selecciona una fecha en el calendario emergente."""
    # Formatear la fecha seleccionada
    dia = fecha_seleccionada.strftime("%d").lstrip("0")  # Día sin ceros a la izquierda (ej. "5" en lugar de "05")
    mes = fecha_seleccionada.strftime("%b").upper()      # Mes abreviado en mayúsculas (ej. "MAR")
    año = fecha_seleccionada.strftime("%Y")              # Año completo (ej. "2025")

    # Hacer clic en el botón del calendario para abrirlo
    calendario_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.eui-date-range-selector__toggler"))
    )
    calendario_button.click()

    # Esperar a que el calendario esté visible
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.TAG_NAME, "mat-calendar"))
    )

    # Seleccionar el año
    year_selector = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//button[contains(@aria-label, '{año}')]"))
    )
    year_selector.click()

    # Seleccionar el mes
    month_selector = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//button[contains(@aria-label, '{mes}')]"))
    )
    month_selector.click()

    # Seleccionar el día
    day_selector = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//td[contains(@aria-label, '{dia} {mes} {año}')]"))
    )
    day_selector.click()

    # El calendario debería cerrarse automáticamente y llenar ambos campos si se selecciona el rango

def extraer_alertas(fecha_seleccionada):
    """Extrae las alertas del sistema RASFF filtrando por la fecha seleccionada en la barra de búsqueda."""
    driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
    alertas = []

    try:
        print("Navegando a la página del RASFF...")
        driver.get(RASFF_URL)

        # Esperar a que el formulario de búsqueda esté presente
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "search"))
        )

        # Seleccionar la fecha en el calendario para ambos campos (inicio y fin)
        seleccionar_fecha_en_calendario(driver, fecha_seleccionada)

        # Verificar que los campos se hayan llenado (opcional, para depuración)
        start_date_input = driver.find_element(By.CSS_SELECTOR, "input.mat-start-date")
        end_date_input = driver.find_element(By.CSS_SELECTOR, "input.mat-end-date")
        print(f"Valor en Start Date: {start_date_input.get_attribute('value')}")
        print(f"Valor en End Date: {end_date_input.get_attribute('value')}")

        # Hacer clic en el botón "Search"
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Submit Search']"))
        )
        search_button.click()

        # Esperar a que la tabla se actualice con los resultados filtrados
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.eui-table"))
        )
        time.sleep(2)  # Pausa adicional para asegurar la carga completa

        # Localizar la tabla específica
        tabla = driver.find_element(By.CSS_SELECTOR, "table.eui-table.eui-table--hoverable.eui-table--responsive")

        # Extraer los encabezados de la tabla
        encabezados = [th.text.strip() for th in tabla.find_elements(By.CSS_SELECTOR, "thead th")]

        # Extraer las filas de datos
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")

        # Extraer todas las filas sin filtrar manualmente (ya están filtradas por la búsqueda)
        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            datos_fila = {encabezados[i]: celdas[i].text.strip() for i in range(len(celdas))}
            alertas.append(datos_fila)

        print(f"Se encontraron {len(alertas)} alertas para la fecha {fecha_seleccionada.strftime('%d/%m/%Y')}.")

    except Exception as e:
        print(f"Error al extraer las alertas: {e}")

    finally:
        driver.quit()

    return alertas

def guardar_en_excel(alertas, fecha_seleccionada):
    """Guarda las alertas en un archivo Excel en la carpeta 'Descargas'."""
    if not alertas:
        print("No hay datos para guardar.")
        return None

    df = pd.DataFrame(alertas)
    fecha_formateada = fecha_seleccionada.strftime("%Y-%m-%d")
    nombre_archivo = f"alertas_RASFF_{fecha_formateada}.xlsx"

    # Ruta de la carpeta "Descargas"
    ruta_guardar = os.path.join(os.path.expanduser("~"), "Downloads")
    ruta_completa = os.path.join(ruta_guardar, nombre_archivo)

    df.to_excel(ruta_completa, index=False)
    print(f"Datos guardados en '{ruta_completa}'.")
    return ruta_completa

def main(page: ft.Page):
    """Función principal de la interfaz gráfica con Flet."""
    page.title = "Extractor de Alertas RASFF"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    # Campo de texto para introducir la fecha manualmente
    campo_fecha = ft.TextField(
        label="Introduce la fecha (DD/MM/YYYY)",
        hint_text="Ejemplo: 25/10/2023",
        width=200
    )

    # Texto para mostrar el estado
    estado_texto = ft.Text(value="Introduce una fecha y genera el Excel.", size=16)

    def generar_excel(e):
        """Función que se ejecuta al presionar el botón de generar Excel."""
        fecha_texto = campo_fecha.value.strip()

        if not fecha_texto:
            estado_texto.value = "Por favor, introduce una fecha."
            page.update()
            return

        try:
            # Convertir el texto ingresado a un objeto datetime
            fecha_seleccionada = datetime.strptime(fecha_texto, "%d/%m/%Y")
        except ValueError:
            estado_texto.value = "Formato de fecha inválido. Usa DD/MM/YYYY (ejemplo: 25/10/2023)."
            page.update()
            return

        estado_texto.value = "Extrayendo datos, por favor espera..."
        page.update()

        # Extraer alertas con la fecha seleccionada usando la barra de búsqueda
        alertas = extraer_alertas(fecha_seleccionada)

        if alertas:
            ruta = guardar_en_excel(alertas, fecha_seleccionada)
            estado_texto.value = f"Excel generado en: {ruta}"
        else:
            estado_texto.value = "No se encontraron datos para la fecha seleccionada."

        page.update()

    # Botón para generar el Excel
    boton_generar = ft.ElevatedButton(
        text="Generar Excel",
        on_click=generar_excel
    )

    # Añadir elementos a la página
    page.add(
        ft.Column(
            [
                estado_texto,
                campo_fecha,
                boton_generar
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=20
        )
    )

# Ejecutar la aplicación
if __name__ == "__main__":
    ft.app(target=main)