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

# Configuración navegador
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36"
)

driver_path = ChromeDriverManager().install()

RASFF_URL = "https://webgate.ec.europa.eu/rasff-window/screen/search?searchQueries=eyJkYXRlIjp7InN0YXJ0UmFuZ2UiOiIiLCJlbmRSYW5nZSI6IiJ9LCJjb3VudHJpZXMiOnt9LCJ0eXBlIjp7fSwibm90aWZpY2F0aW9uU3RhdHVzIjp7fSwicHJvZHVjdCI6e30sInJpc2siOnt9LCJyZWZlcmVuY2UiOiIiLCJzdWJqZWN0IjoiIn0%3D"

# Extraer alertas
def extraer_alertas(fecha_seleccionada):
    driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
    alertas = []

    try:
        driver.get(RASFF_URL)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.eui-table"))
        )
        time.sleep(2)

        page_size_select = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "select.page-size__select.eui-select"))
        )
        driver.execute_script(
            "arguments[0].value = '100'; arguments[0].dispatchEvent(new Event('change'));",
            page_size_select,
        )
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.eui-table"))
        )
        time.sleep(3)

        fecha_formateada = fecha_seleccionada.strftime("%d %b %Y").lstrip("0").upper()

        tabla = driver.find_element(By.CSS_SELECTOR, "table.eui-table.eui-table--hoverable.eui-table--responsive")
        encabezados = [th.text.strip() for th in tabla.find_elements(By.CSS_SELECTOR, "thead th")]
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")

        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            datos_fila = {encabezados[i]: celdas[i].text.strip() for i in range(len(celdas))}
            if datos_fila.get("Date") == fecha_formateada:
                alertas.append(datos_fila)

    except Exception as e:
        print(f"Error al extraer alertas: {e}")

    finally:
        driver.quit()

    return alertas

# Guardar en Excel
def guardar_en_excel(alertas, ruta):
    if not alertas:
        return False

    df_nuevo = pd.DataFrame(alertas)

    try:
        if os.path.exists(ruta):
            df_historico = pd.read_excel(ruta)
            df_combined = pd.concat([df_historico, df_nuevo], ignore_index=True)
            df_combined.drop_duplicates(inplace=True)
            df_combined.to_excel(ruta, index=False)
        else:
            df_nuevo.to_excel(ruta, index=False)
        return True
    except Exception as e:
        print(f"Error al guardar en Excel: {e}")
        return False

# Interfaz Flet
def main(page: ft.Page):
    page.title = "RASFF Alerts Extractor"
    page.padding = 20
    page.scroll = "auto"
    ainia_orange = "#F26C21"
    bg_color = "#FFF8F0"

    page.bgcolor = bg_color

    selected_date = {"value": None}

    estado_texto = ft.Text(
        value="Selecciona una fecha y un archivo Excel.",
        size=16,
        color="#444444",
        text_align=ft.TextAlign.CENTER,
        animate_opacity=300,
    )

    selected_date_text = ft.Text("No se ha seleccionado fecha", color="#666666")

    def date_picker_changed(e):
        fecha_str = e.data
        fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%dT%H:%M:%S.%f")
        selected_date["value"] = fecha_obj
        selected_date_text.value = f"Fecha seleccionada: {fecha_obj.strftime('%d/%m/%Y')}"
        page.update()

    date_picker = ft.DatePicker(
        on_change=date_picker_changed,
        first_date=datetime(2020, 1, 1),
        last_date=datetime.now(),
    )
    page.overlay.append(date_picker)

    def open_date_picker(_):
        date_picker.open = True
        page.update()

    date_button = ft.ElevatedButton(
        "Seleccionar Fecha",
        icon=ft.Icons.CALENDAR_MONTH,
        on_click=open_date_picker,
        style=ft.ButtonStyle(
            color="white",
            bgcolor=ainia_orange,
            shape=ft.RoundedRectangleBorder(radius=10),
            overlay_color="#ff8c42",
            elevation=2,
        ),
    )

    file_picker = ft.FilePicker(on_result=lambda e: file_picker_result(e))
    page.overlay.append(file_picker)

    def file_picker_result(e):
        if e.files and len(e.files) > 0:
            selected_file = e.files[0]
            campo_ruta.value = selected_file.path
            page.update()

    campo_ruta = ft.TextField(
        label="Ruta archivo Excel",
        hint_text="Ej: C:/.../Historico.xlsx",
        width=400,
        value="C:/Users/bec-smi/Documents/Historico RASFF.xlsx",
        border_color=ainia_orange,
        focused_border_color=ainia_orange,
        read_only=True,
        cursor_color=ainia_orange,
    )

    browse_button = ft.ElevatedButton(
        "Examinar",
        icon=ft.Icons.FOLDER_OPEN,
        on_click=lambda _: file_picker.pick_files(
            allowed_extensions=["xlsx", "xls"],
            dialog_title="Seleccionar archivo Excel"
        ),
        style=ft.ButtonStyle(
            color="white",
            bgcolor=ainia_orange,
            shape=ft.RoundedRectangleBorder(radius=10),
            overlay_color="#ff8c42",
            elevation=2,
        ),
    )

    file_row = ft.Row(
        [campo_ruta, browse_button],
        alignment=ft.MainAxisAlignment.CENTER,
        spacing=10,
    )

    def extraer_y_guardar(e):
        if not selected_date["value"]:
            estado_texto.value = "⚠️ Por favor, selecciona una fecha válida."
        elif not campo_ruta.value.strip():
            estado_texto.value = "⚠️ Por favor, ingresa la ruta del Excel."
        else:
            estado_texto.value = "⏳ Extrayendo alertas..."
            page.update()

            alertas = extraer_alertas(selected_date["value"])

            if not alertas:
                estado_texto.value = f"ℹ️ No se encontraron alertas para {selected_date['value'].strftime('%d/%m/%Y')}."
            else:
                guardado = guardar_en_excel(alertas, campo_ruta.value)

                if guardado:
                    estado_texto.value = f"✅ {len(alertas)} alertas guardadas correctamente."
                else:
                    estado_texto.value = "❌ Error al guardar en Excel. Revisa la ruta o permisos."

        estado_texto.opacity = 1
        page.update()

    boton_extraer = ft.ElevatedButton(
        text="Extraer y Guardar",
        on_click=extraer_y_guardar,
        style=ft.ButtonStyle(
            color="white",
            bgcolor=ainia_orange,
            padding=15,
            shape=ft.RoundedRectangleBorder(radius=12),
            overlay_color="#ff8c42",
            elevation=3,
        ),
        width=220,
    )

    logo = ft.Image(
        src="https://www.ainia.com/wp-content/uploads/2022/01/LOGO-AINIA-simple-alta-resolucion-sin-fondo-1.png",
        width=200,
        fit=ft.ImageFit.CONTAIN,
    )

    titulo = ft.Text(
        "RASFF Alerts Extractor",
        size=30,
        weight=ft.FontWeight.BOLD,
        color=ainia_orange,
    )

    # Layout principal
    page.add(
        ft.Column(
            [
                logo,
                titulo,
                estado_texto,
                ft.Divider(color=ainia_orange, thickness=1),
                ft.Row([date_button, selected_date_text], alignment=ft.MainAxisAlignment.CENTER),
                file_row,
                boton_extraer,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=25,
        )
    )

if __name__ == "__main__":
    ft.app(target=main)
