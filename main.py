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

# Browser configuration
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36"
)

# Pre-install driver
driver_path = ChromeDriverManager().install()

# RASFF URL
RASFF_URL = "https://webgate.ec.europa.eu/rasff-window/screen/search?searchQueries=eyJkYXRlIjp7InN0YXJ0UmFuZ2UiOiIiLCJlbmRSYW5nZSI6IiJ9LCJjb3VudHJpZXMiOnt9LCJ0eXBlIjp7fSwibm90aWZpY2F0aW9nU3RhdHVzIjp7fSwicHJvZHVjdCI6e30sInJpc2siOnt9LCJyZWZlcmVuY2UiOiIiLCJzdWJqZWN0IjoiIn0%3D"

def extraer_alertas(fecha_seleccionada):
    """Extracts RASFF alerts for the specified date from the first page with 100 records."""
    driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
    alertas = []
    
    try:
        driver.get(RASFF_URL)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.eui-table"))
        )
        time.sleep(2)

        # Change page size to 100
        page_size_select = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "select.page-size__select.eui-select"))
        )
        driver.execute_script("arguments[0].value = '100'; arguments[0].dispatchEvent(new Event('change'));",
                            page_size_select)
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.eui-table"))
        )
        time.sleep(3)

        # Format the selected date
        fecha_formateada = fecha_seleccionada.strftime("%d %b %Y").replace(" 0", " ").upper()

        # Locate table
        tabla = driver.find_element(By.CSS_SELECTOR, "table.eui-table.eui-table--hoverable.eui-table--responsive")
        encabezados = [th.text.strip() for th in tabla.find_elements(By.CSS_SELECTOR, "thead th")]
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")

        # Extract matching rows
        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            datos_fila = {encabezados[i]: celdas[i].text.strip() for i in range(len(celdas))}
            if datos_fila.get("Date") == fecha_formateada:
                alertas.append(datos_fila)

    except Exception as e:
        print(f"Error extracting alerts: {e}")
    
    finally:
        driver.quit()
    
    return alertas

def guardar_en_excel(alertas, ruta):
    """Adds alerts to the existing historical Excel file."""
    if not alertas:
        return None

    df_nuevo = pd.DataFrame(alertas)
    
    if os.path.exists(ruta):
        df_historico = pd.read_excel(ruta)
        df_combined = pd.concat([df_historico, df_nuevo], ignore_index=True)
        df_combined.to_excel(ruta, index=False)
        return ruta
    else:
        return None

def main(page: ft.Page):
    """Main function for the Flet GUI."""
    page.title = "RASFF Alerts Extractor"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.bgcolor = "#f5f5f5"  # Light gray background
    page.padding = 20

    # AINIA Logo
    logo = ft.Image(
        src="https://www.ainia.com/wp-content/uploads/2022/01/LOGO-AINIA-simple-alta-resolucion-sin-fondo-1.png",
        width=200,
        height=100,
        fit=ft.ImageFit.CONTAIN
    )
    
    # Title with styling
    title = ft.Text(
        "RASFF Alerts Extractor",
        size=28,
        weight=ft.FontWeight.BOLD,
        color="#00529b"  # AINIA blue color
    )
    
    # Text field for date with improved styling
    campo_fecha = ft.TextField(
        label="Enter date (DD/MM/YYYY)",
        hint_text="Example: 25/10/2023",
        width=300,
        border_color="#00529b",
        focused_border_color="#00529b",
        cursor_color="#00529b"
    )

    # Text field for file path with improved styling
    campo_ruta = ft.TextField(
        label="Enter Excel file path",
        hint_text="Example: C:/Users/username/Documents/Historico RAFF.xlsx",
        width=500,
        value="C:/Users/bec-smi/Documents/Historico RAFF.xlsx",  # Default value
        border_color="#00529b",
        focused_border_color="#00529b",
        cursor_color="#00529b"
    )

    # Status text with improved styling
    estado_texto = ft.Text(
        value="Enter date and file path, then click to extract.",
        size=16,
        color="#555555",
        text_align=ft.TextAlign.CENTER
    )

    def extraer_y_guardar(e):
        """Function executed when extract button is clicked."""
        # ... existing code ...

    # Extract button with improved styling
    boton_extraer = ft.ElevatedButton(
        text="Extract and Save",
        on_click=extraer_y_guardar,
        style=ft.ButtonStyle(
            color="white",
            bgcolor="#00529b",
            padding=15,
            shape=ft.RoundedRectangleBorder(radius=8)
        ),
        width=200,
        height=50
    )

    # Card container for form elements
    form_card = ft.Card(
        content=ft.Container(
            content=ft.Column(
                [
                    campo_fecha,
                    campo_ruta,
                    boton_extraer
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=20
            ),
            padding=20,
            border_radius=10
        ),
        elevation=5,
        margin=10
    )

    # Add elements to page
    page.add(
        ft.Column(
            [
                logo,
                title,
                estado_texto,
                form_card
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=20
        )
    )

if __name__ == "__main__":
    ft.app(target=main)
    
    
'''                                                    
        ██████╗  █████╗ ██████╗ ██╗      ██████╗ 
        ██╔══██╗██╔══██╗██╔══██╗██║     ██╔═══██╗
        ██████╔╝███████║██████╔╝██║     ██║   ██║
        ██╔═══╝ ██╔══██║██╔══██╗██║     ██║   ██║
        ██║     ██║  ██║██████╔╝███████╗╚██████╔╝
        ╚═╝     ╚═╝  ╚═╝╚═════╝ ╚══════╝ ╚═════╝ 
                                            
               /\      /\      /\      /\            
              /  \    /  \    /  \    /  \           
             /    \  /    \  /    \  /    \          
            /      \/      \/      \/      \         
           /|      ||      ||      ||      |\        
          / |      ||      ||      ||      | \       
         /  |      ||      ||      ||      |  \      
        /   |      ||      ||      ||      |   \     
       /    |      ||      ||      ||      |    \    
      /     |______||______||______||______|     \   
     /                                            \  
    /______________________________________________\
'''