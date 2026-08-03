import asyncio
from playwright.async_api import async_playwright
import pandas as pd
 
async def scrape_ciiu_sunat():
    url = "https://ww3.sunat.gob.pe/ol-ti-itinsrucsol/utilAlias?proceso=A&accion=cargarBusquedaCIIU4&paginaInvocadora=PaginaInicialInscripRUC02.jsp&formularioInvocador=inscRuc&campoCIIU3=ciiuselected&campoCIIU4=codCiiuSelected4&campoCIIU4_Desc=ciiuSelected4"
    data = []
 
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=50)
        page = await browser.new_page()
        await page.goto(url)
        await page.wait_for_load_state('networkidle')
 
        # 🔹 Ir a pestaña "Búsqueda por Act. Económica"
        await page.click('a[href="#dvBusquedaPorActEconomica"]')
        await page.wait_for_selector('#dvBusquedaPorActEconomica', timeout=60000)
        await page.wait_for_timeout(1000)
 
        # 🔹 Extraer secciones
        secciones = await page.query_selector_all('select[name="selectSeccion"] option')
        for seccion in secciones[1:]:
            seccion_value = await seccion.get_attribute('value')
            seccion_text = await seccion.inner_text()
 
            await page.select_option('select[name="selectSeccion"]', seccion_value)
            await page.wait_for_timeout(500)
 
            # 🔹 Extraer divisiones
            divisiones = await page.query_selector_all('select[name="selectDivision"] option')
            for division in divisiones[1:]:
                division_value = await division.get_attribute('value')
                division_text = await division.inner_text()
 
                await page.select_option('select[name="selectDivision"]', division_value)
                await page.wait_for_timeout(500)
 
                # 🔹 Extraer clases
                clases = await page.query_selector_all('select[name="selectClase"] option')
                for clase in clases[1:]:
                    clase_value = await clase.get_attribute('value')
                    clase_text = await clase.inner_text()
 
                    if not clase_value:
                        continue
 
                    await page.select_option('select[name="selectClase"]', clase_value)
                    await page.wait_for_timeout(500)
 
                    # 🔹 Extraer CIIU final (5 dígitos)
                    ciius = await page.query_selector_all('select[name="listaacteconv3_2"] option')
                    for ciiu in ciius:
                        ciiu_value = await ciiu.get_attribute('value')
                        ciiu_text = await ciiu.inner_text()
 
                        if not ciiu_value:
                            continue
 
                        data.append({
                            "Sección": seccion_text,
                            "Código Sección": seccion_value,
                            "División": division_text,
                            "Código División": division_value,
                            "Clase": clase_text.strip(),
                            "Código Clase": clase_value,
                            "CIIU": ciiu_text.strip()  # Código final de 5 dígitos
                        })
 
        await browser.close()
 
    df = pd.DataFrame(data)
    df.to_excel("diccionario_ciiu_sunat_completo.xlsx", index=False)
    print("✅ Diccionario CIIU exportado como diccionario_ciiu_sunat_completo.xlsx")
 
# Ejecutar
asyncio.run(scrape_ciiu_sunat())