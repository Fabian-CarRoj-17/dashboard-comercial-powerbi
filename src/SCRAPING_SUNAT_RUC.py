import asyncio
import pandas as pd
import random
from tqdm import tqdm
from playwright.async_api import async_playwright

ARCHIVO_ENTRADA = "pendientes_ruc.xlsx"
ARCHIVO_SALIDA = "resultado_rucs.xlsx"
COLUMNA = "RUC_LIMPIO"

URL = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"

MAX_REINTENTOS = 2
WORKERS = 20
NUM_CONTEXTS = 5

USER_AGENTS = [
    "Mozilla/5.0 Windows NT 10.0 Chrome/120.0.0.0",
    "Mozilla/5.0 Windows NT 10.0 Chrome/119.0.0.0",
    "Mozilla/5.0 Mac OS X Chrome/118.0.0.0",
    "Mozilla/5.0 Linux x86_64 Chrome/117.0.0.0",
    "Mozilla/5.0 Windows NT 10.0 Chrome/116.0.0.0"
]

# =============================
# ESTRUCTURA BASE
# =============================

def estructura_base(ruc):
    return {
        "DNI/RUC": ruc,
        "Razón Social": "",
        "Fecha Inscripción": "",
        "Fecha Inicio Actividades": "",
        "Estado Contribuyente": "",
        "Condición Contribuyente": "",
        "Domicilio Fiscal": "",
        "Sistema Emisión": "",
        "Actividad Comercio Exterior": "",
        "Actividad Económica": "",
        "Actividad Secundaria 1": "",
        "Actividad Secundaria 2": "",
        "Emisor Electrónico Desde": "",
        "Periodo": "",
        "Trabajadores": "",
        "Pensionistas": "",
        "Prestadores": "",
        "Trabajadores Estado": "",
        "Estado Scraping": ""
    }

# =============================
# FUNCIONES DE EXTRACCIÓN PRECISAS
# =============================

async def get_text(page, label, nth=0):
    try:
        return (await page.locator(f"text={label}").locator("..").locator("..").locator("p").nth(nth).inner_text()).strip()
    except:
        return ""

# =============================
# SCRAPER RUC (HÍBRIDO)
# =============================

async def procesar_ruc(context, semaphore, ruc):
    async with semaphore:
        page = await context.new_page()
        try:
            for _ in range(MAX_REINTENTOS):
                try:
                    await page.goto(URL, wait_until="domcontentloaded")
                    await page.fill("#txtRuc", ruc)
                    await page.click("#btnAceptar")
                    await page.wait_for_selector(".list-group", timeout=20000)

                    contenido = await page.content()
                    if "no se encuentra registrado" in contenido.lower():
                        data = estructura_base(ruc)
                        data["Estado Scraping"] = "no_existe"
                        await page.close()
                        return data

                    data = estructura_base(ruc)

                    # Razón Social
                    try:
                        nombre = await page.locator(f"h4.list-group-item-heading >> text={ruc}").inner_text()
                        data["Razón Social"] = nombre.split(" - ", 1)[1].strip() if " - " in nombre else nombre.strip()
                    except:
                        pass

                    # Campos básicos
                    data["Fecha Inscripción"] = await get_text(page, "Fecha de Inscripción:")
                    data["Fecha Inicio Actividades"] = await get_text(page, "Fecha de Inicio de Actividades:", nth=1)
                    data["Estado Contribuyente"] = await get_text(page, "Estado del Contribuyente:")
                    data["Condición Contribuyente"] = await get_text(page, "Condición del Contribuyente:")
                    data["Domicilio Fiscal"] = await get_text(page, "Domicilio Fiscal:")
                    data["Sistema Emisión"] = await get_text(page, "Sistema Emisión de Comprobante:")
                    data["Actividad Comercio Exterior"] = await get_text(page, "Actividad Comercio Exterior:", nth=1)
                    data["Emisor Electrónico Desde"] = await get_text(page, "Emisor electrónico desde:")

                    # Actividades económicas
                    actividad_principal, actividad_sec1, actividad_sec2 = "", "", ""
                    tabla_actividades = page.locator("text=Actividad(es) Económica(s):").locator("..").locator("..").locator("tr")
                    cantidad = await tabla_actividades.count()
                    for i in range(cantidad):
                        texto = (await tabla_actividades.nth(i).inner_text()).strip()
                        if texto.startswith("Principal"):
                            actividad_principal = texto
                        elif texto.startswith("Secundaria 1"):
                            actividad_sec1 = texto
                        elif texto.startswith("Secundaria 2"):
                            actividad_sec2 = texto
                    data["Actividad Económica"] = actividad_principal
                    data["Actividad Secundaria 1"] = actividad_sec1
                    data["Actividad Secundaria 2"] = actividad_sec2

                    # Trabajadores
                    try:
                        await page.click(".btnInfNumTra")
                        await page.wait_for_load_state("domcontentloaded")
                        contenido_trab = await page.content()

                        if "No existen declaraciones presentadas" in contenido_trab:
                            data["Trabajadores Estado"] = "sin_declaraciones"
                            data["Periodo"] = ""
                            data["Trabajadores"] = "0"
                            data["Pensionistas"] = "0"
                            data["Prestadores"] = "0"
                        else:
                            filas = page.locator("table tbody tr")
                            total = await filas.count()
                            if total > 0:
                                ultima = filas.last
                                columnas = ultima.locator("td")
                                data["Periodo"] = (await columnas.nth(0).inner_text()).strip()
                                data["Trabajadores"] = (await columnas.nth(1).inner_text()).strip()
                                data["Pensionistas"] = (await columnas.nth(2).inner_text()).strip()
                                data["Prestadores"] = (await columnas.nth(3).inner_text()).strip()
                                data["Trabajadores Estado"] = "ok"
                    except:
                        data["Trabajadores Estado"] = "error_trabajadores"

                    data["Estado Scraping"] = "ok"
                    await page.close()
                    return data

                except:
                    await asyncio.sleep(random.uniform(1.5, 2.5))

            data = estructura_base(ruc)
            data["Estado Scraping"] = "error_total"
            await page.close()
            return data

        except:
            data = estructura_base(ruc)
            data["Estado Scraping"] = "error_critico"
            await page.close()
            return data

# =============================
# MAIN
# =============================

async def main():
    df = pd.read_excel(ARCHIVO_ENTRADA, dtype=str)
    df[COLUMNA] = df[COLUMNA].astype(str).str.strip()
    rucs = df[COLUMNA].tolist()

    resultados = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"]
        )

        contexts = []
        for i in range(NUM_CONTEXTS):
            ua = USER_AGENTS[i % len(USER_AGENTS)]
            context = await browser.new_context(user_agent=ua)
            await context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
            """)
            contexts.append(context)

        semaphore = asyncio.Semaphore(WORKERS)

        tareas = [
            procesar_ruc(contexts[i % NUM_CONTEXTS], semaphore, ruc)
            for i, ruc in enumerate(rucs)
        ]

        for future in tqdm(asyncio.as_completed(tareas), total=len(tareas)):
            resultado = await future
            resultados.append(resultado)

            if len(resultados) % 100 == 0:
                pd.DataFrame(resultados).to_excel(ARCHIVO_SALIDA, index=False)

        await browser.close()

    pd.DataFrame(resultados).to_excel(ARCHIVO_SALIDA, index=False)
    print("Proceso finalizado correctamente.")

# =============================
# RUN
# =============================

asyncio.run(main())