import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError
import re
import random
import time
import asyncio
import sys
import os
from datetime import datetime
BASE_URL = "https://tramites.cancilleria.gov.co/apostillalegalizacion/solicitud/inicio.aspx"


# ==============================
# 📌 FUNCIONES DE EXCEL
# ==============================
def leer_excel(path="entrada.xlsx"):
    """Lee el Excel de entrada y devuelve DataFrame validado"""
    
    # Definimos las columnas esperadas en el orden correcto
    COLUMNAS_ESPERADAS = ["#", "NOMBRE", "CEDULA", "FECHA_EXP", "CODIGO", "LINK", "OBSERVACIONES"]

    # Leemos forzando CEDULA como texto
    df = pd.read_excel(path, dtype={"CEDULA": str})

    # Normalizamos nombres de columnas (quitamos espacios y mayúsculas)
    df.columns = [c.strip().upper() for c in df.columns]

    # Validamos que las columnas sean exactamente las esperadas
    columnas_actuales = df.columns.tolist()
    if columnas_actuales != COLUMNAS_ESPERADAS:
        raise ValueError(
            f"❌ El archivo no cumple con el formato.\n"
            f"Se esperaban columnas: {COLUMNAS_ESPERADAS}\n"
            f"Se encontraron: {columnas_actuales}"
        )

    # Limpiamos la columna CEDULA por si quedó algún .0
    df["CEDULA"] = df["CEDULA"].str.replace(r"\.0$", "", regex=True)

    return df

def guardar_excel(df, path="entrada.xlsx"):
    """Guarda los resultados en el mismo archivo (sin alterar el orden de columnas)"""
    df.to_excel(path, index=False)
    print(f"💾 Resultados actualizados en {path}")


# ==============================
# 📌 PÁGINA 1
# ==============================
def pagina1_inicio(page):
    """Selecciona tipo de trámite y acepta términos iniciales"""
    page.goto(BASE_URL, timeout=60_000)
    print("🌐 Página 1: Crear Solicitud")

    page.wait_for_selector("#contenido_ddlTipoSeleccion", timeout=15_000)
    page.select_option("#contenido_ddlTipoSeleccion", "21")

    page.wait_for_selector("#contenido_ddlTipoDocumento", timeout=5_000)
    page.select_option("#contenido_ddlTipoDocumento", "1")

    # Cerrar modal si aparece
    try:
        page.wait_for_selector("#contenido_ucInfor_panInformmacion", timeout=5_000)
        page.click("#contenido_ucInfor_lbClose")
    except TimeoutError:
        pass

    # Forzar aceptación del checkbox
    time.sleep(5)
    cb = page.locator("#contenido_cbAcepto")
    if not cb.is_checked():
        cb.check(force=True)
    print("✅ Checkbox marcado correctamente")

    # Clic en iniciar
    page.click("#contenido_btnIniciar", timeout=10_000)

    # Ahora esperamos: o bien carga Página 2, o bien aparece CAPTCHA
    try:
        page.wait_for_selector("#contenido_Wizard3_tbCedula", timeout=5000)
        print("✅ Página 2 cargada correctamente (sin CAPTCHA)")
        return True
    except TimeoutError:
        # No cargó página 2 aún → probablemente salió CAPTCHA
        return False

def validar_captcha_hibrido(page, max_intentos=3):
    for intento in range(max_intentos):
        if page.is_visible("#contenido_ucInfor_lbMensajeEnPopup"):
            print(f"🤖 CAPTCHA detectado. Intento {intento+1}/{max_intentos}...")
            try:
                # Cerrar modal (sin esperar navegación)
                page.click("#contenido_ucInfor_lbClose", timeout=5000)

                # Re-seleccionar selects y checkbox
                page.select_option("#contenido_ddlTipoSeleccion", "21")
                page.select_option("#contenido_ddlTipoDocumento", "1")

                # Forzar aceptación del checkbox
                cb = page.locator("#contenido_cbAcepto")
                cb.wait_for()
                if not cb.is_checked():
                    cb.check(force=True)
                print("✅ Checkbox marcado correctamente")

                # Dar clic en Continuar y esperar navegación real
                with page.expect_navigation(wait_until="networkidle", timeout=20_000):
                    page.click("#contenido_btnIniciar")

                # Esperar campo cédula hasta 30s
                page.wait_for_selector("#contenido_Wizard3_tbCedula", timeout=10_000)
                print("✅ Página 2 cargada correctamente tras CAPTCHA")
                return True

            except Exception as e:
                print(f"   (Error en intento: {e})")
                continue
        else:
            try:
                page.wait_for_selector("#contenido_Wizard3_tbCedula", timeout=30_000)
                print("✅ Página 2 cargada correctamente (sin CAPTCHA visible)")
                return True
            except TimeoutError:
                pass
    print("⚠️ CAPTCHA no se resolvió automáticamente")
    return False


# ==============================
# 📌 PÁGINA 2
# ==============================
def pagina2_cedula_correo(page, cedula: str, correo: str, max_retries=3):
    """
    Llena cédula y correo en la Página 2 (solicitud.aspx).
    Maneja timeouts y reintenta hasta `max_retries`.
    """
    for attempt in range(max_retries):
        try:
            # 🚩 Confirmar que cargó el formulario de Página 2
            page.wait_for_selector("#contenido_Wizard3_tbCedula", timeout=15_000)
            print("🌐 Página 2: Cédula - Correo")

            # Limpiar y escribir la cédula
            campo_cedula = page.locator("#contenido_Wizard3_tbCedula")
            campo_cedula.fill("")
            campo_cedula.type(cedula, delay=50)

            # Limpiar y escribir correo
            campo_correo = page.locator("#contenido_Wizard3_ucCorreoElectronico_tbCorreoElectronico")
            campo_correo.fill("")
            campo_correo.type(correo, delay=50)

            # Confirmación correo
            campo_confirm = page.locator("#contenido_Wizard3_ucCorreoElectronico_tbCorfirmCorreo")
            campo_confirm.fill("")
            campo_confirm.type(correo, delay=50)

            # Continuar → Página 3
            with page.expect_navigation(wait_until="networkidle", timeout=30_000):
                page.click("#contenido_Wizard3_StartNavigationTemplateContainerID_StartNextButton")

            return True

        except TimeoutError:
            print(f"⚠️ Timeout en Página 2 (intento {attempt+1}/{max_retries})")
            # Evitar navegación directa: solo reintentar
            continue

    return False

def retroceder_a_pagina2(page, max_intentos=10):
    """Vuelve dinámicamente a Página 2 mediante go_back."""
    for i in range(max_intentos):
        if page.is_visible("#contenido_Wizard3_tbCedula"):
            print("📄 Confirmado: estamos de nuevo en Página 2")
            return True
        try:
            page.go_back(wait_until="commit")
            print(f"↩️ Retroceso {i+1}/{max_intentos} completado")
        except Exception as e:
            print(f"⚠️ Error en retroceso {i+1}: {e}")
    print("⚠️ No se pudo confirmar que volvimos a Página 2")
    return False


# ==============================
# 📌 PÁGINA 3
# ==============================
def pagina3_checkboxes_fecha(page, fecha_expedicion: str):
    """
    Marca los checks, llena la fecha de expedición y maneja error de fecha inválida.
    Si la fecha es ambigua (ej: 8/6/1997), prueba ambos formatos (%m/%d/%Y y %d/%m/%Y).
    Retorna:
      (True, None)  -> si avanzó a página 4
      (False, msg)  -> si hubo error de fecha
    """
    print("🌐 Página 3: Check - Fecha Exp")

    page.wait_for_selector("#contenido_Wizard3_rbConFinMigratorio", timeout=12_000)

    # 1️⃣ Seleccionar "SI" en radio migratorio
    page.check("#contenido_Wizard3_rbConFinMigratorio", force=True)

    # 2️⃣ Marcar checkbox
    cb = page.locator("#contenido_Wizard3_cbInformacionReservada")
    for intento in range(3):
        try:
            cb.scroll_into_view_if_needed()
            cb.click(force=True)
            page.wait_for_timeout(1500)
            if cb.is_checked():
                print("✅ Checkbox 'Acepto' marcado correctamente")
                break
        except Exception as e:
            print(f"⚠️ Error al marcar checkbox (intento {intento+1}): {e}")
    else:
        raise Exception("❌ El checkbox 'Acepto' no pudo marcarse después de 3 intentos")

    # 3️⃣ Normalizar fecha de entrada
    fecha_expedicion = str(fecha_expedicion).strip().replace("-", "/")

    if len(fecha_expedicion) == 8 and "/" not in fecha_expedicion:
        fecha_expedicion = f"{fecha_expedicion[:2]}/{fecha_expedicion[2:4]}/{fecha_expedicion[4:]}"
        print(f"📅 Fecha formateada automáticamente como {fecha_expedicion}")

    # 4️⃣ Intentar con los dos posibles formatos
    posibles_formatos = ["%m/%d/%Y", "%d/%m/%Y"]
    ultimo_error = None

    for idx, formato in enumerate(posibles_formatos, start=1):
        try:
            fecha = datetime.strptime(fecha_expedicion, formato)
        except ValueError:
            # Fecha imposible en este formato (ej: mes 29)
            continue

        fecha_str_final = fecha.strftime("%d%m%Y")
        print(f"🧩 Probando formato {formato} → {fecha_str_final} (intento {idx})")

        # 👉 Digitar fecha
        campo_fecha = page.locator("#contenido_Wizard3_tbExpedicionCedula_tbFecha")
        campo_fecha.click()
        page.wait_for_timeout(400)
        campo_fecha.press("Control+A")
        campo_fecha.press("Delete")
        page.wait_for_timeout(300)

        for ch in fecha_str_final:
            page.keyboard.type(ch, delay=100)
        page.wait_for_timeout(800)

        # 👉 Intentar avanzar
        page.click("#contenido_Wizard3_StepNavigationTemplateContainerID_StepNextButton")

        try:
            page.wait_for_selector("#contenido_Wizard3_ucTramitePorPais_ddlPais", timeout=8000)
            print(f"✅ Fecha válida y aceptada ({formato})")
            return True, None

        except Exception:
            # 📌 Verificar si apareció el modal de fecha inválida
            modal_selector = "#contenido_ucInfor_panInformmacion"
            if page.is_visible(modal_selector):
                mensaje = page.inner_text("#contenido_ucInfor_lbMensajeEnPopup")
                print(f"❌ Error detectado con formato {formato}: {mensaje}")
                ultimo_error = mensaje

                # Intentar cerrar modal, pero SIN retroceder a Página 2 todavía
                try:
                    if page.is_visible("#contenido_ucInfor_lbClose"):
                        page.click("#contenido_ucInfor_lbClose", timeout=2000)
                        print("🔒 Modal cerrado para reintentar otro formato")
                    else:
                        print("ℹ️ El modal ya no estaba visible")
                except Exception as e:
                    print(f"⚠️ No se pudo cerrar el modal: {e}")

                # Reintentar siguiente formato SIN retroceder
                continue
            else:
                ultimo_error = "❌ No avanzó ni apareció modal"
                print(ultimo_error)
                continue

    # 🧱 Si llega aquí, ninguno de los formatos funcionó → retroceder recién ahora
    print("⚠️ Ningún formato de fecha fue aceptado — retrocediendo a Página 2")
    retroceder_a_pagina2(page)
    return False, ultimo_error or "⚠️ Validar fecha de expedición CC (ningún formato aceptado)"

# ==============================
# 📌 PÁGINA 4
# ==============================
def pagina4_seleccionar_pais(page, pais_value: str):
    """
    Selecciona el país (por value) y avanza a la página 5.
    Si aparece la alerta de la DIJIN, guarda 'Validar con dijin' y vuelve a Página 2.
    Retorna:
      (True, None)             -> si avanzó a página 5
      (False, "Validar con dijin")  -> si apareció la alerta y retrocedió
    """
    print("🌐 Página 4: Seleccionar país")

    select_locator = "#contenido_Wizard3_ucTramitePorPais_ddlPais"

    # Esperar a que exista el <select> en el DOM
    page.wait_for_selector(select_locator, state="attached", timeout=12_000)

    # Forzar selección y disparar onchange
    page.evaluate(
        """({ selector, value }) => {
            const sel = document.querySelector(selector);
            if (sel) {
                sel.value = value;
                sel.dispatchEvent(new Event("change", { bubbles: true }));
            }
        }""",
        {"selector": select_locator, "value": pais_value}
    )
    print(f"✅ País seleccionado (value={pais_value})")

    # Esperar al postback parcial
    page.wait_for_timeout(5000)

    # Click en "Continuar"
    page.click("#contenido_Wizard3_StepNavigationTemplateContainerID_StepNextButton")

    # Ahora esperar: o cargamos página 5 o aparece alerta DIJIN
    try:
        # Esperar a que cargue algo de página 5
        page.wait_for_selector("#contenido_Wizard3_rbSi", timeout=8_000)
        print("➡️ Avanzando a Página 5")
        return True, None

    except Exception:
        # 📌 Verificar si apareció la alerta DIJIN
        try:
            time.sleep(5)
            if page.is_visible("#contenido_ucInfor_panInformativo", timeout=3_000):
                mensaje = page.inner_text("#contenido_ucInfor_lblMensajes2")
                print(f"❌ Alerta detectada en Página 4: {mensaje}...")

                mensaje_simplificado = "Validar con dijin"

                # 🔄 Retroceder dinámicamente hasta Página 2
                retroceder_a_pagina2(page)

                return False, mensaje_simplificado
        except Exception:
            pass

        raise Exception("❌ Error desconocido en Página 4: no cargó Página 5 ni apareció alerta DIJIN")


# ==============================
# 📌 PÁGINA 5
# ==============================
def extraer_codigo_modal(page, modal_selector, mensaje_selector):
    """Extrae el código de solicitud del modal (si existe) y guarda screenshot."""
    codigo = None
    try:
        if page.is_visible(modal_selector):
            mensaje = page.inner_text(mensaje_selector)
            print(f"📝 Modal detectado: {mensaje[:120]}...")

            # Guardar screenshot con timestamp
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"modal_{ts}.png"
            page.screenshot(path=filename)
            print(f"📸 Screenshot guardado: {filename}")

            # Buscar número de solicitud que empiece con 52
            match = re.search(r"\b52\d+\b", mensaje)
            if match:
                codigo = match.group(0)
                print(f"✅ Código detectado en el modal: {codigo}")
            else:
                print("⚠️ No se encontró código en el modal")
    except Exception as e:
        print(f"⚠️ Error al leer modal: {e}")
    return codigo

def pagina5_confirmar_datos(page, max_reintentos=3):
    """
    Página 5: confirmar datos.
    Retorna:
      (codigo, False) -> si apareció modal con solicitud previa
      (None, True)    -> si avanzó correctamente a Página 6
      (None, False)   -> si no se pudo avanzar
    """
    print("🌐 Página 5: Confirmar datos")

    radio_si = page.locator("#contenido_Wizard3_rbSi")
    boton_continuar = page.locator("#contenido_Wizard3_StepNavigationTemplateContainerID_StepNextButton")
    modal_selector = "#contenido_ucInfor_panInformmacion"
    mensaje_selector = "#contenido_ucInfor_lbMensajeEnPopup"


    # --- Caso 2: Flujo normal ---
    try:
        radio_si.scroll_into_view_if_needed()
        radio_si.click(force=True)
        print("✅ Radio 'Sí' marcado")
    except Exception:
        page.evaluate("document.getElementById('contenido_Wizard3_rbSi').checked = true;")
        print("⚠️ Radio 'Sí' forzado por JS")

    # --- Intentar avanzar a Página 6 ---
    for intento in range(1, max_reintentos + 1):
        try:
            print(f"➡️ Click en 'Continuar' (intento {intento})")
            if not radio_si.is_checked():
                radio_si.scroll_into_view_if_needed()
                radio_si.click(force=True)
                print("✅ Radio 'Sí' marcado antes de CONTINUAR")
            else:
                print("ℹ️ Radio 'Sí' ya estaba marcado")
            boton_continuar.click()
            # Validar que la URL cambie a página 6
            page.wait_for_url(
                re.compile(r".*/capturaDatosPagos\.aspx.*"),
                timeout=8_000
            )
            print("✅ Avanzamos correctamente a Página 6")
            return None, True
        except TimeoutError:
            print(f"⚠️ No se avanzó a Página 6 en intento {intento}")
            if intento < max_reintentos:
                time.sleep(2)
            else:
                print("❌ Error: no se pudo avanzar a Página 6 después de varios intentos")
                return None, False



# ==============================
# 📌 PÁGINA 6
# ==============================
def pagina6_codigo(page):
    """Captura el número de solicitud en la página 6 y retrocede hasta la página 2."""
    print("🌐 Página 6: Capturando número de solicitud...")

    candidatos = [
        "#contenido_Wizard2_infoNumeroSolicitud_lblMensajes2",
        "#contenido_Wizard3_infoNumeroSolicitud_lblMensajes2",
        "#contenido_Wizard2_lblMensajes2",
        "#contenido_Wizard3_lblMensajes2",
    ]

    selector_mensaje = None
    for s in candidatos:
        if page.is_visible(s):
            selector_mensaje = s
            break
    if not selector_mensaje:
        for s in candidatos:
            try:
                page.wait_for_selector(s, timeout=8_000)
                selector_mensaje = s
                break
            except TimeoutError:
                continue

    codigo = None
    if selector_mensaje:
        for intento in range(5):
            try:
                mensaje = page.inner_text(selector_mensaje).strip()
                match = re.search(r"\b52\d+\b", mensaje)
                if match:
                    codigo = match.group(0)
                    print(f"✅ Código de solicitud encontrado en intento {intento+1}: {codigo}")
                    break
            except Exception as e:
                print(f"⚠️ No se pudo leer el mensaje en intento {intento+1}: {e}")
            time.sleep(1)
    else:
        print("⚠️ No se detectó el contenedor del código en Página 6. Puede que haya cambiado el ID.")

    if not codigo:
        print("⚠️ No se pudo capturar código en Página 6.")

    # ✅ Reutilizamos la función para volver a página 2
    retroceder_a_pagina2(page)
    return codigo


# ==============================
# 📌 PROCESAR UNA PERSONA
# ==============================
def procesar_persona(page, cedula, correo, fecha_expedicion, pais_value="173"):
    """
    Procesa una fila del Excel desde Página 2 en adelante.
    Asume que ya se creó la solicitud y se pasó Página 1 con CAPTCHA válido.
    """
    # Página 2
    if not pagina2_cedula_correo(page, cedula, correo):
        return None, "No se pudo completar la página 2"

    # Página 3
    ok, mensaje = pagina3_checkboxes_fecha(page, fecha_expedicion)
    if not ok:
        return None, mensaje

    # Página 4
    ok, mensaje = pagina4_seleccionar_pais(page, pais_value)
    if not ok:
        return None, mensaje

    # Página 5
    codigo, continuar = pagina5_confirmar_datos(page)
    if not continuar:
        return codigo, None

    # Página 6
    codigo = pagina6_codigo(page)
    if codigo:
        return codigo, None
    else:
        return None, "No se detectó código en Página 6"

# ==============================
# 📌 MAIN
# ==============================
def main():
    correo = "apostillamen@gmail.com"
    pais_value = "173"

    df = leer_excel("entrada.xlsx")

    # Asegurar que columnas existan
    if "OBSERVACIONES" not in df.columns:
        df["OBSERVACIONES"] = ""
    if "CODIGO" not in df.columns:
        df["CODIGO"] = ""

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=50)
        page = browser.new_page()

        try:
            # 🌐 Página 1: Crear Solicitud
            pagina1_inicio(page)

            # Intentar CAPTCHA automáticamente
            captcha_ok = validar_captcha_hibrido(page)

            if not captcha_ok:
                print("⚠️ CAPTCHA requiere intervención manual...")
                try:
                    page.wait_for_selector("#contenido_ucInfor_lbClose", timeout=15_000)
                    page.click("#contenido_ucInfor_lbClose")
                    print("✅ Modal de CAPTCHA cerrado manualmente")
                    time.sleep(2)
                except:
                    print("⚠️ No apareció modal de error de captcha, seguimos...")

                cb = page.locator("#contenido_cbAcepto")
                if not cb.is_checked():
                    cb.check(force=True)

                boton = page.locator("#contenido_btnIniciar")
                boton.wait_for(state="visible", timeout=5000)
                with page.expect_navigation(wait_until="networkidle", timeout=30_000):
                    boton.click()

                print("✅ Página 2 cargada manualmente")
            else:
                print("✅ Página 2 cargada automáticamente tras pasar el captcha")

            # 🔄 Recorrido de filas
            for i, row in df.iterrows():
                nombre = str(row["NOMBRE"])
                cedula = str(row["CEDULA"])
                fecha_expedicion = row["FECHA_EXP"]

                # ✅ Validar si ya existe un código en la fila
                if pd.notna(row["CODIGO"]) and str(row["CODIGO"]).strip() != "":
                    print(f"\n⏭️ Fila {i+1} ({nombre}) ya tiene código: {row['CODIGO']} -> se omite")
                    continue

                print(f"\nFila {i+1} Procesando..\n👤 {nombre} - cédula {cedula} ({i+1}/{len(df)})")

                try:
                    # Formatear fecha
                    try:
                        fecha_str = pd.to_datetime(fecha_expedicion).strftime("%d%m%Y")
                    except Exception:
                        df.at[i, "OBSERVACIONES"] = "Fecha inválida - Validar formato dd/mm/aa"
                        df.at[i, "CODIGO"] = ""
                        continue

                    # Procesar desde Página 2 en adelante
                    codigo, observacion = procesar_persona(
                        page, cedula, correo, fecha_str, pais_value
                    )

                    if observacion:
                        df.at[i, "OBSERVACIONES"] = observacion
                        df.at[i, "CODIGO"] = ""
                    else:
                        df.at[i, "OBSERVACIONES"] = ""
                        df.at[i, "CODIGO"] = str(codigo)

                except Exception as e:
                    print(f"❌ Error procesando fila {i+1}: {e}")
                    df.at[i, "OBSERVACIONES"] = f"Error inesperado: {str(e)}"
                    df.at[i, "CODIGO"] = ""

                # 💾 Guardar progreso después de cada fila
                guardar_excel(df, "entrada.xlsx")

        finally:
            browser.close()


if __name__ == "__main__":
    main()
