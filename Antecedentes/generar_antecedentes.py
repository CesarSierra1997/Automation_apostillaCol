import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError
import re
import random
import time
import asyncio
import sys
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

    page.wait_for_selector("#contenido_ddlTipoDocumento", timeout=3_000)
    page.select_option("#contenido_ddlTipoDocumento", "1")

    # Cerrar modal si aparece
    try:
        page.wait_for_selector("#contenido_ucInfor_panInformmacion", timeout=5_000)
        page.click("#contenido_ucInfor_lbClose")
    except TimeoutError:
        pass

    # Forzar aceptación del checkbox
    cb = page.locator("#contenido_cbAcepto")
    if not cb.is_checked():
        cb.check(force=True)
    print("✅ Checkbox marcado correctamente")

    # Clic en iniciar
    page.click("#contenido_btnIniciar", timeout=5_000)

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
                page.click("#contenido_ucInfor_lbClose", timeout=2000)

                # Re-seleccionar selects y checkbox
                page.select_option("#contenido_ddlTipoSeleccion", "21")
                page.select_option("#contenido_ddlTipoDocumento", "1")

                cb = page.locator("#contenido_cbAcepto")
                if not cb.is_checked():
                    cb.check(force=True)

                # Dar clic en Continuar y esperar navegación real
                with page.expect_navigation(wait_until="networkidle", timeout=40_000):
                    page.click("#contenido_btnIniciar")

                # Esperar campo cédula hasta 30s
                page.wait_for_selector("#contenido_Wizard3_tbCedula", timeout=30_000)
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

            # Intentar recuperar manualmente
            try:
                page.goto(
                    "https://tramites.cancilleria.gov.co/apostillalegalizacion/PolNal/solicitud.aspx",
                    wait_until="networkidle",
                    timeout=30_000
                )
                continue
            except Exception as e:
                print(f"⚠️ No se pudo recuperar Página 2: {e}")
                continue

    return False


# ==============================
# 📌 PÁGINA 3
# ==============================
def pagina3_checkboxes_fecha(page, fecha_expedicion: str):
    """
    Marca los checks, llena fecha de expedición y maneja error de fecha inválida.
    Retorna:
      (True, None)  -> si avanzó a página 4
      (False, msg)  -> si hubo error de fecha (msg = texto de error)
    """
    print("🌐 Página 3: Check - Fecha Exp")

    page.wait_for_selector("#contenido_Wizard3_rbConFinMigratorio", timeout=12_000)

    # 1. Seleccionar "SI" en radio migratorio
    page.check("#contenido_Wizard3_rbConFinMigratorio", force=True)

    # 2. Marcar checkbox
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

    # 3. Llenar fecha
    campo_fecha = page.locator("#contenido_Wizard3_tbExpedicionCedula_tbFecha")
    campo_fecha.click()
    page.wait_for_timeout(500)
    campo_fecha.press("Control+A")
    campo_fecha.press("Delete")
    page.wait_for_timeout(500)

    for ch in fecha_expedicion:  # ej: "02122015"
        page.keyboard.type(ch, delay=100)

    page.wait_for_timeout(1500)

    # 4. Intentar continuar
    page.click("#contenido_Wizard3_StepNavigationTemplateContainerID_StepNextButton")

    # Esperar: o aparece la página 4, o aparece el modal de error
    try:
        page.wait_for_selector("#contenido_Wizard3_ucTramitePorPais_ddlPais", timeout=8_000)
        print("➡️ Avanzando a Página 4")
        return True, None

    except Exception:
        # 📌 Verificar si apareció el modal de fecha inválida
        modal_selector = "#contenido_ucInfor_panInformmacion"
        if page.is_visible(modal_selector):
            mensaje = page.inner_text("#contenido_ucInfor_lbMensajeEnPopup")
            print(f"❌ Error detectado en Página 3: {mensaje}")

            # Intentar cerrar modal si aún está visible
            try:
                if page.is_visible("#contenido_ucInfor_lbClose"):
                    page.click("#contenido_ucInfor_lbClose", timeout=2000)
                    print("🔒 Modal cerrado")
                else:
                    print("ℹ️ El modal ya no estaba visible")
            except Exception as e:
                print(f"⚠️ No se pudo cerrar el modal: {e}")

            # 🔄 Retroceder dinámicamente hasta Página 2
            for i in range(6):  # máximo 6 retrocesos
                if page.is_visible("#contenido_Wizard3_tbCedula"):
                    print("📄 Confirmado: estamos de nuevo en Página 2")
                    break
                try:
                    page.go_back(wait_until="commit")
                    print(f"↩️ Retroceso {i+1}/6 completado")
                except Exception as e:
                    print(f"⚠️ Error en retroceso {i+1}: {e}")
            else:
                print("⚠️ No se pudo confirmar que volvimos a Página 2")

            return False, mensaje
        else:
            raise Exception("❌ Error desconocido: no cargó Página 4 ni apareció el modal de error")


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
    page.wait_for_timeout(2000)

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
            if page.is_visible("#contenido_ucInfor_panInformativo", timeout=3_000):
                mensaje = page.inner_text("#contenido_ucInfor_lblMensajes2")
                print(f"❌ Alerta detectada en Página 4: {mensaje}...")

                mensaje_simplificado = "Validar con dijin"

                # 🔄 Retroceder dinámicamente hasta Página 2
                for i in range(6):  # máximo 6 retrocesos
                    if page.is_visible("#contenido_Wizard3_tbCedula"):
                        print("📄 Confirmado: estamos de nuevo en Página 2")
                        break
                    try:
                        page.go_back(wait_until="commit")
                        print(f"↩️ Retroceso {i+1}/6 completado")
                    except Exception as e:
                        print(f"⚠️ Error en retroceso {i+1}: {e}")
                else:
                    print("⚠️ No se pudo confirmar que volvimos a Página 2")

                return False, mensaje_simplificado
        except Exception:
            pass

        raise Exception("❌ Error desconocido en Página 4: no cargó Página 5 ni apareció alerta DIJIN")


# ==============================
# 📌 PÁGINA 5
# ==============================
def pagina5_confirmar_datos(page):
    """Confirma los datos y maneja el modal de solicitud existente.
    Retorna:
      (codigo, False) -> si apareció el modal con solicitud existente
      (None, True)    -> si avanzó al flujo normal
    """
    print("🌐 Página 5: Confirmar datos")

    radio_si = page.locator("#contenido_Wizard3_rbSi")
    boton_continuar = page.locator("#contenido_Wizard3_StepNavigationTemplateContainerID_StepNextButton")
    modal_selector = "#contenido_ucInfor_panInformmacion"
    mensaje_selector = "#contenido_ucInfor_lbMensajeEnPopup"

    # 1. Marcar el radio 'Sí' y verificar su estado
    try:
        radio_si.scroll_into_view_if_needed()
        # En lugar de .check(), usamos .click() y forzamos el estado con JavaScript
        radio_si.click(force=True)
        # Forzar el estado 'checked' directamente en el DOM
        page.evaluate("document.getElementById('contenido_Wizard3_rbSi').checked = true;")
        print("✅ Radio 'Sí' marcado y su estado forzado a 'checked'")
        
        # Esperar un breve momento por si hay eventos asíncronos en la página
        time.sleep(2)
    except Exception as e:
        # Aquí capturamos cualquier error en el clic o en la evaluación
        raise Exception(f"❌ No se pudo marcar o forzar el radio 'Sí': {e}")

    # 2. NUEVA LÓGICA: Verificar si ya apareció el modal ANTES de hacer clic en continuar
    if page.is_visible(modal_selector, timeout=3_000):
        print("📝 Modal detectado inmediatamente después de marcar radio 'Sí'")
        mensaje = page.inner_text(mensaje_selector)
        print(f"📝 Mensaje: {mensaje[:120]}...")

        # Extraer número de solicitud (empieza con 52)
        match = re.search(r"\b52\d+\b", mensaje)
        codigo = match.group(0) if match else None
        if codigo:
            print(f"✅ Código de solicitud encontrado: {codigo}")
        else:
            print("⚠️ No se encontró código en el modal")

        # 🔄 Retroceder dinámicamente hasta Página 2
        for i in range(7):
            if page.is_visible("#contenido_Wizard3_tbCedula"):
                print("📄 Confirmado: estamos de nuevo en Página 2")
                break
            try:
                page.go_back(wait_until="commit")
                print(f"↩️ Retroceso {i+1}/7 completado")
            except Exception as e:
                print(f"⚠️ Error en retroceso {i+1}: {e}")
        else:
            print("⚠️ No se pudo confirmar que volvimos a Página 2")

        return codigo, False

    # 3. Si no hay modal, proceder con el clic en continuar
    try:
        boton_continuar.click()

        # Esperar explícitamente a un elemento de la página 6
        page.wait_for_selector("#contenido_Wizard2_infoNumeroSolicitud_lblMensajes2", timeout=30_000)
        print("✅ Navegación a Página 6 confirmada, continuando con el flujo normal")
        return None, True
    except Exception as e:
        print(f"⚠️ Fallo en la navegación a Página 6: {e}. Validando si apareció un modal...")

    # =========================
    # 4. Verificar si apareció modal DESPUÉS del clic (caso de respaldo)
    # =========================
    if page.is_visible(modal_selector):
        mensaje = page.inner_text(mensaje_selector)
        print(f"📝 Modal detectado después del clic: {mensaje[:120]}...")

        # Extraer número de solicitud (empieza con 52)
        match = re.search(r"\b52\d+\b", mensaje)
        codigo = match.group(0) if match else None
        if codigo:
            print(f"✅ Código de solicitud encontrado: {codigo}")
        else:
            print("⚠️ No se encontró código en el modal")

        # 🔄 Retroceder dinámicamente hasta Página 2
        for i in range(7):
            if page.is_visible("#contenido_Wizard3_tbCedula"):
                print("📄 Confirmado: estamos de nuevo en Página 2")
                break
            try:
                page.go_back(wait_until="commit")
                print(f"↩️ Retroceso {i+1}/7 completado")
            except Exception as e:
                print(f"⚠️ Error en retroceso {i+1}: {e}")
        else:
            print("⚠️ No se pudo confirmar que volvimos a Página 2")

        return codigo, False
    else:
        print("❌ La página no navegó y no se detectó un modal. Algo inesperado ocurrió.")
        return None, False

# ==============================
# 📌 PÁGINA 6
# ==============================
def pagina6_codigo(page):
    """Captura el número de solicitud en la página 6 y retrocede hasta la página 2."""
    print("🌐 Página 6: Capturando número de solicitud...")

    selector_mensaje = "#contenido_Wizard2_infoNumeroSolicitud_lblMensajes2"

    # Intentar leer el mensaje varias veces
    codigo = None
    for intento in range(5):
        try:
            mensaje = page.inner_text(selector_mensaje).strip()
            # Buscar un número de 10 dígitos o más que empiece con '52'
            match = re.search(r"\b52\d{9,}\b", mensaje)
            if match:
                codigo = match.group(0)
                print(f"✅ Código de solicitud encontrado en intento {intento+1}: {codigo}")
                break
        except:
            pass
        time.sleep(1)

    if not codigo:
        print("⚠️ No se pudo capturar código en Página 6, continuando con retrocesos igualmente.")

    # Retroceder hasta Página 2
    for i in range(10):
        if page.is_visible("#contenido_Wizard3_tbCedula"):
            print("📄 Confirmación: regresamos correctamente a Página 2")
            break
        try:
            page.go_back(wait_until="commit")
            print(f"↩️ Retroceso {i+1}/10 realizado")
        except Exception as e:
            print(f"⚠️ Problema en retroceso {i+1}: {e}")
    else:
        print("❌ No fue posible regresar a Página 2 después de Página 6")

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

    # Asegurar que columnas existan y sean tipo string
    if "OBSERVACIONES" not in df.columns:
        df["OBSERVACIONES"] = ""
    if "CODIGO" not in df.columns:
        df["CODIGO"] = ""

    df["OBSERVACIONES"] = df["OBSERVACIONES"].astype(str)
    df["CODIGO"] = df["CODIGO"].astype(str)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=50)
        page = browser.new_page()

        try:
            # 🌐 Página 1: Crear Solicitud
            pagina1_inicio(page)

            # Intentar CAPTCHA automáticamente
            captcha_ok = validar_captcha_hibrido(page)

            if not captcha_ok:
                print("⚠️ CAPTCHA requiere intervención manual. Espera que el usuario lo resuelva...")

                # Esperar a que aparezca y cerrar modal de error si sale
                try:
                    page.wait_for_selector("#contenido_ucInfor_lbClose", timeout=15_000)
                    page.click("#contenido_ucInfor_lbClose")
                    print("✅ Modal de CAPTCHA cerrado manualmente")
                    time.sleep(2)
                except:
                    print("⚠️ No apareció modal de error de captcha, seguimos...")

                # Reforzar checkbox antes de continuar
                cb = page.locator("#contenido_cbAcepto")
                if not cb.is_checked():
                    cb.check(force=True)

                # Dar clic en continuar
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

                print(f"\nFila {i+1} Procesando..\n👤 {nombre} - cédula {cedula} ({i+1}/{len(df)})")

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

            guardar_excel(df, "entrada.xlsx")

        finally:
            browser.close()


if __name__ == "__main__":
    main()
