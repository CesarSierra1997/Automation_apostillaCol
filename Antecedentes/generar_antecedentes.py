#!/usr/bin/env python3
import random
import time
import re
from playwright.sync_api import sync_playwright, TimeoutError

BASE_URL = "https://tramites.cancilleria.gov.co/apostillalegalizacion/solicitud/inicio.aspx"


# ==============================
# 📌 PÁGINA 1
# ==============================
def pagina1_inicio(page):
    """Selecciona tipo de trámite y acepta términos iniciales"""
    page.goto(BASE_URL, timeout=15_000)
    print("🌐 Página 1: Crear Solicitud")

    page.wait_for_selector("#contenido_ddlTipoSeleccion", timeout=10_000)
    page.select_option("#contenido_ddlTipoSeleccion", "21")

    page.wait_for_selector("#contenido_ddlTipoDocumento", timeout=10_000)
    page.select_option("#contenido_ddlTipoDocumento", "1")

    # Cerrar modal si aparece
    try:
        page.wait_for_selector("#contenido_ucInfor_panInformmacion", timeout=5000)
        page.click("#contenido_ucInfor_lbClose")
    except TimeoutError:
        pass

    # Forzar aceptación del checkbox
    cb = page.locator("#contenido_cbAcepto")
    if not cb.is_checked():
        cb.check(force=True)
    print("✅ Checkbox marcado correctamente")

def validar_captcha(page, attempt):
    """Detecta y simula clics para intentar evitar reCAPTCHA"""
    if page.locator("#contenido_ucInfor_lbMensajeEnPopup:has-text('reCaptcha')").is_visible():
        print(f"🤖 CAPTCHA detectado. Intento {attempt+1}. Simulando clics...")
        try:
            page.click("#contenido_ucInfor_lbClose", timeout=2000)
            time.sleep(random.uniform(0.5, 1.2))
            page.click('h1:has-text("Apostilla")', timeout=2000)
            time.sleep(random.uniform(0.4, 1.0))
            page.click('p:has-text("correo electrónico")', timeout=2000)
            time.sleep(random.uniform(0.4, 1.0))
        except Exception as e:
            print(f"   (No se pudo simular clics: {e})")
        return True
    return False


# ==============================
# 📌 PÁGINA 2
# ==============================
def pagina2_cedula_correo(page, cedula: str, correo: str, max_retries=3):
    """Llena cédula y correo en la página 2"""
    for attempt in range(max_retries):
        try:
            with page.expect_navigation(wait_until="networkidle", timeout=15_000):
                page.click("#contenido_btnIniciar")

            # Esperar campo de cédula
            page.wait_for_selector("#contenido_Wizard3_tbCedula", timeout=12_000)
            print("🌐 Página 2: Cedula - Correo")

            page.fill("#contenido_Wizard3_tbCedula", cedula)
            page.fill("#contenido_Wizard3_ucCorreoElectronico_tbCorreoElectronico", correo)
            page.fill("#contenido_Wizard3_ucCorreoElectronico_tbCorfirmCorreo", correo)

            # Continuar (esperando navegación a la Página 3)
            with page.expect_navigation(wait_until="networkidle", timeout=15_000):
                page.click("#contenido_Wizard3_StartNavigationTemplateContainerID_StartNextButton")
            return True

        except TimeoutError:
            if validar_captcha(page, attempt):
                continue
            else:
                print("❌ No se pudo avanzar a la página 2")
                return False
    return False


# ==============================
# 📌 PÁGINA 3
# ==============================
def pagina3_checkboxes_fecha(page, fecha_expedicion: str):
    """Marca los checks y llena fecha de expedición de cédula"""
    print("🌐 Página 3: Check - Fecha Exp")

    page.wait_for_selector("#contenido_Wizard3_rbConFinMigratorio", timeout=12_000)

    # 1. Seleccionar "SI" en radio migratorio
    page.check("#contenido_Wizard3_rbConFinMigratorio", force=True)

    # 2. Marcar checkbox con varios intentos
    cb = page.locator("#contenido_Wizard3_cbInformacionReservada")

    for intento in range(3):
        try:
            cb.scroll_into_view_if_needed()
            cb.click(force=True)
            page.wait_for_timeout(1500)  # darle tiempo al postback

            if cb.is_checked():
                print("✅ Checkbox 'Acepto' marcado correctamente")
                break
        except Exception as e:
            print(f"⚠️ Error al marcar checkbox (intento {intento+1}): {e}")
    else:
        raise Exception("❌ El checkbox 'Acepto' no pudo marcarse después de 3 intentos")

    # 3. Fecha de expedición (solo números → máscara agrega '/')
    campo_fecha = page.locator("#contenido_Wizard3_tbExpedicionCedula_tbFecha")
    campo_fecha.click()
    page.wait_for_timeout(500)

    campo_fecha.press("Control+A")
    campo_fecha.press("Delete")
    page.wait_for_timeout(500)

    for ch in fecha_expedicion:  # ej: "02122015"
        page.keyboard.type(ch, delay=100)

    page.wait_for_timeout(1500)

    # 4. Continuar a página 4
    with page.expect_navigation(wait_until="networkidle", timeout=15_000):
        page.click("#contenido_Wizard3_StepNavigationTemplateContainerID_StepNextButton")

    print("➡️ Avanzando a Página 4")
    

# ==============================
# 📌 PÁGINA 4
# ==============================
def pagina4_seleccionar_pais(page, pais_value: str):
    """Selecciona el país (por value) y avanza a la página 5"""
    print("🌐 Página 4: Seleccionar país")

    select_locator = "#contenido_Wizard3_ucTramitePorPais_ddlPais"

    # Esperar a que exista el <select> en el DOM (aunque esté oculto)
    page.wait_for_selector(select_locator, state="attached", timeout=12_000)

    # Forzar selección y disparar el evento onchange
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
    with page.expect_navigation(wait_until="networkidle", timeout=15_000):
        page.click("#contenido_Wizard3_StepNavigationTemplateContainerID_StepNextButton")

    print("➡️ Avanzando a Página 5")


# ==============================
# 📌 PÁGINA 5
# ==============================
def pagina5_confirmar_datos(page):
    """Confirma los datos y maneja el modal de solicitud existente"""
    print("🌐 Página 5: Confirmar datos")

    radio_si = page.locator("#contenido_Wizard3_rbSi")

    # Intentar marcar el radio "Sí"
    try:
        radio_si.scroll_into_view_if_needed()
        radio_si.click(force=True)
        page.wait_for_timeout(1000)
        print("✅ Radio 'Sí' marcado correctamente")
    except Exception as e:
        raise Exception(f"❌ No se pudo marcar el radio 'Sí': {e}")

    # Clic en continuar
    try:
        with page.expect_navigation(wait_until="networkidle", timeout=15_000):
            page.click("#contenido_Wizard3_StepNavigationTemplateContainerID_StepNextButton")
    except Exception:
        print("⚠️ No hubo navegación después de continuar (posible modal)")

    # =========================
    # Verificar si apareció modal
    # =========================
    modal_selector = "#contenido_ucInfor_panInformmacion"
    mensaje_selector = "#contenido_ucInfor_lbMensajeEnPopup"
    close_selector = "#contenido_ucInfor_lbClose"

    if page.is_visible(modal_selector, timeout=3000):
        mensaje = page.inner_text(mensaje_selector)
        print(f"📝 Modal detectado: {mensaje[:120]}...")

        # Extraer número de solicitud (empieza con 52)
        match = re.search(r"\b52\d+\b", mensaje)
        codigo = match.group(0) if match else None
        if codigo:
            print(f"✅ Código de solicitud encontrado: {codigo}")
        else:
            print("⚠️ No se encontró código en el modal")

        # Cerrar modal
        page.click(close_selector)
        page.wait_for_timeout(1000)
        print("❌ Modal cerrado")

        # Retroceder 7 veces en historial (volver a Página 2)
        for i in range(7):
            page.go_back(wait_until="networkidle")
            print(f"↩️ Retroceso {i+1}/7 completado")

        return codigo, False  # False = no continuar flujo normal
    else:
        print("✅ No apareció modal, continuar flujo normal")
        return None, True

# ==============================
# 📌 PÁGINA 6
# ==============================
def pagina6_codigo(page):
    """Extrae el codigo encontrado"""
    print()



# ==============================
# 📌 MAIN
# ==============================
def main():
    cedula = "1073249475"
    correo = "apostillamen@gmail.com"
    fecha_expedicion = "02122015"  # Formato solo números
    pais_value = "173"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=10)
        page = browser.new_page()

        try:
            # Páginas 1 a 4
            pagina1_inicio(page)
            if not pagina2_cedula_correo(page, cedula, correo):
                raise Exception("No se pudo completar la página 2")

            pagina3_checkboxes_fecha(page, fecha_expedicion)
            pagina4_seleccionar_pais(page, pais_value)

            # Página 5
            codigo, continuar = pagina5_confirmar_datos(page)

            if not continuar:
                print(f"🔁 Solicitud previa detectada (código: {codigo}). Se volvió a Página 2.")
            else:
                print("➡️ No hubo solicitud previa, se puede continuar con el flujo normal.")

        except Exception as e:
            print(f"❌ Error general: {e}")

        input("Presiona ENTER para cerrar el navegador...")
        browser.close()


if __name__ == "__main__":
    main()
