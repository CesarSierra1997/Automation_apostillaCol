import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError
import time
import sys

BASE_URL = "https://tramites.cancilleria.gov.co/apostillalegalizacion/consulta/tramite.aspx"
CORREO = "APOSTILLAMEN@GMAIL.COM"

# Datos constantes para el pago
COD_LIQUIDACION = "2020"
NOMBRE_COMPLETO = "GIOVANI AYALA"
TELEFONO = "3213337777"


# ==============================
# 📌 FUNCIONES DE EXCEL
# ==============================
def leer_excel(path="entrada.xlsx"):
    """Lee el Excel de entrada y devuelve DataFrame validado"""
    
    # Definimos las columnas esperadas en el orden correcto
    COLUMNAS_ESPERADAS = ["#", "CODIGO", "OBSERVACIONES"]

    # Leemos el archivo
    df = pd.read_excel(path, dtype={"CODIGO": str})

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

    # Limpiamos la columna CODIGO por si quedó algún .0
    df["CODIGO"] = df["CODIGO"].str.replace(r"\.0$", "", regex=True)

    return df


def guardar_excel(df, path="entrada.xlsx"):
    """Guarda los resultados en el mismo archivo (sin alterar el orden de columnas)"""
    df.to_excel(path, index=False)
    print(f"💾 Resultados actualizados en {path}")


# ==============================
# 📌 PÁGINA 1 - CONSULTA DE TRÁMITE
# ==============================
def pagina1_consulta(page, codigo, correo, max_intentos=3):
    """
    Ingresa código y correo en la página de consulta.
    Maneja el CAPTCHA con reintentos automáticos.
    Retorna: (True, None) si exitoso, (False, mensaje_error) si falla
    """
    
    for intento in range(max_intentos):
        print(f"🌐 Página 1: Consulta de Trámite (Intento {intento + 1}/{max_intentos})")
        
        try:
            # Navegar a la página
            if intento == 0:
                page.goto(BASE_URL, timeout=60_000)
            else:
                print(f"   ⏳ Reintentando después de error de CAPTCHA...")
                time.sleep(2)
            
            # Esperar que cargue el formulario
            page.wait_for_selector("#contenido_tbNumeroSolicitud", timeout=15_000)
            
            # Limpiar y llenar el campo de código
            page.fill("#contenido_tbNumeroSolicitud", "")
            page.fill("#contenido_tbNumeroSolicitud", codigo)
            print(f"   ✏️ Código ingresado: {codigo}")
            
            # Limpiar y llenar el campo de correo
            page.fill("#contenido_ucCorreoElectronico_tbCorreoElectronico", "")
            page.fill("#contenido_ucCorreoElectronico_tbCorreoElectronico", correo)
            print(f"   ✏️ Correo ingresado: {correo}")
            
            # Esperar 5 segundos antes de continuar
            print("   ⏳ Esperando 5 segundos...")
            time.sleep(5)
            
            # Dar clic en Continuar y esperar navegación
            try:
                with page.expect_navigation(wait_until="networkidle", timeout=20_000):
                    page.click("#contenido_btnBuscar")
                print("   ✅ Clic en Continuar ejecutado")
            except TimeoutError:
                # A veces no hay navegación si sale error
                time.sleep(2)
            
            # Verificar si apareció el error de CAPTCHA
            time.sleep(2)
            if page.is_visible("#contenido_validadorCaptcha"):
                print(f"   🤖 Error de CAPTCHA detectado en intento {intento + 1}")
                if intento < max_intentos - 1:
                    continue
                else:
                    return False, "Error de CAPTCHA - Máximo de intentos alcanzado"
            
            # Verificar si hubo otros errores visibles
            try:
                # Buscar mensajes de error comunes
                error_selectors = [
                    ".error_validacion",
                    ".alert-danger"
                ]
                
                for selector in error_selectors:
                    if page.is_visible(selector):
                        error_text = page.locator(selector).text_content()
                        if error_text and error_text.strip() and "Requerido" not in error_text:
                            return False, f"Error en formulario: {error_text.strip()}"
            except:
                pass
            
            # Si llegamos aquí sin errores, asumimos éxito
            print("   ✅ Página 1 completada exitosamente")
            time.sleep(2)
            return True, None
            
        except TimeoutError as e:
            print(f"   ⚠️ Timeout en intento {intento + 1}: {e}")
            if intento < max_intentos - 1:
                continue
            else:
                return False, "Timeout al cargar la página de consulta"
        
        except Exception as e:
            print(f"   ❌ Error inesperado en intento {intento + 1}: {e}")
            if intento < max_intentos - 1:
                continue
            else:
                return False, f"Error inesperado: {str(e)}"
    
    return False, "No se pudo completar la consulta después de todos los intentos"


# ==============================
# 📌 PÁGINA 2 - CONFIRMACIÓN DE DATOS
# ==============================
def pagina2_confirmacion(page):
    """
    Confirma que los datos están correctos marcando 'Sí' y continuando.
    Retorna: (True, None) si exitoso, (False, mensaje_error) si falla
    """
    print("🌐 Página 2: Confirmación de datos")
    
    try:
        # Esperar que cargue el radio button
        page.wait_for_selector("#contenido_rbSi", timeout=15_000)
        
        # Marcar el radio button "Sí"
        page.click("#contenido_rbSi")
        print("   ✅ Radio button 'Sí' marcado")
        
        # Esperar 5 segundos
        print("   ⏳ Esperando 5 segundos...")
        time.sleep(5)
        
        # Dar clic en Continuar
        try:
            with page.expect_navigation(wait_until="networkidle", timeout=20_000):
                page.click("#contenido_btnPagar")
            print("   ✅ Clic en Continuar ejecutado")
        except TimeoutError:
            time.sleep(2)
        
        # Verificar si hay error de validación
        time.sleep(2)
        if page.is_visible("#contenido_cvConfirmarTramite"):
            error_text = page.locator("#contenido_cvConfirmarTramite").text_content()
            return False, f"Error de validación: {error_text}"
        
        print("   ✅ Página 2 completada exitosamente")
        return True, None
        
    except TimeoutError:
        return False, "Timeout al cargar página de confirmación"
    
    except Exception as e:
        return False, f"Error en página 2: {str(e)}"


# ==============================
# 📌 PÁGINA 3 - SELECCIÓN DE MEDIO DE PAGO
# ==============================
def pagina3_medio_pago(page):
    """
    Selecciona medio de pago presencial y banco Sudameris.
    Retorna: (True, None) si exitoso, (False, mensaje_error) si falla
    """
    print("🌐 Página 3: Selección de medio de pago")
    
    try:
        # Esperar que cargue la página de pagos
        page.wait_for_selector("#contenido_Wizard2_rbExterior", timeout=15_000)
        
        # Marcar radio button de pago presencial
        page.click("#contenido_Wizard2_rbExterior")
        print("   ✅ Pago presencial marcado")
        
        # Esperar 3 segundos
        print("   ⏳ Esperando 3 segundos...")
        time.sleep(3)
        
        # Seleccionar Banco Sudameris
        page.select_option("#contenido_Wizard2_ddlPagoEn", "1")
        print("   ✅ Banco Sudameris seleccionado")
        
        time.sleep(2)
        
        # Dar clic en Continuar
        try:
            page.click("#contenido_Wizard2_StepNavigationTemplateContainerID_StepNextButton")
            time.sleep(3)
            print("   ✅ Clic en Continuar ejecutado")
        except:
            pass
        
        print("   ✅ Página 3 completada exitosamente")
        return True, None
        
    except TimeoutError:
        return False, "Timeout al cargar página de medio de pago"
    
    except Exception as e:
        return False, f"Error en página 3: {str(e)}"


# ==============================
# 📌 PÁGINA 4 - DATOS DEL PAGADOR
# ==============================
def pagina4_datos_pago(page):
    """
    Llena los datos del pagador y confirma.
    Retorna: (True, mensaje_exito) si exitoso, (False, mensaje_error) si falla
    """
    print("🌐 Página 4: Datos del pagador")
    
    try:
        # Esperar que cargue el formulario
        page.wait_for_selector("#contenido_Wizard2_ucTitularPago_ddlTipoDocumento", timeout=15_000)
        
        # 1. Seleccionar tipo de documento: Cédula de ciudadanía
        page.select_option("#contenido_Wizard2_ucTitularPago_ddlTipoDocumento", "2")
        print(f"   ✅ Tipo de documento: Cédula de ciudadanía")
        
        time.sleep(1)
        
        # 2. Número de identificación
        page.fill("#contenido_Wizard2_ucTitularPago_tbnumeroDocumento", "")
        page.fill("#contenido_Wizard2_ucTitularPago_tbnumeroDocumento", COD_LIQUIDACION)
        print(f"   ✅ Número de identificación: {COD_LIQUIDACION}")
        
        # 3. Nombre completo
        page.fill("#contenido_Wizard2_ucTitularPago_tbNombres", "")
        page.fill("#contenido_Wizard2_ucTitularPago_tbNombres", NOMBRE_COMPLETO)
        print(f"   ✅ Nombre completo: {NOMBRE_COMPLETO}")
        
        # 4. Teléfono
        page.fill("#contenido_Wizard2_ucTitularPago_tbTelefonoDepositante", "")
        page.fill("#contenido_Wizard2_ucTitularPago_tbTelefonoDepositante", TELEFONO)
        print(f"   ✅ Teléfono: {TELEFONO}")
        
        # 5. Correo
        page.fill("#contenido_Wizard2_ucTitularPago_ucCorreoElectronico_tbCorreoElectronico", "")
        page.fill("#contenido_Wizard2_ucTitularPago_ucCorreoElectronico_tbCorreoElectronico", CORREO)
        print(f"   ✅ Correo: {CORREO}")
        
        # 6. Confirmar correo
        page.fill("#contenido_Wizard2_ucTitularPago_ucCorreoElectronico_tbCorfirmCorreo", "")
        page.fill("#contenido_Wizard2_ucTitularPago_ucCorreoElectronico_tbCorfirmCorreo", CORREO)
        print(f"   ✅ Correo confirmado: {CORREO}")
        
        time.sleep(2)
        
        # 7. Dar clic en Continuar
        page.click("#contenido_Wizard2_FinishNavigationTemplateContainerID_FinishButton")
        print("   ✅ Clic en Continuar ejecutado")
        
        # Esperar 4 segundos
        print("   ⏳ Esperando 4 segundos...")
        time.sleep(4)
        
        # Verificar confirmación exitosa
        try:
            if page.is_visible("#contenido_Wizard2_ucInfoBanco_lbMensajeEnPopup"):
                mensaje = page.locator("#contenido_Wizard2_ucInfoBanco_lbMensajeEnPopup").text_content()
                print(f"   ✅ Confirmación exitosa")
                print(f"   📄 Mensaje: {mensaje[:100]}...")
                return True, "Proceso completado exitosamente"
            else:
                # Buscar errores
                error_selectors = [".error_validacion", ".alert-danger"]
                for selector in error_selectors:
                    if page.is_visible(selector):
                        error_text = page.locator(selector).text_content()
                        return False, f"Error en formulario: {error_text}"
                
                return True, "Proceso completado"
        except:
            return True, "Proceso completado"
        
    except TimeoutError:
        return False, "Timeout al cargar página de datos de pago"
    
    except Exception as e:
        return False, f"Error en página 4: {str(e)}"


# ==============================
# 📌 PROCESAR UN CÓDIGO COMPLETO
# ==============================
def procesar_codigo(page, codigo, correo):
    """
    Procesa un código completo desde página 1 hasta página 4.
    Retorna: (True, mensaje_exito) si exitoso, (False, mensaje_error) si falla
    """
    
    # Página 1: Consulta de trámite
    exito, mensaje = pagina1_consulta(page, codigo, correo)
    if not exito:
        return False, mensaje
    
    # Página 2: Confirmación de datos
    exito, mensaje = pagina2_confirmacion(page)
    if not exito:
        return False, mensaje
    
    # Página 3: Selección de medio de pago
    exito, mensaje = pagina3_medio_pago(page)
    if not exito:
        return False, mensaje
    
    # Página 4: Datos del pagador
    exito, mensaje = pagina4_datos_pago(page)
    if not exito:
        return False, mensaje
    
    return True, mensaje


# ==============================
# 📌 MAIN
# ==============================
def main():
    # Leer el archivo Excel
    try:
        df = leer_excel("entrada.xlsx")
    except Exception as e:
        print(f"❌ Error al leer el archivo Excel: {e}")
        sys.exit(1)
    
    # Asegurar que la columna OBSERVACIONES exista
    if "OBSERVACIONES" not in df.columns:
        df["OBSERVACIONES"] = ""
    
    with sync_playwright() as p:
        # Configuración del navegador
        browser = p.chromium.launch(
            headless=False,
            slow_mo=50,
            args=["--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage"]
        )
        
        # Crear contexto con tamaño visible
        context = browser.new_context(viewport={"width": 1280, "height": 800})
        page = context.new_page()
        
        try:
            # Recorrer todas las filas del Excel
            for i, row in df.iterrows():
                codigo = str(row["CODIGO"]).strip()
                
                # Validar que el código no esté vacío
                if not codigo or codigo == "" or codigo == "nan":
                    print(f"\n⏭️ Fila {i+1} tiene código vacío -> se omite")
                    df.at[i, "OBSERVACIONES"] = "Código vacío"
                    guardar_excel(df, "entrada.xlsx")
                    continue
                
                # Validar si ya tiene una observación exitosa (ya fue procesada)
                if pd.notna(row["OBSERVACIONES"]) and str(row["OBSERVACIONES"]).strip() != "":
                    obs = str(row["OBSERVACIONES"]).strip()
                    if "exitoso" in obs.lower() or "completado" in obs.lower():
                        print(f"\n⏭️ Fila {i+1} (código: {codigo}) ya fue procesada exitosamente -> se omite")
                        continue
                
                print(f"\n{'='*60}")
                print(f"📋 Procesando Fila {i+1}/{len(df)}")
                print(f"   Código: {codigo}")
                print(f"{'='*60}")
                
                try:
                    # Procesar el código completo
                    exito, mensaje = procesar_codigo(page, codigo, CORREO)
                    
                    if exito:
                        print(f"✅ Fila {i+1} procesada exitosamente")
                        df.at[i, "OBSERVACIONES"] = mensaje
                    else:
                        print(f"⚠️ Fila {i+1} con error: {mensaje}")
                        df.at[i, "OBSERVACIONES"] = mensaje
                
                except Exception as e:
                    print(f"❌ Error inesperado procesando fila {i+1}: {e}")
                    df.at[i, "OBSERVACIONES"] = f"Error inesperado: {str(e)}"
                
                # Guardar progreso después de cada fila
                guardar_excel(df, "entrada.xlsx")
                
                # Pequeña pausa entre filas
                time.sleep(2)
            
            print("\n" + "="*60)
            print("🎉 Proceso completado para todos los códigos")
            print("="*60)
        
        except Exception as e:
            print(f"\n❌ Error crítico durante la ejecución: {e}")
        
        finally:
            print("\n🔒 Cerrando navegador...")
            browser.close()


if __name__ == "__main__":
    main()