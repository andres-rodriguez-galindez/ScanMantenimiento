# modules/scanner.py
import platform
import psutil
import socket
import getpass
import os
import subprocess

def mostrar_checklist_y_comandos():
    print("="*60)
    print("IMPORTANTE: Ejecute este programa como ADMINISTRADOR.")
    print("Ingrese como usuario administrador de la máquina.\n")
    print("Querido técnico, por favor tenga presente este check list:\n")
    print("1. Toma de fotografías y videos de como recibe el equipo.")
    print("2. Verificación de Funcionamiento.")
    print("3. Validar Serial del Equipo.")
    print("4. Validar Garantía de la máquina, fecha inicial y final, desde la página del fabricante.")
    print("5. Validación de los usuarios administrador y el usuario soporte, estén activos, actualizar las credenciales entregadas.")
    print("6. Levantar Inventario de Software.")
    print("7. Diagnóstico de Disco Duro.")
    print("8. Borrado de cache, temporales, eventos, papelera por usuario registrado en la máquina.")
    print("9. Actualización o instalación de Antivirus sea la última versión (si no está, instalarlo).")
    print("10. Limpieza interna de componentes.")
    print("11. Limpieza externa de dispositivos.")
    print("12. Revisión de Batería (Portátil).")
    print("13. Validación de Navegación en los equipos, que no navegue libremente a cualquier página web.")
    print("14. Validación al ingresar USB, el puerto debe estar bloqueado (si no, informar inmediatamente).")
    print("15. Actualización o instalación One Drive, Google Drive (Validar que todos los archivos estén sincronizados).")
    print("16. Actualizar parches del sistema operativo.")
    print("17. Fotografía o video de entrega de equipo funcional y operativo.\n")
    print("Puede usar estos comandos para agilizar su trabajo (ejecútelos desde CMD o Windows + R como administrador):\n")
    print("Visor de eventos: eventvwr")
    print("Papelera de reciclaje: shell:RecycleBinFolder")
    print("Visualizar datos del equipo: Dxdiag")
    print("Visualizar nombre del equipo: hostname")
    print("Visualizar el serial de la bios: wmic bios get serialnumber")
    print("Visualizar las cuentas de usuario: netplwiz")
    print("Visualizar los controladores de red: ncpa.cpl")
    print("Visualizar la información del equipo: systeminfo")
    print("Visualizar temporales del usuario: %temp%")
    print("Visualizar temporales del sistema: temp")
    print("Validar inicio de arranque: msconfig")
    print("Validar la versión del Windows: winver")
    print("Ingresar al control panel: control")
    print("Ingresar al administrador de tareas: taskmgr ó (CRTL + SHIFT + SCAPE)")
    print("Información del sistema: msinfo32")
    print("Propiedades del sistema: sysdm.cpl")
    print("Explorador de archivos: explorer")
    print("Programas y características: appwiz.cpl")
    print("Firewall de Windows: firewall.cpl")
    print("Informe de la bateria: powercfg\n")
    input("Presione cualquier tecla para continuar...")

def limpiar_temporales_y_cache():
    print("\nLimpiando archivos temporales, cache, visor de eventos y papelera de todos los usuarios...\n")
    # Limpiar %temp% y temp del usuario actual
    temp_paths = [os.environ.get('TEMP'), os.environ.get('TMP'), r'C:\Windows\Temp']
    for path in temp_paths:
        if path and os.path.exists(path):
            try:
                for archivo in os.listdir(path):
                    archivo_path = os.path.join(path, archivo)
                    try:
                        if os.path.isfile(archivo_path) or os.path.islink(archivo_path):
                            os.unlink(archivo_path)
                        elif os.path.isdir(archivo_path):
                            os.rmdir(archivo_path)
                    except Exception:
                        pass
            except Exception:
                pass
    # Limpiar papelera
    try:
        subprocess.run('PowerShell.exe -Command "Clear-RecycleBin -Force"', shell=True)
    except Exception:
        pass
    # Limpiar visor de eventos
    try:
        subprocess.run('wevtutil cl Application', shell=True)
        subprocess.run('wevtutil cl System', shell=True)
        subprocess.run('wevtutil cl Security', shell=True)
    except Exception:
        pass
    print("Limpieza completada.\n")

def obtener_valor_wmic_o_powershell(comando_wmic, comando_powershell, clave):
    try:
        valor = subprocess.check_output(comando_wmic, shell=True, text=True).split('\n')[1].strip()
        if valor and valor.lower() != clave.lower():
            return valor
    except Exception:
        pass
    # Si falla o está vacío, intenta con PowerShell
    try:
        valor = subprocess.check_output(
            ["powershell", "-Command", comando_powershell],
            text=True
        ).split('\n')[0].strip()
        return valor if valor else "No disponible"
    except Exception:
        return "No disponible"

def extraer_info_maquina():
    print("\nExtrayendo información de la máquina...\n")
    info = {}
    info["Nombre del equipo"] = socket.gethostname()
    info["Usuario actual"] = getpass.getuser()
    info["Sistema operativo"] = platform.platform()
    info["Arquitectura"] = platform.architecture()[0]
    info["Procesador"] = platform.processor()
    info["Memoria RAM (GB)"] = round(psutil.virtual_memory().total / (1024**3), 2)
    # Serial, marca y modelo
    info["Serial"] = obtener_valor_wmic_o_powershell(
        'wmic bios get serialnumber',
        "(Get-WmiObject win32_bios).SerialNumber",
        "SerialNumber"
    )
    info["Marca"] = obtener_valor_wmic_o_powershell(
        'wmic computersystem get manufacturer',
        "(Get-WmiObject win32_computersystem).Manufacturer",
        "Manufacturer"
    )
    info["Modelo"] = obtener_valor_wmic_o_powershell(
        'wmic computersystem get model',
        "(Get-WmiObject win32_computersystem).Model",
        "Model"
    )
    # Discos
    discos = []
    for disk in psutil.disk_partitions():
        try:
            usage = psutil.disk_usage(disk.mountpoint)
            discos.append({
                "Disco": disk.device,
                "Tamaño total (GB)": round(usage.total / (1024**3), 2),
                "Disponible (GB)": round(usage.free / (1024**3), 2)
            })
        except Exception:
            pass
    info["Discos"] = discos
    # Tarjeta gráfica
    try:
        salida = subprocess.check_output('wmic path win32_VideoController get name', shell=True, text=True).split('\n')
        tarjetas = [line.strip() for line in salida[1:] if line.strip()]
        info["Tarjeta gráfica"] = tarjetas
    except Exception:
        info["Tarjeta gráfica"] = "No disponible"
    # Aplicativos instalados (Panel de control)
    aplicativos = []
    try:
        salida = subprocess.check_output('wmic product get name,installdate,version,vendor', shell=True, text=True, errors='ignore')
        lineas = salida.split('\n')[1:]
        for linea in lineas:
            if linea.strip():
                aplicativos.append(linea.strip())
    except Exception:
        aplicativos.append("No disponible")
    info["Aplicativos instalados"] = aplicativos

    # Mostrar la información en pantalla
    for k, v in info.items():
        if k == "Discos" and isinstance(v, list):
            print(f"{k}:")
            for disco in v:
                print(f"  - Unidad: {disco['Disco']}")
                print(f"    Tamaño total (GB): {disco['Tamaño total (GB)']}")
                print(f"    Disponible (GB): {disco['Disponible (GB)']}")
        elif isinstance(v, list):
            print(f"{k}:")
            for item in v:
                print(f"  {item}")
        else:
            print(f"{k}: {v}")
    print("\nPara consultar la garantía, use el serial en la web del fabricante.\n")

    # Guardar la información en HTML
    html = "<html><head><meta charset='utf-8'><title>Reporte de Información de la Máquina</title></head><body>"
    html += "<h1>Reporte de Información de la Máquina</h1><ul>"
    for k, v in info.items():
        if k == "Discos" and isinstance(v, list):
            html += f"<li><b>{k}:</b><ul>"
            for disco in v:
                html += "<li>"
                html += f"Unidad: {disco['Disco']}<br>"
                html += f"Tamaño total (GB): {disco['Tamaño total (GB)']}<br>"
                html += f"Disponible (GB): {disco['Disponible (GB)']}"
                html += "</li>"
            html += "</ul></li>"
        elif isinstance(v, list):
            html += f"<li><b>{k}:</b><ul>"
            for item in v:
                html += f"<li>{item}</li>"
            html += "</ul></li>"
        else:
            html += f"<li><b>{k}:</b> {v}</li>"
    html += "</ul>"
    html += "<p><i>Para consultar la garantía, use el serial en la web del fabricante.</i></p>"
    html += "</body></html>"
    with open("reporte_maquina.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("También se ha guardado un archivo 'reporte_maquina.html' en la carpeta actual. Puede abrirlo con su navegador y exportarlo a PDF.\n")

def mostrar_menu():
    print("="*60)
    print("Seleccione una opción:")
    print("1. Limpiar temporales, cache, visor de eventos y papelera de todos los usuarios")
    print("2. Extraer información de la máquina")
    print("3. Salir")

if __name__ == "__main__":
    mostrar_checklist_y_comandos()
    while True:
        mostrar_menu()
        opcion = input("Ingrese el número de la opción: ")
        if opcion == "1":
            limpiar_temporales_y_cache()
        elif opcion == "2":
            extraer_info_maquina()
        elif opcion == "3":
            print("Saliendo...")
            break
        else:
            print("Opción no válida. Intente de nuevo.")
        print("-" * 40)
