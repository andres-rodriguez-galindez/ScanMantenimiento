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

def obtener_lista_usuarios_netuser():
    try:
        salida = subprocess.check_output('net user', shell=True, text=True, errors='ignore')
        lineas = salida.splitlines()
        usuarios = []
        captura = False
        for linea in lineas:
            if "----" in linea:
                captura = not captura
                continue
            if captura:
                if linea.strip() and "El comando se ha ejecutado correctamente." not in linea:
                    usuarios += [u for u in linea.split() if u.strip()]
        # Filtrar posibles palabras que no sean usuarios válidos
        palabras_invalidas = {"El", "comando", "se", "ha", "ejecutado", "correctamente.", "The", "command", "completed", "successfully."}
        usuarios = [u for u in usuarios if u not in palabras_invalidas]
        return usuarios
    except Exception:
        return ["No disponible"]

def obtener_usuarios_estado():
    usuarios = []
    try:
        salida = subprocess.check_output('net user', shell=True, text=True, errors='ignore')
        lineas = salida.splitlines()
        usuarios_encontrados = []
        captura = False
        for linea in lineas:
            if "----" in linea:
                captura = not captura
                continue
            if captura:
                # Solo líneas con texto y sin mensaje final
                if linea.strip() and "El comando se ha ejecutado correctamente." not in linea:
                    usuarios_encontrados += [u for u in linea.split() if u.strip()]
        # Filtrar posibles palabras que no sean usuarios válidos
        palabras_invalidas = {"El", "comando", "se", "ha", "ejecutado", "correctamente.", "The", "command", "completed", "successfully."}
        usuarios_encontrados = [u for u in usuarios_encontrados if u not in palabras_invalidas]
        # Depuración: imprime los usuarios encontrados
        print("Usuarios detectados:", usuarios_encontrados)
        # Consultar el estado de cada usuario real
        for usuario in usuarios_encontrados:
            try:
                detalle = subprocess.check_output(f'net user "{usuario}"', shell=True, text=True, errors='ignore')
                # Busca tanto en español como en inglés
                if "Cuenta activa" in detalle:
                    if "Cuenta activa" in detalle and "Sí" in detalle:
                        estado = "Activo"
                    elif "Cuenta activa" in detalle and "No" in detalle:
                        estado = "Inactivo"
                    else:
                        continue
                elif "Account active" in detalle:
                    if "Account active" in detalle and "Yes" in detalle:
                        estado = "Activo"
                    elif "Account active" in detalle and "No" in detalle:
                        estado = "Inactivo"
                    else:
                        continue
                else:
                    continue  # Si no se puede determinar, no lo muestra
                usuarios.append(f"{usuario} ({estado})")
            except Exception:
                continue  # Si hay error, no lo muestra
    except Exception:
        usuarios.append("No disponible")
    return usuarios

def obtener_aplicativos_instalados():
    try:
        comando = [
            "powershell",
            "-Command",
            r"""
            Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
                              HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
            Where-Object { $_.DisplayName } |
            Select-Object DisplayName, Publisher, InstallDate, DisplayVersion, EstimatedSize |
            Sort-Object DisplayName
            """
        ]
        salida = subprocess.check_output(comando, text=True, errors='ignore')
        lineas = salida.splitlines()
        # Buscar el encabezado y los datos
        datos = []
        encabezado = []
        for i, linea in enumerate(lineas):
            if "DisplayName" in linea and "Publisher" in linea:
                encabezado = [col.strip() for col in linea.split()]
                start = i + 2  # Los datos empiezan dos líneas después
                break
        for linea in lineas[start:]:
            if linea.strip():
                # Separar por columnas fijas (puede requerir ajuste según el sistema)
                partes = [linea[0:40].strip(), linea[40:70].strip(), linea[70:80].strip(), linea[80:100].strip(), linea[100:].strip()]
                datos.append(partes)
        return encabezado, datos
    except Exception:
        return ["Nombre", "Editor", "Fecha de instalación", "Versión", "Tamaño (KB)"], [["No disponible"]*5]

# Obtén la salida de query user
def obtener_query_user():
    try:
        salida = subprocess.check_output('query user', shell=True, text=True, errors='ignore')
        lineas = salida.strip().split('\n')
        return lineas
    except Exception:
        return ["No disponible"]

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
    # Usuarios y estado
    info["Usuarios"] = obtener_usuarios_estado()
    info["UsuariosNetUser"] = obtener_lista_usuarios_netuser()
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
    # Aplicativos instalados (más parecido a Programas y características)
    encabezado, aplicativos = obtener_aplicativos_instalados()
    info["Aplicativos instalados"] = aplicativos
    info["Encabezado aplicativos"] = encabezado
    # Sesiones activas
    info["QueryUser"] = obtener_query_user()

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
    html = f"""
<html>
<head>
    <meta charset='utf-8'>
    <title>Reporte de Información de la Máquina</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 30px; }}
        h1 {{ color: #2c3e50; }}
        h2 {{ color: #34495e; }}
        ul {{ margin-bottom: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 30px; }}
        th, td {{ border: 1px solid #888; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        .section {{ margin-bottom: 30px; }}
    </style>
</head>
<body>
    <h1>Reporte de Información de la Máquina</h1>

    <div class="section">
        <h2>Datos Generales</h2>
        <ul>
            <li><b>Nombre del equipo:</b> {info["Nombre del equipo"]}</li>
            <li><b>Usuario actual:</b> {info["Usuario actual"]}</li>
            <li><b>Sistema operativo:</b> {info["Sistema operativo"]}</li>
            <li><b>Arquitectura:</b> {info["Arquitectura"]}</li>
            <li><b>Procesador:</b> {info["Procesador"]}</li>
            <li><b>Memoria RAM (GB):</b> {info["Memoria RAM (GB)"]}</li>
            <li><b>Serial:</b> {info["Serial"]}</li>
            <li><b>Marca:</b> {info["Marca"]}</li>
            <li><b>Modelo:</b> {info["Modelo"]}</li>
        </ul>
    </div>

    <div class="section">
        <h2>Discos</h2>
        <ul>
"""
    for disco in info["Discos"]:
        html += f"<li><b>Unidad:</b> {disco['Disco']} | <b>Tamaño total (GB):</b> {disco['Tamaño total (GB)']} | <b>Disponible (GB):</b> {disco['Disponible (GB)']}</li>"

    html += """
        </ul>
    </div>

    <div class="section">
        <h2>Tarjeta Gráfica</h2>
        <ul>
"""
    if isinstance(info["Tarjeta gráfica"], list):
        for tarjeta in info["Tarjeta gráfica"]:
            html += f"<li>{tarjeta}</li>"
    else:
        html += f"<li>{info['Tarjeta gráfica']}</li>"

    html += """
        </ul>
    </div>

    <div class="section">
        <h2>Usuarios del Sistema</h2>
        <ul>
"""
    for usuario in info["Usuarios"]:
        html += f"<li>{usuario}</li>"

    html += """
        </ul>
    </div>
    <div class="section">
        <h2>Usuarios del Sistema (net user)</h2>
        <table>
            <tr>
                <th>Nombre de usuario</th>
            </tr>
"""
    for usuario in info["UsuariosNetUser"]:
        html += f"<tr><td>{usuario}</td></tr>"

    html += """
        </table>
    </div>
    <div class="section">
        <h2>Sesiones activas (query user)</h2>
        <pre>
"""
    for linea in info["QueryUser"]:
        html += f"{linea}\n"
    html += """
        </pre>
    </div>

    <div class="section">
        <h2>Aplicativos Instalados</h2>
        <table>
            <tr>
"""
    for col in info["Encabezado aplicativos"]:
        html += f"<th>{col}</th>"
    html += "</tr>\n"

    for app in info["Aplicativos instalados"]:
        html += "<tr>"
        for campo in app:
            html += f"<td>{campo}</td>"
        html += "</tr>\n"

    html += """
        </table>
    </div>
    <p><i>Para consultar la garantía, use el serial en la web del fabricante.</i></p>
</body>
</html>
"""

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
