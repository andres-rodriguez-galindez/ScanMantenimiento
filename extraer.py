# modules/scanner.py
import platform
import psutil
import socket
import getpass
import os
import subprocess

def get_basic_info():
    info = {
        "hostname": socket.gethostname(),
        "usuario_actual": getpass.getuser(),
        "sistema_operativo": platform.platform(),
        "arquitectura": platform.architecture(),
        "procesador": platform.processor(),
        "memoria_total": round(psutil.virtual_memory().total / (1024**3), 2),  # GB
        "discos": [],
        "usuarios": []
    }

    # Listar discos
    for disk in psutil.disk_partitions():
        usage = psutil.disk_usage(disk.mountpoint)
        info["discos"].append({
            "device": disk.device,
            "mountpoint": disk.mountpoint,
            "total_GB": round(usage.total / (1024**3), 2)
        })

    # Listar usuarios registrados
    try:
        users = subprocess.check_output('net user', shell=True, text=True)
        info["usuarios"] = users.split('\n')[4:]
    except Exception as e:
        info["usuarios"] = [str(e)]

    return info

def mostrar_menu():
    print("Seleccione una opción:")
    print("1. Mostrar información básica del sistema")
    print("2. Mostrar discos")
    print("3. Mostrar usuarios registrados")
    print("4. Salir")

if __name__ == "__main__":
    while True:
        mostrar_menu()
        opcion = input("Ingrese el número de la opción: ")
        data = get_basic_info()
        if opcion == "1":
            for k, v in data.items():
                if k not in ["discos", "usuarios"]:
                    print(f"{k}: {v}")
        elif opcion == "2":
            print("Discos:")
            for disco in data["discos"]:
                print(disco)
        elif opcion == "3":
            print("Usuarios registrados:")
            for usuario in data["usuarios"]:
                print(usuario.strip())
        elif opcion == "4":
            print("Saliendo...")
            break
        else:
            print("Opción no válida. Intente de nuevo.")
        print("-" * 40)
