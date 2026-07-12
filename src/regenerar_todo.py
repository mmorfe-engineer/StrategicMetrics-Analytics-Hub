#!/usr/bin/env python3
"""
Script Maestro para Regenerar Todos los Módulos de La Tiendita UVM

Este script ejecuta secuencialmente los scripts de generacion para los modulos 1-8,
permitiendo regenerar el proyecto completo desde cero.

Uso:
    python regenerar_todo.py

Opciones:
    --modulo N    : Regenerar solo el modulo N (1-8)
    --verificar   : Verificar archivos generados sin regenerar
    --limpio      : Eliminar archivos existentes antes de regenerar
"""

import subprocess
import sys
import os
import argparse

# Ruta base del proyecto
BASE_PATH = "/home/bitcoinpapa/Documentos/uvm estadistica modulo 8"
SCRIPTS_PATH = os.path.join(BASE_PATH, "StrategicMetrics-Analytics-Hub", "src", "La_Tiendita_UVM")
DASHBOARDS_PATH = os.path.join(BASE_PATH, "StrategicMetrics-Analytics-Hub", "dashboards", "La_Tiendita_UVM")

# Mapeo de modulos: numero -> nombre de script y archivo de salida
MODULOS = {
    1: {"script": "generar_modulo1.py", "output": "Modulo1_VentasGlobales.xlsx"},
    2: {"script": "generar_modulo2.py", "output": "Modulo2_AnalisisSucursales.xlsx"},
    3: {"script": "generar_modulo3.py", "output": "Modulo3_PortafolioProductos.xlsx"},
    4: {"script": "generar_modulo4.py", "output": "Modulo4_ComportamientoTemporal.xlsx"},
    5: {"script": "generar_modulo5.py", "output": "Modulo5_InventarioRotacion.xlsx"},
    6: {"script": "generar_modulo6.py", "output": "Modulo6_PreciosMargenes.xlsx"},
    7: {"script": "generar_modulo7.py", "output": "Modulo7_SatisfaccionClientes.xlsx"},
    8: {"script": "generar_modulo8.py", "output": "Modulo8_Dashboard_TienditaUVM.xlsx"},
}


def verificar_archivo(modulo_num):
    """Verificar si el archivo del modulo existe"""
    output_file = MODULOS[modulo_num]["output"]
    file_path = os.path.join(DASHBOARDS_PATH, output_file)
    return os.path.exists(file_path)


def ejecutar_modulo(modulo_num):
    """Ejecutar el script de un modulo"""
    script_name = MODULOS[modulo_num]["script"]
    script_path = os.path.join(SCRIPTS_PATH, script_name)
    
    print(f"\n{'='*70}")
    print(f"🔄 Ejecutando Módulo {modulo_num}: {script_name}")
    print(f"{'='*70}")
    
    try:
        # Cambiar al directorio base
        os.chdir(BASE_PATH)
        
        # Ejecutar el script con Python
        result = subprocess.run(
            ["python3", script_path],
            capture_output=True,
            text=True,
            timeout=300
        )
        
        # Mostrar salida
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(f"⚠️  Advertencias: {result.stderr}", file=sys.stderr)
        
        # Verificar resultado
        if result.returncode == 0:
            if verificar_archivo(modulo_num):
                print(f"✅ Módulo {modulo_num} completado: {MODULOS[modulo_num]['output']} generado")
                return True
            else:
                print(f"❌ Error: Archivo {MODULOS[modulo_num]['output']} no generado")
                return False
        else:
            print(f"❌ Error en ejecución del Módulo {modulo_num}")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"❌ Timeout: Módulo {modulo_num} tardó demasiado")
        return False
    except Exception as e:
        print(f"❌ Excepción: {str(e)}")
        return False


def limpiar_archivos():
    """Eliminar archivos de salida existentes"""
    print("\n🧹 Limpianado archivos existentes...")
    for modulo_num in range(1, 9):
        output_file = MODULOS[modulo_num]["output"]
        file_path = os.path.join(DASHBOARDS_PATH, output_file)
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"   🗑️  Eliminado: {output_file}")


def verificar_todo():
    """Verificar todos los archivos generados"""
    print("\n🔍 Verificando archivos generados...")
    todos_ok = True
    for modulo_num in range(1, 9):
        if verificar_archivo(modulo_num):
            print(f"   ✅ Módulo {modulo_num}: {MODULOS[modulo_num]['output']}")
        else:
            print(f"   ❌ Módulo {modulo_num}: {MODULOS[modulo_num]['output']} FALTANTE")
            todos_ok = False
    return todos_ok


def main():
    parser = argparse.ArgumentParser(
        description="Regenerar módulos de La Tiendita UVM",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python regenerar_todo.py              # Regenerar todo
  python regenerar_todo.py --modulo 8   # Solo Módulo 8
  python regenerar_todo.py --limpio   # Limpiar antes de regenerar
  python regenerar_todo.py --verificar # Solo verificar
        """
    )
    parser.add_argument("--modulo", type=int, choices=range(1, 9), 
                       help="Regenerar solo un módulo específico (1-8)")
    parser.add_argument("--verificar", action="store_true", 
                       help="Verificar archivos sin regenerar")
    parser.add_argument("--limpio", action="store_true", 
                       help="Eliminar archivos existentes antes de regenerar")
    
    args = parser.parse_args()
    
    # Mostrar encabezado
    print("""
╔══════════════════════════════════════════════════════════════════════╗
║     SCRIPT MAESTRO - La Tiendita UVM - Regenerar Todos los Módulos     ║
║                  Universidad Valle del Momboy - 2026B                   ║
║                   Responsable: Martin Morfe                            ║
╚══════════════════════════════════════════════════════════════════════╝
    """)
    
    # Verificar solo
    if args.verificar:
        if verificar_todo():
            print("\n✅ Todos los módulos están generados correctamente")
            sys.exit(0)
        else:
            print("\n❌ Algunos módulos faltan")
            sys.exit(1)
    
    # Limpiar archivos existentes
    if args.limpio:
        limpiar_archivos()
    
    # Ejecutar modulo especifico
    if args.modulo:
        if ejecutar_modulo(args.modulo):
            print(f"\n✅ Módulo {args.modulo} regenerado exitosamente")
            sys.exit(0)
        else:
            print(f"\n❌ Falló regeneración del Módulo {args.modulo}")
            sys.exit(1)
    
    # Ejecutar todos los modulos
    print("\n🚀 Regenerando TODOS los módulos (1-8)...")
    print("   Esto puede tardar varios minutos...")
    
    exitosos = 0
    fallidos = 0
    
    for modulo_num in range(1, 9):
        if ejecutar_modulo(modulo_num):
            exitosos += 1
        else:
            fallidos += 1
            # Preguntar si continuar
            if modulo_num < 8:
                respuesta = input(f"\n⚠️  Módulo {modulo_num} falló. ¿Continuar con el siguiente? (s/n): ").lower()
                if respuesta != 's':
                    print("\n❌ Proceso detenido por el usuario")
                    sys.exit(1)
    
    # Resumen final
    print(f"\n{'='*70}")
    print(f"📊 RESUMEN DE REGENERACIÓN")
    print(f"{'='*70}")
    print(f"   ✅ Éxitos: {exitosos}/8 módulos")
    print(f"   ❌ Fallos: {fallidos}/8 módulos")
    
    if fallidos == 0:
        print(f"\n✅ Todos los módulos regenerados exitosamente")
        print(f"\n📁 Archivos generados en: {DASHBOARDS_PATH}")
        sys.exit(0)
    else:
        print(f"\n⚠️  {fallidos} módulo(s) fallaron. Revisa los mensajes de error.")
        sys.exit(1)


if __name__ == "__main__":
    main()
