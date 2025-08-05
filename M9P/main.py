import pyttsx3
import time
import os
import comtypes.client

def mostrar_titulo():
    os.system("cls" if os.name == "nt" else "clear")
    print("=" * 40)
    print("🗣️  APLICACIÓN DE SÍNTESIS DE VOZ")
    print("=" * 40)

def listar_voces_sapi5():
    sapi = comtypes.client.CreateObject("SAPI.SpVoice")
    token_enum = sapi.GetVoices()
    voces = []
    for i in range(token_enum.Count):
        token = token_enum.Item(i)
        voces.append((token.GetDescription(), token.Id))
    return voces

def crear_motor(voice_id=None):
    engine = pyttsx3.init()
    engine.setProperty('rate', 150)
    engine.setProperty('volume', 0.8)
    if voice_id:
        engine.setProperty('voice', voice_id)
    return engine

def seleccionar_voz():
    voces = listar_voces_sapi5()
    print("\nVoces disponibles:\n")
    for i, (nombre, _) in enumerate(voces):
        print(f"[{i}] {nombre}")
    while True:
        try:
            seleccion = int(input("\nEscribe el número de la voz que quieres usar: "))
            if 0 <= seleccion < len(voces):
                print(f"\n✅ Voz seleccionada: {voces[seleccion][0]}\n")
                return voces[seleccion][1]
            else:
                print("Número fuera de rango, intenta otra vez.")
        except ValueError:
            print("Entrada inválida, escribe un número.")

def main():
    voice_id = seleccionar_voz()
    engine = crear_motor(voice_id)

    while True:
        mostrar_titulo()
        texto = input("📝 Escribe el texto que quieres que la computadora diga (o 'salir' para terminar): ").strip()
        if texto.lower() == 'salir':
            print("👋 Cerrando la aplicación...")
            time.sleep(1)
            break
        elif texto == "":
            print("⚠️ No escribiste nada. Intenta de nuevo.")
            time.sleep(1.5)
            continue

        print("🔊 Hablando...")
        engine.say(texto)
        engine.runAndWait()

        # Opciones después de hablar
        while True:
            print("\n¿Quieres cambiar la voz o continuar con la misma?")
            print("1. Cambiar voz")
            print("2. Continuar con la misma voz")
            opcion = input("Elige una opción (1 o 2): ").strip()

            if opcion == '1':
                voice_id = seleccionar_voz()
                engine.stop()          # Detener cualquier síntesis pendiente
                del engine             # Eliminar motor viejo
                engine = crear_motor(voice_id)  # Crear nuevo motor con nueva voz
                break
            elif opcion == '2':
                break
            else:
                print("⚠️ Opción inválida, intenta otra vez.")

if __name__ == "__main__":
    main()
