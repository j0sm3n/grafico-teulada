import pyperclip

NUM_AGENTES = 19
AGENTES = [
    'OCHOA BAEZA, ANTONIO',
    'SANJUAN IZQUIERDO, YOLANDA',
    'DE LA TORRE HERRERA, RICARDO',
    'BIELSA NOGUES, ERNESTO',
    'COLLADO MASCAROS, IVAN',
    'GOMEZ SEGURA, JOSE',
    'LARA NAVARRO, MANUEL',
    'PASTOR BERTOMEU, JAIME JUAN',
    'MENDOZA RODRIGUEZ, JOSE ANTONIO',
    'MENA DIAZ, VICTOR',
    'ESTEVE GIMENO, ENRIQUETA',
    'PEREZ SANTAMARIA, MONICA',
    'CARMONA GARCIA, DIEGO JOSE',
    'BOU SERRALTA, MIGUEL ANGEL',
    'LLOPES MOLINA, JUAN MANUEL',
    'CABRERA GUERRERO, ENRIQUE',
    'DE LAMA GONZALEZ, CARLOS',
    'GOMEZ MARTIN, JAIME',
    'MORALES ASENSI, ENRIQUE']

texto = pyperclip.paste()
texto.replace('\n', '')
texto_lista = texto.split()
dias_lista = []
dias = []

for i in range(0, len(texto_lista), NUM_AGENTES + 2):
    dias_lista.append(texto_lista[i:i + NUM_AGENTES + 2])

for d in dias_lista:
    dia = {
        'num_mes': d[0],
        'num_sem': d[1],
        'turnos': d[2:]
    }
    dias.append(dia)
    print()

# for dia in dias:
#     for i, turno in enumerate(dia):
#         if i < 2:
#             print(turno, end=' ')
#         else:
#             print(f"{AGENTES[i - 2]}, turno {turno}")

# print(type(dias))
# print(dias)