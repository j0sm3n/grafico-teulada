import os
import pyperclip
import openpyxl

NOMBRE_EXCEL = 'test.xlsx'
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
MESES = [
    'JULIO 2020',
    'AGOSTO 2020',
    'SEPTIEMBRE 2020',
    'OCTUBRE 2020',
    'NOVIEMBRE 2020',
    'DICIEMBRE 2020']


def lee_portapapeles():
    """
    Lee el contenido del portapapeles y formatea el texto, creando una 
    lista 'dias' con el siguiente contenido:
    dias[0] -> {'num_mes': '10-jul.'}
    dias[1] -> {'num_sem': 'vi.'}
    dias[2:] -> {'turnos': ['3', 'D', 'D', '7', '8'...]}
    """
    texto = pyperclip.paste()
    texto.replace('\n', '')
    texto_lista = texto.split()
    dias = []

    for i in range(0, len(texto_lista), NUM_AGENTES + 2):
        d = texto_lista[i:i + NUM_AGENTES + 2]
        dia = {
            'num_mes': d[0],
            'num_sem': d[1],
            'turnos': d[2:]
        }
        dias.append(dia)

    return dias


def crea_excel(agentes, dias, mes):
    if os.path.isfile(NOMBRE_EXCEL):
        wb = openpyxl.load_workbook(NOMBRE_EXCEL)
    else:
        wb = openpyxl.Workbook()
    wb.create_sheet(title=mes, index=MESES.index(mes))
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    sheet = wb[mes]
    # if sheet.title == 'Sheet':
    #     sheet.title = mes
    # else:
    #     wb.create_sheet(title=mes, index=MESES.index(mes))
    sheet['A1'] = mes

    for fila in range(3, 3 + len(AGENTES)):
        sheet.cell(row=fila, column=1).value = AGENTES[fila - 3].split(',')[0].lstrip()
        sheet.cell(row=fila, column=2).value = AGENTES[fila - 3].split(',')[1].lstrip()

    for columna in range(3, 3 + len(dias)):
        dia = dias[columna - 3]
        sheet.cell(row=1, column=columna).value = dia['num_mes']
        sheet.cell(row=2, column=columna).value = dia['num_sem']
        for fila in range(3, 3 + len(dia['turnos'])):
            sheet.cell(row=fila, column=columna).value = dia['turnos'][fila - 3]

    wb.save(NOMBRE_EXCEL)


def main():
    for mes in MESES:
        print(f"Copia el mes {mes} en el portapapeles y pulsa enter para continuar. Pulsa C para cancelar.")
        resp = input("> ")
        if resp == 'c' or resp == 'C':
            break
        else:
            dias = lee_portapapeles()
            crea_excel(AGENTES, dias, mes)


if __name__ == "__main__":
    main()
