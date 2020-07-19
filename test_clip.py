import os
import sys
from datetime import datetime
import pyperclip
import locale
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, NamedStyle, Color, PatternFill


NOMBRE_EXCEL = 'test2.xlsx'
AGENTES = [
    '677, OCHOA BAEZA, ANTONIO',
    '2560, SANJUAN IZQUIERDO, YOLANDA',
    '2082, DE LA TORRE HERRERA, RICARDO',
    '109, BIELSA NOGUES, ERNESTO',
    '2049, COLLADO MASCAROS, IVAN',
    '409, GOMEZ SEGURA, JOSE',
    '2154, LARA NAVARRO, MANUEL',
    '1651, PASTOR BERTOMEU, JAIME JUAN',
    '2076, MENDOZA RODRIGUEZ, JOSE ANTONIO',
    '597, MENA DIAZ, VICTOR',
    '2316, ESTEVE GIMENO, ENRIQUETA',
    '2586, PEREZ SANTAMARIA, MONICA',
    '159, CARMONA GARCIA, DIEGO JOSE',
    '133, BOU SERRALTA, MIGUEL ANGEL',
    '2588, LLOPES MOLINA, JUAN MANUEL',
    '2369, CABRERA GUERRERO, ENRIQUE',
    '1204, DE LAMA GONZALEZ, CARLOS',
    '1472, GOMEZ MARTIN, JAIME',
    '2210, MORALES ASENSI, ENRIQUE']
NUM_AGENTES = len(AGENTES)
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

    # Para que pueda detectar el nombre de los meses en español
    locale.setlocale(locale.LC_ALL, 'es_ES')
    
    texto = pyperclip.paste()
    texto.replace('\n', '')
    texto_lista = texto.split()
    dias = []

    for i in range(0, len(texto_lista), NUM_AGENTES + 2):
        d = texto_lista[i:i + NUM_AGENTES + 2]
        fecha = datetime.strptime(d[0].replace('.', '-2020'), '%d-%b-%Y')
        dia = {
            'num_mes': fecha,
            'num_sem': d[1],
            'turnos': d[2:]
        }
        dias.append(dia)

    return dias


def crea_excel(agentes, dias, mes):
    """
    Crea un fichero excel con los datos obtenidos anteriormente.
    """

    # Si ya existe el fichero lo abre.
    if os.path.isfile(NOMBRE_EXCEL):
        wb = openpyxl.load_workbook(NOMBRE_EXCEL)
    # En caso de que no exista lo crea.
    else:
        wb = openpyxl.Workbook()
    # Crea una nueva hoja con el nombre del mes y año.
    wb.create_sheet(title=mes, index=MESES.index(mes))
    # Borra la hoja inicial 'Sheet'
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    sheet = wb[mes]
    sheet['A1'] = mes

    # Introduce apellidos y nombre de los agentes en las columnas A y B
    for fila in range(3, 3 + NUM_AGENTES):
        sheet.cell(row=fila, column=1).value = AGENTES[fila - 3].split(',')[1].lstrip()
        sheet.cell(row=fila, column=2).value = AGENTES[fila - 3].split(',')[2].lstrip()

    # Introduce los turnos del mes en columnas por día, utilizando la
    # primera y segunda fila para el día del mes, y el resto de filas para
    # los turnos de cada agente.
    for columna in range(3, 3 + len(dias)):
        dia = dias[columna - 3]
        sheet.cell(row=1, column=columna).value = dia['num_mes']
        sheet.cell(row=2, column=columna).value = dia['num_mes']
        for fila in range(3, 3 + len(dia['turnos'])):
            if dia['turnos'][fila - 3].isdigit():
                sheet.cell(row=fila, column=columna).value = int(dia['turnos'][fila - 3])
            else:
                sheet.cell(row=fila, column=columna).value = dia['turnos'][fila - 3]

    # Graba los datos en el fichero excel.
    wb.save(NOMBRE_EXCEL)


def formatea_excel():
    """
    Da formato al fichero excel
    """
    centrado = Alignment(horizontal='center', vertical='center')
    negrita = Font(bold=True)
    dia_mes = NamedStyle(name='dia_mes', number_format='dd mmm')
    dia_mes.font = Font(bold=True)
    dia_mes.alignment = Alignment(horizontal='center', vertical='center')
    dia_sem = NamedStyle(name='dia_sem', number_format='ddd')
    dia_sem.font = Font(bold=True)
    dia_sem.alignment = Alignment(horizontal='center', vertical='center')
    fondo_gris = NamedStyle(
        name='gris',
        fill=PatternFill(
            patternType='solid',
            fill_type='solid',
            fgColor=Color('CBCBCB')))
    letra_roja = Font(color='FF2600')
    
    # Comprueba si existe el fichero
    if not os.path.isfile(NOMBRE_EXCEL):
        print(f"No se encuentra el fichero excel {NOMBRE_EXCEL}.")
        sys.exit(1)
    else:
        wb = openpyxl.load_workbook(NOMBRE_EXCEL)
    
    for sheet in wb.sheetnames:
        # Ancho de la columna para los apellidos
        wb[sheet].column_dimensions['A'].width = 18
        # Ancho de la columna para el nombre
        wb[sheet].column_dimensions['B'].width = 12
        # Fusiona las celdas para el nombre del mes
        wb[sheet].merge_cells('A1:B2')
        wb[sheet]['A1'].alignment = centrado
        wb[sheet]['A1'].font = negrita
        # Pinta lineas grises
        for columna in range(1, wb[sheet].max_column + 1):
            letra_col = get_column_letter(columna)
            for fila in range(3, wb[sheet].max_row + 1):
                if fila % 2 != 0:
                    wb[sheet][f'{letra_col}{fila}'].style = fondo_gris
                
        for columna in range(3, wb[sheet].max_column + 1):
            letra_col = get_column_letter(columna)
            # Ancho de la columna para los turnos
            wb[sheet].column_dimensions[letra_col].width = 7
            for fila in range(1, wb[sheet].max_row + 1):
                # Fila día del mes
                if fila == 1:
                    wb[sheet][f'{letra_col}{fila}'].style = dia_mes
                    # Si es sábado o domingo utiliza letra roja
                    if wb[sheet][f'{letra_col}{fila}'].value.weekday() == 5 or wb[sheet][f'{letra_col}{fila}'].value.weekday() == 6:
                        wb[sheet][f'{letra_col}{fila}'].font = letra_roja
                # Fila día de la semana
                elif fila == 2:
                    wb[sheet][f'{letra_col}{fila}'].style = dia_sem
                    # Si es sábado o domingo utiliza letra roja
                    if wb[sheet][f'{letra_col}{fila}'].value.weekday() == 5 or wb[sheet][f'{letra_col}{fila}'].value.weekday() == 6:
                        wb[sheet][f'{letra_col}{fila}'].font = letra_roja
                # Filas de turnos
                else:
                    wb[sheet][f'{letra_col}{fila}'].alignment = centrado

    wb.save(NOMBRE_EXCEL)


def crea_bd(agentes, dias):
    """
    Crea una base de datos con los turnos de cada agente
    """
    pass


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
    formatea_excel()
