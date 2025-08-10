import pandas as pd
import datetime
import numpy as np
import warnings
from corregir_nombres import corregir_nombre
from corregir_grados import corregir_grado
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import locale
import seaborn as sns


locale.setlocale(locale.LC_TIME, "es_ES.utf8")  # ajustar al español

hoy_dt = datetime.datetime.now()
dia = hoy_dt.strftime("%d")
mes = hoy_dt.strftime("%B").capitalize()
anio = hoy_dt.strftime("%Y")
hoy = f"{dia} de {mes} de {anio}"

# Configuración del servidor SMTP de Outlook
SMTP_SERVER = "smtp.office365.com"  
SMTP_PORT = 587
EMAIL_USER = "ejemplo@gmail.com" # aqui va el correo institucional desde el cual se envian los informes para todos los estudiantes del colegio
EMAIL_PASSWORD = "contraseña"  # contraseña del correo

a = 'NO'
while a != 'SI':
    # Ingresar periodo actual
    periodo_actual = int(input('Ingrese el periodo actual: '))
    # Ingresar la semana actual
    semana_actual = int(input('Ingrese la semana actual: '))
    print(f'Es el periodo {periodo_actual} y semana {semana_actual}, es correcto ? escribe SI o NO')
    a = input ('')

# Leer los datos necesarios, lista de estudiantes y base de datos
notas = pd.read_excel('C:/Users/Admin/Desktop/GK/BASE DE DATOS 2025.xlsx', sheet_name='GK2025')
estudiantes = pd.read_excel('C:/Users/Admin/Desktop/GK/PS.xlsx', sheet_name = 'g')
planeacion_primaria = pd.read_excel('C:/Users/Admin/Desktop/GK/PS.xlsx', sheet_name = 'pp')
planeacion_bachillerato = pd.read_excel('C:/Users/Admin/Desktop/GK/PS.xlsx', sheet_name = 'pb')


###################################################  ARREGLAR HORARIO PARA EL PDF ########################################

#PRIMARIA

# ACA SE MUEVEN LAS COLUMNAS DEL HORARIO PARA QUE QUEDEN COMO EN LA APGINA QUE SE PUBLICABA EN TEAMS
# ES DECIR QUE QUEDEN LOS DOS MODULOS DE CADA DIA JUNTOS 

columna_a_mover = 'L.1'
planeacion_primaria.insert(3, columna_a_mover, planeacion_primaria.pop(columna_a_mover))
columna_a_mover = 'M.1'
planeacion_primaria.insert(5, columna_a_mover, planeacion_primaria.pop(columna_a_mover))
columna_a_mover = 'W.1'
planeacion_primaria.insert(7, columna_a_mover, planeacion_primaria.pop(columna_a_mover))
columna_a_mover = 'J.1'
planeacion_primaria.insert(9, columna_a_mover, planeacion_primaria.pop(columna_a_mover))
columna_a_mover = 'V.1'
planeacion_primaria.insert(11, columna_a_mover, planeacion_primaria.pop(columna_a_mover))

# Aca se filtran solo las columnas para el horario para al final (en la sugunda pagina del pdf) solo filtrar el estudiante y crear la tabla

columnas_a_conservar = ['Estudiante', 'L', 'L.1', 'M', 'M.1', 'W', 'W.1', 'J', 'J.1', 'V', 'V.1']

planeacion_primaria = planeacion_primaria[columnas_a_conservar]

# Función para modificar el nombre de las columnas y que queden L.1, L.2,M.1,M.2,....
def modificar_nombre(columna):
    if columna.endswith('.1'):
        return columna.replace('.1', '.2')  # Cambiar el '1' por '2'
    elif columna in ['L', 'M', 'W', 'J', 'V']:
        return columna + '.1'  # Agregar '.1' a las letras
    else:
        return columna  # No cambiar el resto

# Aplicar la función a todos los nombres de las columnas
planeacion_primaria.columns = [modificar_nombre(col) for col in planeacion_primaria.columns]


# BACHILLERATO

# ACA SE MUEVEN LAS COLUMNAS DEL HORARIO PARA QUE QUEDEN COMO EN LA APGINA QUE SE PUBLICABA EN TEAMS
# ES DECIR QUE QUEDEN LOS DOS MODULOS DE CADA DIA JUNTOS 

columna_a_mover = 'L.1'
planeacion_bachillerato.insert(3, columna_a_mover, planeacion_bachillerato.pop(columna_a_mover))
columna_a_mover = 'M.1'
planeacion_bachillerato.insert(5, columna_a_mover, planeacion_bachillerato.pop(columna_a_mover))
columna_a_mover = 'W.1'
planeacion_bachillerato.insert(7, columna_a_mover, planeacion_bachillerato.pop(columna_a_mover))
columna_a_mover = 'J.1'
planeacion_bachillerato.insert(9, columna_a_mover, planeacion_bachillerato.pop(columna_a_mover))
columna_a_mover = 'V.1'
planeacion_bachillerato.insert(11, columna_a_mover, planeacion_bachillerato.pop(columna_a_mover))

# Aca se filtran solo las columnas para el horario para al final (en la sugunda pagina del pdf) solo filtrar el estudiante y crear la tabla

columnas_a_conservar = ['Estudiante', 'L', 'L.1', 'M', 'M.1', 'W', 'W.1', 'J', 'J.1', 'V', 'V.1']

planeacion_bachillerato = planeacion_bachillerato[columnas_a_conservar]

# Función para modificar el nombre de las columnas
def modificar_nombre(columna):
    if columna.endswith('.1'):
        return columna.replace('.1', '.2')  # Cambiar el '1' por '2'
    elif columna in ['L', 'M', 'W', 'J', 'V']:
        return columna + '.1'  # Agregar '.1' a las letras
    else:
        return columna  # No cambiar el resto

# Aplicar la función a todos los nombres de las columnas
planeacion_bachillerato.columns = [modificar_nombre(col) for col in planeacion_bachillerato.columns]


# Corregir nombres y asignaturas de todos los formatos
estudiantes['ESTUDIANTE'] = estudiantes['ESTUDIANTE'].apply(corregir_nombre)
estudiantes['GRADO'] = estudiantes['GRADO'].apply(corregir_grado)
notas['ESTUDIANTE'] = notas['ESTUDIANTE'].apply(corregir_nombre)
notas['FECHA'] = pd.to_datetime(notas['FECHA'], errors='coerce')

# Crear la lista de estudiantes para luego iterar sobre ellas
lista_estudiantes = estudiantes['ESTUDIANTE'].tolist()


# Crear el calendario en el que tenemos las semanas del periodo (mismo de cargar notas), cambiar cada periodo (AÑO,MES,DIA)
inicio = datetime.date(2025, 4, 14)
fin = datetime.date(2025, 12, 31)
rango_fechas = pd.date_range(inicio, fin)

cal = pd.DataFrame({'Fecha': rango_fechas, 'Día de la Semana': rango_fechas.day_name()})
df_semana = pd.DataFrame({'Número': np.repeat(range(1, 11), 7)})

calendario = pd.concat([cal, df_semana], ignore_index=True, axis=1)

calendario = calendario.rename(columns={
    0: 'FECHA',
    1: 'DIA',
    2: 'SEMANA'
})

calendario_por_semana = {}  

# Crear un diccionario en el que la entrada es la semana

for i in range(1, 11):
    calendario_semana_i = calendario[calendario.iloc[:, 2] == i]
    calendario_por_semana[i] = calendario_semana_i

semanas = [f"SEMANA {i}" for i in range(1, 11)] + ["TOTAL"]


asignaturas_1_5= ['Biología','Química','Medio ambiente','Física',
                  'Historia', 'Geografía', 'Participación política','Pensamiento religioso',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Inglés - listening','Inglés - speaking','Inglés - writing', 'Inglés - reading',
                  'Aritmética','Animaplanos','Estadística', 'Geometría', 'Dibujo técnico', 'Sistemas']


asignaturas_6_7= ['Biología','Química','Medio ambiente','Física',
                  'Historia', 'Geografía', 'Participación política','Filosofía',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Inglés - listening','Inglés - speaking','Inglés - writing', 'Inglés - reading',
                  'Aritmética','Animaplanos','Estadística', 'Geometría', 'Dibujo técnico', 'Sistemas']

asignaturas_8_9= ['Biología','Química','Medio ambiente','Física',
                  'Historia', 'Geografía', 'Participación política','Filosofía',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Inglés - listening','Inglés - speaking','Inglés - writing', 'Inglés - reading',
                  'Álgebra','Animaplanos', 'Estadística', 'Geometría', 'Dibujo técnico', 'Sistemas']

asignaturas_10=  ['Biología','Química','Medio ambiente','Física',
                  'Metodología','Ciencias económicas', 'Ciencias políticas','Filosofía',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Inglés - listening','Inglés - speaking','Inglés - writing', 'Inglés - reading',
                  'Trigonometría','Animaplanos', 'Estadística', 'Matemática financiera', 'Dibujo técnico', 'Sistemas']

asignaturas_11=  ['Química','Medio ambiente','Física',
                  'Metodología','Ciencias económicas', 'Ciencias políticas','Filosofía',
                  'Comunicación y sistemas simbólicos','Producción e interpretación de textos',
                  'Inglés - listening','Inglés - speaking','Inglés - writing', 'Inglés - reading',
                  'Cálculo','Animaplanos', 'Estadística', 'Matemática financiera', 'Dibujo técnico', 'Sistemas']
    


# Generacion de F10

for estudiante in lista_estudiantes:

    #Borra la variable horario al iniciar cada iteracion
    if 'horario' in locals(): 
        del horario

    #aca se mira si esta en la planeacion de primaria o bachillerato para poner el horario en la segunda pagina del pdf (mas adelante)

    if estudiante in planeacion_primaria['Estudiante'].values:
        horario = planeacion_primaria[planeacion_primaria['Estudiante'] == estudiante]
        horario = horario.drop(columns=["Estudiante"])
    
    if estudiante in planeacion_bachillerato['Estudiante'].values:
        horario = planeacion_bachillerato[planeacion_bachillerato['Estudiante'] == estudiante]
        horario = horario.drop(columns=["Estudiante"])


    print(estudiante)
    grado = estudiantes.loc[estudiantes['ESTUDIANTE'] == estudiante, 'GRADO'].values[0]
    grado = int(grado)
    correo_institucional = estudiantes.loc[estudiantes['ESTUDIANTE'] == estudiante, 'CORREO'].values[0]
    F10 = pd.DataFrame({
    "SEMANA": semanas,
    "DESEMPEÑOS ALCANZADOS": [0] * 11,
    "DESEMPEÑOS FALTANTES": [0] * 11,
    })
    total = 0
    for i in range(1,11):
        notas_semana = notas[(notas['ESTUDIANTE'] == estudiante) & notas['FECHA'].isin(calendario_por_semana[i]['FECHA']) & (notas['CALIFICACIÓN'] != 'H') ]
        desempenos_realizados = len(notas_semana)
        total += desempenos_realizados
        F10.loc[i-1, 'DESEMPEÑOS ALCANZADOS'] = desempenos_realizados
        F10.loc[i-1, 'DESEMPEÑOS FALTANTES'] = estudiantes.loc[estudiantes['ESTUDIANTE'] == estudiante, 'META'].values[0] - total

    F10.loc[10, 'DESEMPEÑOS ALCANZADOS'] = total
    F10.loc[10, 'DESEMPEÑOS FALTANTES'] = estudiantes.loc[estudiantes['ESTUDIANTE'] == estudiante, 'META'].values[0] - total
    F10['DESEMPEÑOS ALCANZADOS'].astype(int)
    F10['DESEMPEÑOS FALTANTES'].astype(int)

    # Crear diagrama circular
    categorias = ['DESEMPEÑOS ALCANZADOS', 'DESEMPEÑOS FALTANTES']
    valores = [total, max(0,F10.loc[10, 'DESEMPEÑOS FALTANTES']) ]

    #Aqui todos los if son para generar el F5_2

    if 1 <= grado <= 5:
        F5_2 = pd.DataFrame(np.full((len(asignaturas_1_5), 20), "", dtype=str), index=asignaturas_1_5)
        largo = {}
        for asignatura,_ in F5_2.iterrows():
            notas_asi = notas[ (notas['ESTUDIANTE'] == estudiante) & (notas['GRADO'] == grado) & (notas['ASIGNATURA'] == asignatura) ] #(notas['ESTUDIANTE'] == estudiante) & (notas['GRADO'] == grado) & (notas['ASIGNATURA'] == asignatura)
            largo[asignatura] = len(notas_asi)
        for asignatura in asignaturas_1_5:
            cantidad_notas = largo[asignatura]
            F5_2.iloc[asignaturas_1_5.index(asignatura), :min(20, cantidad_notas)] = "✔️"
    
    if 6 <= grado <= 7:
        F5_2 = pd.DataFrame(np.full((len(asignaturas_6_7), 20), "", dtype=str), index=asignaturas_6_7)
        largo = {}
        for asignatura,_ in F5_2.iterrows():
            notas_asi = notas[ (notas['ESTUDIANTE'] == estudiante) & (notas['GRADO'] == grado) & (notas['ASIGNATURA'] == asignatura) ] #(notas['ESTUDIANTE'] == estudiante) & (notas['GRADO'] == grado) & (notas['ASIGNATURA'] == asignatura)
            largo[asignatura] = len(notas_asi)
        for asignatura in asignaturas_6_7:
            cantidad_notas = largo[asignatura]
            F5_2.iloc[asignaturas_6_7.index(asignatura), :min(20, cantidad_notas)] = "✔️"
    
    if 8 <= grado <= 9:
        F5_2 = pd.DataFrame(np.full((len(asignaturas_8_9), 20), "", dtype=str), index=asignaturas_8_9)
        largo = {}
        for asignatura,_ in F5_2.iterrows():
            notas_asi = notas[(notas['ESTUDIANTE'] == estudiante) & (notas['GRADO'] == grado) & (notas['ASIGNATURA'] == asignatura) ]
            largo[asignatura] = len(notas_asi)
        for asignatura in asignaturas_8_9:
            cantidad_notas = largo[asignatura]
            F5_2.iloc[asignaturas_8_9.index(asignatura), :min(20, cantidad_notas)] = "✔️"

    if grado == 10:
        F5_2 = pd.DataFrame(np.full((len(asignaturas_10), 20), "", dtype=str), index=asignaturas_10)
        largo = {}
        for asignatura,_ in F5_2.iterrows():
            notas_asi = notas[(notas['ESTUDIANTE'] == estudiante) & (notas['GRADO'] == grado) & (notas['ASIGNATURA'] == asignatura) ]
            largo[asignatura] = len(notas_asi)
        for asignatura in asignaturas_10:
            cantidad_notas = largo[asignatura]
            F5_2.iloc[asignaturas_10.index(asignatura), :min(20, cantidad_notas)] = "✔️"

    if grado == 11:
        F5_2 = pd.DataFrame(np.full((len(asignaturas_11), 20), "", dtype=str), index=asignaturas_11)
        largo = {}
        for asignatura,_ in F5_2.iterrows():
            notas_asi = notas[(notas['ESTUDIANTE'] == estudiante) & (notas['GRADO'] == grado) & (notas['ASIGNATURA'] == asignatura) ]
            largo[asignatura] = len(notas_asi)
        for asignatura in asignaturas_11:
            cantidad_notas = largo[asignatura]
            F5_2.iloc[asignaturas_11.index(asignatura), :min(20, cantidad_notas)] = "✔️"

    # aqui agregamos la primer linea

    fila_asignatura = pd.DataFrame([["A"] * 5 + ["B"] * 5 + ["C"] * 5 + ["D"] * 5], 
                               index=["ASIGNATURA"], 
                               columns=F5_2.columns)

    # Concatenamos 
    F5_2 = pd.concat([fila_asignatura, F5_2])



    # Crear copia del DataFrame para trabajar
    F5_2_modificado = F5_2.copy()

    # Insertar columnas en posiciones específicas (en orden inverso para no afectar los índices)
    F5_2_modificado.insert(15, 'NuevaCol15', ['']*len(F5_2_modificado))
    F5_2_modificado.insert(10, 'NuevaCol10', ['']*len(F5_2_modificado))
    F5_2_modificado.insert(5, 'NuevaCol5', ['']*len(F5_2_modificado))
    F5_2_modificado.insert(0, 'NuevaCol1', ['']*len(F5_2_modificado))


    # CREAR PDF 
    nombre_archivo = f"C:/Users/Admin/Desktop/GK/Informes semanales/{estudiante}.pdf"

    
    with PdfPages(nombre_archivo) as pdf:

        ###################### PRIMERA PAGINA ############################

        # Crear figura con dos subgráficos (1 fila, 2 columnas)
        fig, axs = plt.subplots(1, 2, figsize=(12, 12), gridspec_kw = {'width_ratios':[1,1]})  # 12x6 pulgadas para acomodar ambos
        plt.subplots_adjust(wspace=0.185)

        # LOGOOOOOOS:
        logo_color = plt.imread("C:/Users/Admin/Desktop/GK/Logos/logo color.png")
        colombia_excelente = plt.imread("C:/Users/Admin/Desktop/GK/Logos/COLOMBIA EXCELENTE.png")
        franja_naranja_izquierda = plt.imread("C:/Users/Admin/Desktop/GK/Logos/FRANJA NARANJA IZQUIERDA.png")
        franja_naranja_derecha = plt.imread("C:/Users/Admin/Desktop/GK/Logos/FRANJA NARANJA DERECHA.png")
        pifi = plt.imread("C:/Users/Admin/Desktop/GK/Logos/PIFI.png")

        logo_ax = fig.add_axes([0.05, 0.82, 0.13, 0.13], anchor='NW', zorder=10)  # Cambiado a esquina izquierda
        logo_ax.imshow(logo_color)
        logo_ax.axis('off')

        logo_der = fig.add_axes([0.70, 0.70, 0.25, 0.25], anchor='NE', zorder=10)
        logo_der.imshow(colombia_excelente)
        logo_der.axis('off')

        franja_izq = fig.add_axes([0, 0, 0.2, 0.05], anchor='SW', zorder=10)  # x0, y0, ancho, alto
        franja_izq.imshow(franja_naranja_izquierda)
        franja_izq.axis('off')

        franja_der = fig.add_axes([0.8, 0, 0.2, 0.05], anchor='SE', zorder=10)  # x0, y0, ancho, alto
        franja_der.imshow(franja_naranja_derecha)
        franja_der.axis('off')

        franja_centro = fig.add_axes([0.44, 0.01, 0.12, 0.022], anchor='S', zorder=10)  # ¡Centrado exacto!
        franja_centro.imshow(pifi)
        franja_centro.axis('off')

        #TITULOOOOOOO

        fig.suptitle(f"""        BOGOTÁ D.C, {hoy}
                     
        Buenos días,

        Reciba un cordial saludo de parte del Colegio Gimnasio Kaiporé. A continuación, 
        encontrará el informe de {estudiante}.

        En esta primera página se presenta el reporte semanal de los desempeños realizados, y en la segunda 
        se detalla el avance del grado actual junto con el horario de la semana {semana_actual}.""", 
        fontsize=13, 
        y=0.75,
        x=0.02,      # Posición horizontal en coordenadas normalizadas (0 es el extremo izquierdo)
        ha='left'   # Alineación horizontal a la izquierda
        )

        # Crear la tabla en el primer subgráfico
        axs[0].axis('tight')
        axs[0].axis('off')
        tabla = axs[0].table(cellText=F10.values, colLabels=F10.columns, cellLoc='center', loc='center', bbox=[-0.12, 0.17,1.3, 0.4])

        # Ajustar el tamaño de la fuente en la tabla y alto de las columnas
        tabla.auto_set_font_size(False)
        tabla.set_fontsize(10)  # Ajusta el tamaño del texto en la tabla
        for (i, j), cell in tabla.get_celld().items():
            if j == 0:  #Primera columna
                cell.set_width(0.3)  # Ancho para la primera columna
            else:  # Otras columnas
                cell.set_width(0.5)  # Ancho menor para las demás columnas
            cell.set_height(0.04)  # Ajusta el alto de cada celda
            cell.set_linewidth(1.5)  # Ajusta el grosor del borde de la tabla

        # Crear el gráfico de pastel en el segundo subgráfico
        axs[1].pie(valores, autopct='%1.1f%%', startangle=90, radius = 0.65, textprops={'fontsize': 10}, center=(0, -1.1))

        # Ajustar los límites del eje para que el pie se vea correctamente
        axs[1].set_xlim(-1.4, 1.5)     # Ajusta según necesidad
        axs[1].set_ylim(-1.8, 1.2)     # Más espacio abajo para el pie
        axs[1].legend(categorias, loc = 'lower center', fontsize = 7, ncol = 2, frameon = False, bbox_to_anchor=(0.5, -0.15))
        axs[1].set_title(f"DESEMPEÑOS P{periodo_actual}", fontsize=11, y=0.52)  # Ajustar tamaño del título

        axs[1].set_frame_on(True)  # Activa el marco
        for spine in axs[1].spines.values():
            spine.set_edgecolor('black')  # Color del borde
            spine.set_linewidth(1.5)  # Grosor del borde
        
        axs[1].spines['top'].set_position(('outward', -112))  # Mueve la línea superior
        axs[1].spines['bottom'].set_position(('outward', 61))  # Mueve la línea inferior
        axs[1].spines['left'].set_visible(False)  # Oculta el eje derecho 

        right_spine = axs[1].spines['right']
        # Acortar el spine desde abajo (extenderlo hacia y=-1.5)
        right_spine.set_bounds(-2.37, 0.14)  # ¡Ajusta estos valores!
        

        # Guardar la figura en el PDF
        pdf.savefig(fig)

        # Cerrar la figura para liberar memoria
        plt.close(fig)

        #######################################################################################  SEGUNDA PAGINA #################################

        fig2, ax = plt.subplots(figsize=(12, 8))  # Crear la figura sin subgráficos
        ax.axis("off")  # Ocultar los ejes

        # LOGOOOOOOS:
        logo_color = plt.imread("C:/Users/Admin/Desktop/GK/Logos/logo color.png")
        colombia_excelente = plt.imread("C:/Users/Admin/Desktop/GK/Logos/COLOMBIA EXCELENTE.png")
        franja_naranja_izquierda = plt.imread("C:/Users/Admin/Desktop/GK/Logos/FRANJA NARANJA IZQUIERDA.png")
        franja_naranja_derecha = plt.imread("C:/Users/Admin/Desktop/GK/Logos/FRANJA NARANJA DERECHA.png")
        pifi = plt.imread("C:/Users/Admin/Desktop/GK/Logos/PIFI.png")

        logo_ax = fig2.add_axes([0.05, 0.77, 0.18, 0.18], anchor='NW', zorder=10)  # Cambiado a esquina izquierda
        logo_ax.imshow(logo_color)
        logo_ax.axis('off')

        logo_der = fig2.add_axes([0.70, 0.70, 0.25, 0.25], anchor='NE', zorder=10)
        logo_der.imshow(colombia_excelente)
        logo_der.axis('off')

        franja_izq = fig2.add_axes([0, 0, 0.25, 0.08], anchor='SW', zorder=10)  # x0, y0, ancho, alto
        franja_izq.imshow(franja_naranja_izquierda)
        franja_izq.axis('off')

        franja_der = fig2.add_axes([0.75, 0, 0.25, 0.08], anchor='SE', zorder=10)  # x0, y0, ancho, alto
        franja_der.imshow(franja_naranja_derecha)
        franja_der.axis('off')

        franja_centro = fig2.add_axes([0.4, 0.01, 0.2, 0.03], anchor='S', zorder=10)
        franja_centro.imshow(pifi)
        franja_centro.axis('off')


        # Convertir el DataFrame en tabla y agregarlo al gráfico
        tabla = ax.table(cellText=F5_2_modificado.values, 
                        rowLabels=F5_2_modificado.index, 
                        cellLoc='center', 
                        loc='center',
                        bbox=[0.26, 0.1, 0.72, 0.7])
        
        

        tabla.auto_set_font_size(False)  # Permitir ajuste manual del tamaño de fuente
        tabla.set_fontsize(10)  # Ajustar tamaño de fuente
        tabla.scale(1.2, 1.2)  # Ajustar tamaño de la tabla

        for (i, j), cell in tabla.get_celld().items():
            cell.set_width(0.03)
            cell.set_linewidth(0.8)  # Grosor por defecto para todos los bordes
            cell.set_edgecolor('black')  # Asegurar que todos los bordes sean visibles

            if j==0 or j % 6 ==0 :  
                cell.set_edgecolor('black')  # Asegurar que el borde sea visible
                cell.visible_edges = 'L'  # Mostrar solo el borde izquierdo (Left)
        
        

        ############################################################## HORARIO EN LA PARTE FINAL DE LA SEGUNDA PAGINA ##############3

        try:
            horario
        except NameError:
            pdf.savefig(fig2)
            plt.close(fig2)
            pass  # Si no existe, no hace nada y sigue el código
        else:
            tabla2 = ax.table(
                cellText=horario.values,
                colLabels=horario.columns,   # Mostrar nombres de columnas
                rowLabels=None,              # No mostrar nombres de filas
                cellLoc='center',
                loc='center',
                bbox=[0.01, 0.01, 0.93, 0.06] # x,y,ancho,alto
            )
            
            tabla2.auto_set_font_size(False)
            tabla2.set_fontsize(6)
            tabla2.scale(1.2, 1.2)

            for pos, cell in tabla2.get_celld().items():
                cell.set_width(0.8)       # Ancho menor para las columnas
                cell.set_height(0.04)     # Ajusta el alto de cada celda
                cell.set_linewidth(0.8)   # Ajusta el grosor del borde de la tabla

            pdf.savefig(fig2)
            plt.close(fig2)



    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()  # Habilitar TLS
        server.login(EMAIL_USER, EMAIL_PASSWORD)

        msg = MIMEMultipart()
        msg["From"] = EMAIL_USER
        msg["To"] = correo_institucional
        msg["Subject"] = f'Informe semanal de {estudiante}, semana {semana_actual} '
        msg.attach(MIMEText(
            'Cordial saludo.\n\n'
            'Enviamos el reporte semanal correspondiente.\n\n'
            'Agradecemos su atención.\n\n'
            'Atentamente,\n'
            'Gimnasio Kaiporé',
            "plain"
        ))
        # Ruta del archivo PDF generado
        ruta_pdf = f"C:/Users/Admin/Desktop/GK/Informes semanales/{estudiante}.pdf"
        # Adjuntar el archivo PDF
        with open(ruta_pdf, "rb") as archivo_pdf:
            adjunto = MIMEApplication(archivo_pdf.read(), _subtype="pdf")
            adjunto.add_header("Content-Disposition", f"attachment; filename={estudiante}.pdf")
            msg.attach(adjunto)

        server.sendmail(EMAIL_USER,  correo_institucional, msg.as_string())
        print(f"Correo enviado a {correo_institucional}")

        server.quit()
    except Exception as e:
        print(f"Error al enviar correos: {e}")



   

    

