import re

def corregir_grado(grado):

    grado = str(grado)
    #Usar una expresión regular para eliminar cualquier caracter no numérico (como el '°')
    grado = re.sub(r'[^\d]', '', grado)  # Esto elimina todo lo que no sea un número
    
    #Convertir el valor a un número entero
    try:
        grado = int(grado)
    except ValueError:
        raise ValueError(f"El valor '{grado}' no es un grado válido.")
    
    return grado