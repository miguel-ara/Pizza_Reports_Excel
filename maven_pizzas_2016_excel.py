import pandas as pd
import signal
import sys
import xml.etree.ElementTree as ET

def handler_signal(signal, frame):
    print("\n\n Interrupción! Saliendo del programa de manera ordenada")
    sys.exit(1)

# Señal de control por si el usuario introduce Ctrl + C para parar el programa
signal.signal(signal.SIGINT, handler_signal)

# Estas dos listas sirven para darle un cierto peso a una pizza según su tamaño. Este peso será el que
# multiplique a los ingredientes de dicha pizza.

MULT = [1, 2, 3, 4, 5]
TAM = ['s','m','l','xl', 'xxl']

def extract():

    # Carga los datos de los csvs correspondientes a los pedidos, las pizzas y los ingredintes de las pizzas

    pedidos = pd.read_csv("order_details.csv", sep = ";", encoding = "UTF-8")
    pizzas = pd.read_csv("pizzas.csv", sep = ",", encoding = "UTF-8")
    ingredientes = pd.read_csv("pizza_types.csv", sep = ",", encoding = "LATIN-1")
    fechas = pd.read_csv("orders.csv", sep = ';', encoding = 'UTF-8')

    return pedidos, pizzas, ingredientes, fechas


def transform(pedidos, pizzas, ingredientes, fechas):

    # Recibe los 4 dataframes, pedidos, fechas, pizzas e ingredientes y va trasnformando los datos para obtener
    # un diccionario con los ingredientes a comprar semanalmente. Primero, genera un informe analizando 
    # cada dataframe extraído (ver función informe_datos) en formato xml

    informe_datos(pedidos, pizzas, ingredientes, fechas)

    # El único de los csvs que contiene Nulls y en el que los datos están mal formateados es el de
    # order_details.csv, el cual hemos cargado en el dataframe de pedidos. Por ello, primero deberemos
    # limpiar este dataframe de Nulls, quitandolos de las columnas pizza_id y quantity. El dataframe de
    # fechas también contiene datos que necesitan ser limpiados, pero como no vamos a trabajar con esos
    # datos, no hace falta limpiarlos.

    pedidos = pedidos[pedidos['pizza_id'].isnull() == False]
    pedidos = pedidos[pedidos['quantity'].isnull() == False]
    pizza_id = []
    quantity = []

    # Ahora procesaremos estas dos columnas del dataframe, remplazando ciertos caracteres por otros
    # para así tener estas dos columnas de manera que se correspondan los nombres de las pizzas
    # con los nombres de las pizzas en los dataframe pizzas e ingredientes. Con ello tratamos de
    # tener los nombres de las pizzas en el formato <nombre_separado_por_guiones_bajos_tamaño_pizza>
    # y las cantidades en números enteros positivos.

    for pizza_sin_procesar in pedidos['pizza_id']:
        pizza = pizza_sin_procesar.replace('@', 'a').replace('3', 'e').replace('0', 'o').replace('-', '_').replace(' ', '_')
        pizza_id.append(pizza)

    for cantidad in pedidos['quantity']:    
        cantidad = cantidad.replace('One', '1').replace('one', '1').replace('two', '2')
        cantidad = abs(int(cantidad))
        quantity.append(cantidad)

    pedidos['pizza_id'] = pizza_id  # Reescribimos las dos columnas del dataframe que acabamos de procesar
    pedidos['quantity'] = quantity
    num_pizzas_sem = {}
    dict_ingredientes = {}

    # Para cada pizza obtenida del dataframe de pizzas sumamos el número de veces que se ha pedido en todo el
    # año, lo dividimos entre 52 (semanas de un año) tomando la parte entera y sumamos 1. De esta manera
    # estamos aproximando por arriba el número de pizzas de cada tipo que se piden en una semana.

    for pizza in pizzas['pizza_id']:
        num_pizzas_sem[pizza] = int(pedidos[pedidos['pizza_id'] == pizza]['quantity'].sum() // 52 + 1) # Redondeamos para arriba

    # Para cada pizza, tomamos sus ingredientes. Pero como quiero tener un diccionario con todos los
    # ingredientes, debo procesar los ingredientes de cada pizza, pasándolos a una lista y luego generando
    # el diccionario con todos los valores a 0 siendo las claves los ingredientes

    for ingrediente_sin_procesar in ingredientes['ingredients']:
        a = ingrediente_sin_procesar.split(', ')
        for ingrediente in a:
            dict_ingredientes[ingrediente] = 0

    # Ahora debo procesar el nombre de cada pizza e ir sumando al diccionario los ingredientes correspondientes
    # (ver las funciones de procesar_nombre_pizza y calcular_ingredientes)

    for pizza_sin_procesar in num_pizzas_sem.keys():
        pizza, multiplicador = procesar_nombre_pizza(pizza_sin_procesar)
        calcular_ingredientes(pizza, pizza_sin_procesar, multiplicador, dict_ingredientes, num_pizzas_sem, ingredientes)

    return dict_ingredientes


def load(dict_ingredientes):

    # Recibe el diccionario con los ingredientes, crea un dataframe con dicho diccionario para
    # y lo imprime por pantalla. Adicionalmente, guarda el dataframe en un fichero csv y genera
    # un archivo xml que contiene esa misma información.

    compra_semanal_ingredientes = pd.DataFrame([[key, dict_ingredientes[key]] for key in dict_ingredientes.keys()], columns=['Ingredient', 'Amount (units)'])
    compra_semanal_ingredientes.to_csv('compra_semanal_ingredientes.csv', index=False)

    # Inicializamos y damos nombre a la raíz de lo que será el fichero xml con la lista completa de ingredientes.

    root = ET.Element("Ingredientes_semanales")
    root.text = 'Cantidad estimada de cada ingrediente que se deberá comprar semanalmente en Pizzas Maven 2016'

    for i in range(len(dict_ingredientes)):      # Para cada dataframe empleado en el programa

        key = list(dict_ingredientes.keys())[i]  # Obtenemos cada clave del diccionario (un ingrediente)
        ingrediente = ET.SubElement(root, "Ingrediente")
        ET.SubElement(ingrediente, "Número").text = str(i)
        ET.SubElement(ingrediente, "Ingrediente", Nombre = str(key))  # Nombre ingrediente
        ET.SubElement(ingrediente, "Cantidad", Unidades = str(dict_ingredientes[key]))  # Cantidad de ese ingrediente

    ET.indent(root)              # Para que se vea bonito
    tree = ET.ElementTree(root)  # Escribimos todos los datos organizados en un fichero xml
    tree.write("compra_semanal_ingredientes.xml", xml_declaration = True, encoding = 'UTF-8')
    
    print(f'\nEn una semana, el manager de Pizzas Maven deberá comprar las siguientes cantidades de ingredientes: \n\n {compra_semanal_ingredientes}')
    print('\nSe han generado dos archivos xml como output del programa, "informe_pizas_maven_2016.xml" como análisis'
            ' previo de los datos, y "compra_semanal_ingredientes.xml" para indicar a Pizzas Maven la cantidad de cada'
            ' ingrediente que se deberá comprar cada semana del año. También se genera un csv "compra_semanal_ingredientes.csv"'
            ' que contiene esa misma información')


def procesar_nombre_pizza(pizza_sin_procesar):

    # Recibe como argumento el nombre de una pizza sin procesar, por ejemplo 'bbq_chicken_l'
    # y devuelve el nombre de la pizza sin el tamaño, es decir, 'bbq_chicken' y un número
    # que representa por cuanto se van a multiplicar los ingredientes de la pizza, siguiendo
    # la correspondencia establecida en la lista MULT y TAM

    pizza = pizza_sin_procesar.split('_')
    tamaño = pizza.pop(-1)
    s = '_'
    pizza = s.join(pizza)
    multiplicador = MULT[TAM.index(tamaño)]
    return (pizza, multiplicador)


def calcular_ingredientes(pizza, pizza_sin_procesar, multiplicador, dict_ingredientes, num_pizzas_sem, ingredientes):

    # Recibe como argumento el nombre de la pizza procesado y sin procesar, el número por el que se van
    # a multiplicar los ingredientes, el diccionario con todos los ingredientes, el diccionario con el
    # número de pizzas que se venden por semana y el dataframe de los ingredientes de las pizzas.

    # Para una pizza, obtengo sus ingredientes, los proceso pasándolos a una lista y en el diccionario
    # de ingredientes, busco cada ingrediente de la pizza y le sumo una cierta cantidad. Esta cantidad
    # será el resultado de multiplicar el número de pizzas de ese tipo que se han pedido en una semana
    # por el tamaño de la pizza traducido a número (el multiplicador)

    ingredientes_pizza = ingredientes[ingredientes['pizza_type_id'] == pizza]['ingredients'].item()
    a = ingredientes_pizza.split(', ')
    for ingredient in a:
        dict_ingredientes[ingredient] += num_pizzas_sem[pizza_sin_procesar]*multiplicador


def informe_datos(pedidos, pizzas, ingredientes, fechas):

    dataframes = [pedidos, pizzas, ingredientes, fechas]
    archivos = ["order_details.csv", "pizzas.csv", "pizza_types.csv", "orders.csv"]

    # df.info() aporta información extra que no nos interesa, mejor obtener a mano solo lo que nos interesa del dataframe.
    # Esto lo haremos para cada dataframe que hemos cargado, y guardaremos todos los resultados en un xml conjunto llamado
    # "informe_pizzas_maven_2016.xml" Cabe destacar que la librería Pandas no distingue entre NaN y Null. Por tanto,
    # Nº NaNs = Nº Nulls en todos los casos.

    # Inicializamos y damos nombre a la raíz de lo que será el fichero xml con el informe completo de los datos.

    root = ET.Element("Informe")
    root.text = 'Análisis de los NaNs, Nulls y la tipología de cada uno de los ficheros empleados en el proceso'

    for i in range(len(dataframes)):  # Para cada dataframe empleado en el programa

        fichero = ET.SubElement(root, "fichero", nombre_fichero = archivos[i])  # Nº de NaNs y Nulls totales
        ET.SubElement(fichero, "NaN", Número_de_NaNs_totales = str(dataframes[i].isna().sum().sum()))
        ET.SubElement(fichero, "Null", Número_de_Nulls_totales = str(dataframes[i].isnull().sum().sum()))
        nombres_columnas = list(dataframes[i].columns.values)  # Obtenemos los nombres de las columnas de un dataframe
        for nombre_col in nombres_columnas:                    # Para cada columna de un dataframe
            columna = ET.SubElement(fichero, "columna", nombre_columna = nombre_col)  # # Nº de NaNs, de Nulls y tipología por columnas
            ET.SubElement(columna, "NaN", Número_de_NaNs_columna = str(dataframes[i][nombre_col].isna().sum()))
            ET.SubElement(columna, "Null", Número_de_Nulls_columna = str(dataframes[i][nombre_col].isnull().sum()))
            ET.SubElement(columna, "Tipología", Tipología_columna = str(dataframes[i][nombre_col].dtype))

    ET.indent(root)              # Para que se vea bonito
    tree = ET.ElementTree(root)  # Escribimos todos los datos organizados en un fichero xml
    tree.write("informe_pizzas_maven_2016.xml", xml_declaration = True, encoding = 'UTF-8')


def ETL(): 

    # ETL incluyendo informe de cada dataframe extraído

    pedidos, pizzas, ingredientes, fechas = extract()
    compra = transform(pedidos, pizzas, ingredientes, fechas)
    load(compra)


if __name__ == '__main__':

    ETL()
