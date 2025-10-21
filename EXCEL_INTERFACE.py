"""
EXCEL_INTERFACE
"""

#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import os
from openai import OpenAI
from dotenv import load_dotenv

# Cargar las variables de entorno desde el archivo .env
load_dotenv()

# Obtener clave API de OpenAI desde variable de entorno
api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key)

def consulta_chatgpt(prompt):
    try:
        response = client.chat.completions.create(
            model="o3-mini-2025-01-31",
            messages=[
                {"role": "system", "content": """Siguiendo las siguientes reglas "Transforma textos de IA en textos humanos, es decir, vamos a humanizar los textos introducidos. Sigue las siguientes reglas:
1.	El largo de las oraciones debe ser mayor que el normal.
2.	Utiliza menos puntos separadores de oraciones y conecta las oraciones utilizando palabras conectoras del idioma español.
3.	Evita la redundancia y la divagación: Expón las ideas de manera directa y sin desviaciones innecesarias.
4.	Usa un lenguaje accesible: Prioriza palabras comunes sobre términos innecesariamente complejos.
5.	Adapta el nivel de lectura a una audiencia universitaria: Usa un lenguaje preciso, sin tecnicismos innecesarios que dificulten la comprensión.
6.	Las frases de tus repuestas deben contener dos oraciones largas
7.	Conecta las ideas de manera natural: Evita transiciones forzadas o artificiosas entre párrafos y secciones.
8.	Sigue las normas convencionales de puntuación: Usa signos de puntuación estándar y evita abusar de los mismos.
9.	Varía la estructura de las oraciones: Alterna la construcción sintáctica para evitar la monotonía en el discurso.
10.	La respuesta debe ser una sola frase con varias oraciones
11.	Evita el uso de conectores innecesarios: Reduce la presencia de términos como "de hecho", "además", "por lo tanto", "asimismo", "sin embargo", "aparentemente", "en consecuencia", "específicamente", "notablemente" y "alternativamente", ya que pueden sobrecargar el texto y restar claridad.
12.	Asegúrate de que las oraciones se mantengan coherentes y fáciles de entender a pesar de su longitud.
13.	Considera el tono y el contexto original del texto al hacer las transformaciones.
Output Format
•	El texto debe ser fluido, imitando el estilo natural y conversacional de un hablante nativo del español, sin perder el significado original. Utilizando menos puntos separadores de oraciones y más palabras de conexión del español
 Notes

Esta es la lista de conectores que debes usar entre oraciones :
Aquí tienes una lista de conectores en español, organizados por categorías:
________________________________________
CONECTORES ADITIVOS (Añaden información)
•	además
•	agregado a lo anterior
•	asimismo
•	de igual forma
•	de igual manera
•	de la misma manera
•	del mismo modo
•	en la misma línea
•	encima
•	es más
•	hasta
•	igualmente
•	incluso
•	más aún
•	para colmo
•	por añadidura
•	también
•	y
________________________________________
CONECTORES DE CONTRASTE Y OPOSICIÓN
Conectores concesivos (Muestran dificultad o contradicción)
•	a pesar de que
•	a pesar de todo
•	ahora bien
•	al mismo tiempo
•	aun así
•	aun cuando
•	aunque
•	con todo
•	de cualquier modo
Conectores restrictivos (Limitan o contrastan una idea)
•	al contrario
•	en cambio
•	en cierta medida
•	en cierto modo
•	hasta cierto punto
•	no obstante
•	pero
•	por otra parte
•	sin embargo
________________________________________
CONECTORES DE CAUSA (Explican razones o motivos)
•	dado que
•	debido a que
•	porque
•	pues
•	puesto que
•	ya que
________________________________________
CONECTORES DE CONSECUENCIA (Expresan efectos o resultados)
•	así pues
•	así que
•	de ahí que
•	de manera que
•	de tal forma
•	en consecuencia
•	en ese sentido
•	entonces
•	luego
•	por consiguiente
•	por eso
•	por esta razón
•	por lo que sigue
•	por lo tanto
•	por tanto
________________________________________
CONECTORES COMPARATIVOS (Establecen semejanzas o diferencias)
•	análogamente
•	así como
•	como
•	de modo similar
•	del mismo modo
•	igual… que…
•	igualmente
•	más/menos… que…
•	tan… como…
________________________________________
CONECTORES REFORMULATIVOS
Conectores de explicación (Aclaran o explican una idea)
•	a saber
•	dicho de otro modo
•	en otras palabras
•	es decir
•	esto es
•	o sea
Conectores de recapitulación (Resumen lo dicho anteriormente)
•	dicho de otro modo
•	en breve
•	en otras palabras
•	en resumen
•	en resumidas cuentas
•	en síntesis
•	en suma
•	en una palabra
•	total
Conectores de ejemplificación (Introducen ejemplos)
•	así como
•	específicamente
•	para ilustrar
•	particularmente
•	por ejemplo
•	tal es el caso de
Conectores de corrección (Corrigen o reformulan una idea)
•	a decir verdad
•	mejor dicho
•	o sea
________________________________________
CONECTORES ORDENADORES
Conectores para iniciar el discurso
•	ante todo
•	bueno
•	en primer lugar
•	en principio
•	para comenzar
•	primeramente
Conectores para cerrar el discurso
•	al final
•	en conclusión
•	en fin
•	en suma
•	finalmente
•	para concluir
•	para finalizar
•	para resumir
•	por último
Conectores de transición
•	a continuación
•	acto seguido
•	ahora bien
•	después / luego
•	por otra parte
•	por otro lado
Conectores de digresión (Introducen una idea secundaria)
•	a propósito
•	a todo esto
•	con respecto a
•	en cuanto a
•	por cierto
•	por otra parte
________________________________________
CONECTORES TEMPORALES (Ubican un evento en el tiempo)
•	a partir de
•	actualmente
•	al principio
•	antes de
•	apenas
•	cuando
•	desde (entonces)
•	después de
•	en cuanto
•	en el comienzo
•	hasta que
•	inmediatamente
•	luego
•	no bien
•	temporalmente
________________________________________
CONECTORES ESPACIALES (Ubican un evento en el espacio)
•	a la izquierda/derecha
•	abajo
•	al lado
•	arriba
•	en el fondo/medio
•	adelante
________________________________________
CONECTORES CONDICIONALES (Expresan condiciones o requisitos)
•	a menos que
•	a no ser que
•	con tal que
•	en caso de que
•	mientras que
•	según
•	si
•	siempre que
•	siempre y cuando
________________________________________
CONECTORES DE CERTEZA (Expresan seguridad o certeza)
•	ciertamente
•	claro está
•	como es por muchos conocido
•	como nadie ignora
•	con certeza
•	con seguridad
•	efectivamente
•	en realidad
•	en verdad
•	es evidente
•	indudablemente
•	realmente
________________________________________
CONECTORES DE FINALIDAD (Expresan el propósito de una acción)
•	a fin de
•	con el objetivo de
•	con el propósito de
•	con la intención de
•	de manera que
•	de tal forma que
•	de modo que
•	para
•	con el objeto de

"""},
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error al conectar con la API de OpenAI: {str(e)}"

def obtener_informacion_funcion():
    # Solicita la información al usuario
    capitulo = "1"
    pestaña = "Inicio"
    grupo = "Portapapeles"
    funcion = "Cortar"

    # Genera el prompt para ChatGPT
    prompt = (f"Tengo una función de Excel que se encuentra en el CAPÍTULO '{capitulo}', "
              f"PESTAÑA '{pestaña}', GRUPO '{grupo}', y se llama '{funcion}'. "
              "Necesito que me expliques lo siguiente:\n"
              "1. ¿Cuál es el propósito de esta función?\n"
              "2. ¿Cómo se encamina esta función en su uso típico?\n"
              "3. Proporciona dos ejemplos de uso práctico de esta función.")

    # Consulta a ChatGPT
    respuesta = consulta_chatgpt(prompt)

    # Muestra la respuesta obtenida
    print("\nRespuesta de ChatGPT:")
    print(respuesta)

# Ejecuta el programa
if __name__ == "__main__":
    obtener_informacion_funcion()


# In[4]:


import os
from openai import OpenAI
from dotenv import load_dotenv

# Cargar las variables de entorno desde el archivo .env
load_dotenv()

# Obtener clave API de OpenAI desde variable de entorno
api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key)

def consulta_chatgpt(prompt):
    """
    Realiza una consulta a la API de OpenAI con el prompt dado.
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": """Siguiendo las siguientes reglas "Transforma textos de IA en textos humanos, es decir, vamos a humanizar los textos introducidos. Sigue las siguientes reglas:
1.	El largo de las oraciones debe ser mayor que el normal.
2.	Utiliza menos puntos separadores de oraciones y conecta las oraciones utilizando palabras conectoras del idioma español.
3.	Evita la redundancia y la divagación: Expón las ideas de manera directa y sin desviaciones innecesarias.
4.	Usa un lenguaje accesible: Prioriza palabras comunes sobre términos innecesariamente complejos.
5.	Adapta el nivel de lectura a una audiencia universitaria: Usa un lenguaje preciso, sin tecnicismos innecesarios que dificulten la comprensión.
6.	Las frases de tus repuestas deben contener dos oraciones largas
7.	Conecta las ideas de manera natural: Evita transiciones forzadas o artificiosas entre párrafos y secciones.
8.	Sigue las normas convencionales de puntuación: Usa signos de puntuación estándar y evita abusar de los mismos.
9.	Varía la estructura de las oraciones: Alterna la construcción sintáctica para evitar la monotonía en el discurso.
10.	La respuesta debe ser una sola frase con varias oraciones
11.	Evita el uso de conectores innecesarios: Reduce la presencia de términos como "de hecho", "además", "por lo tanto", "asimismo", "sin embargo", "aparentemente", "en consecuencia", "específicamente", "notablemente" y "alternativamente", ya que pueden sobrecargar el texto y restar claridad.
12.	Asegúrate de que las oraciones se mantengan coherentes y fáciles de entender a pesar de su longitud.
13.	Considera el tono y el contexto original del texto al hacer las transformaciones.
Output Format
•	El texto debe ser fluido, imitando el estilo natural y conversacional de un hablante nativo del español, sin perder el significado original. Utilizando menos puntos separadores de oraciones y más palabras de conexión del español
 Notes

Esta es la lista de conectores que debes usar entre oraciones :
Aquí tienes una lista de conectores en español, organizados por categorías:
________________________________________
CONECTORES ADITIVOS (Añaden información)
•	además
•	agregado a lo anterior
•	asimismo
•	de igual forma
•	de igual manera
•	de la misma manera
•	del mismo modo
•	en la misma línea
•	encima
•	es más
•	hasta
•	igualmente
•	incluso
•	más aún
•	para colmo
•	por añadidura
•	también
•	y
________________________________________
CONECTORES DE CONTRASTE Y OPOSICIÓN
Conectores concesivos (Muestran dificultad o contradicción)
•	a pesar de que
•	a pesar de todo
•	ahora bien
•	al mismo tiempo
•	aun así
•	aun cuando
•	aunque
•	con todo
•	de cualquier modo
Conectores restrictivos (Limitan o contrastan una idea)
•	al contrario
•	en cambio
•	en cierta medida
•	en cierto modo
•	hasta cierto punto
•	no obstante
•	pero
•	por otra parte
•	sin embargo
________________________________________
CONECTORES DE CAUSA (Explican razones o motivos)
•	dado que
•	debido a que
•	porque
•	pues
•	puesto que
•	ya que
________________________________________
CONECTORES DE CONSECUENCIA (Expresan efectos o resultados)
•	así pues
•	así que
•	de ahí que
•	de manera que
•	de tal forma
•	en consecuencia
•	en ese sentido
•	entonces
•	luego
•	por consiguiente
•	por eso
•	por esta razón
•	por lo que sigue
•	por lo tanto
•	por tanto
________________________________________
CONECTORES COMPARATIVOS (Establecen semejanzas o diferencias)
•	análogamente
•	así como
•	como
•	de modo similar
•	del mismo modo
•	igual… que…
•	igualmente
•	más/menos… que…
•	tan… como…
________________________________________
CONECTORES REFORMULATIVOS
Conectores de explicación (Aclaran o explican una idea)
•	a saber
•	dicho de otro modo
•	en otras palabras
•	es decir
•	esto es
•	o sea
Conectores de recapitulación (Resumen lo dicho anteriormente)
•	dicho de otro modo
•	en breve
•	en otras palabras
•	en resumen
•	en resumidas cuentas
•	en síntesis
•	en suma
•	en una palabra
•	total
Conectores de ejemplificación (Introducen ejemplos)
•	así como
•	específicamente
•	para ilustrar
•	particularmente
•	por ejemplo
•	tal es el caso de
Conectores de corrección (Corrigen o reformulan una idea)
•	a decir verdad
•	mejor dicho
•	o sea
________________________________________
CONECTORES ORDENADORES
Conectores para iniciar el discurso
•	ante todo
•	bueno
•	en primer lugar
•	en principio
•	para comenzar
•	primeramente
Conectores para cerrar el discurso
•	al final
•	en conclusión
•	en fin
•	en suma
•	finalmente
•	para concluir
•	para finalizar
•	para resumir
•	por último
Conectores de transición
•	a continuación
•	acto seguido
•	ahora bien
•	después / luego
•	por otra parte
•	por otro lado
Conectores de digresión (Introducen una idea secundaria)
•	a propósito
•	a todo esto
•	con respecto a
•	en cuanto a
•	por cierto
•	por otra parte
________________________________________
CONECTORES TEMPORALES (Ubican un evento en el tiempo)
•	a partir de
•	actualmente
•	al principio
•	antes de
•	apenas
•	cuando
•	desde (entonces)
•	después de
•	en cuanto
•	en el comienzo
•	hasta que
•	inmediatamente
•	luego
•	no bien
•	temporalmente
________________________________________
CONECTORES ESPACIALES (Ubican un evento en el espacio)
•	a la izquierda/derecha
•	abajo
•	al lado
•	arriba
•	en el fondo/medio
•	adelante
________________________________________
CONECTORES CONDICIONALES (Expresan condiciones o requisitos)
•	a menos que
•	a no ser que
•	con tal que
•	en caso de que
•	mientras que
•	según
•	si
•	siempre que
•	siempre y cuando
________________________________________
CONECTORES DE CERTEZA (Expresan seguridad o certeza)
•	ciertamente
•	claro está
•	como es por muchos conocido
•	como nadie ignora
•	con certeza
•	con seguridad
•	efectivamente
•	en realidad
•	en verdad
•	es evidente
•	indudablemente
•	realmente
________________________________________
CONECTORES DE FINALIDAD (Expresan el propósito de una acción)
•	a fin de
•	con el objetivo de
•	con el propósito de
•	con la intención de
•	de manera que
•	de tal forma que
•	de modo que
•	para
•	con el objeto de

"""},
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error al conectar con la API de OpenAI: {str(e)}"

def obtener_proposito(capitulo, pestaña, grupo, funcion):
    """
    Genera el propósito de la función.
    """
    prompt = (
        f"Tengo una función de Excel que se encuentra en el CAPÍTULO '{capitulo}', "
        f"PESTAÑA '{pestaña}', GRUPO '{grupo}', y se llama '{funcion}'. "
        "Explícame lo siguiente:\n"
        "Cuál es el propósito de esta función"
    )
    respuesta = consulta_chatgpt(prompt)
    return respuesta

def obtener_uso(capitulo, pestaña, grupo, funcion):
    """
    Genera el uso típico de la función.
    """
    prompt = (
        f"Tengo una función de Excel que se encuentra en el CAPÍTULO '{capitulo}', "
        f"PESTAÑA '{pestaña}', GRUPO '{grupo}', y se llama '{funcion}'. "
        "Explícame lo siguiente:\n"
        "Cómo se encamina esta función en su uso típico"
    )
    respuesta = consulta_chatgpt(prompt)
    return respuesta

def obtener_ejemplos(capitulo, pestaña, grupo, funcion):
    """
    Genera dos ejemplos prácticos de uso para la función.
    """
    prompt = (
        f"Tengo una función de Excel que se encuentra en el CAPÍTULO '{capitulo}', "
        f"PESTAÑA '{pestaña}', GRUPO '{grupo}', y se llama '{funcion}'. "
        "Explícame lo siguiente:\n"
        "Proporciona dos ejemplos de uso práctico de esta función."
    )
    respuesta = consulta_chatgpt(prompt)
    return respuesta

def obtener_informacion_funcion():
    """
    Genera las respuestas para propósito, uso y ejemplos de una función
    y las guarda en variables separadas.
    """
    # Información de la función
    capitulo = "1"
    pestaña = "Inicio"
    grupo = "Portapapeles"
    funcion = "Cortar"

    # Obtener cada respuesta mediante funciones específicas
    proposito = obtener_proposito(capitulo, pestaña, grupo, funcion)
    uso = obtener_uso(capitulo, pestaña, grupo, funcion)
    ejemplos = obtener_ejemplos(capitulo, pestaña, grupo, funcion)

    # Mostrar las respuestas
    print("\nInformación generada para la función:")
    print(f"Propósito: {proposito}")
    print(f"Uso: {uso}")
    print(f"Ejemplos: {ejemplos}")

# Ejecuta el programa
if __name__ == "__main__":
    obtener_informacion_funcion()


# In[ ]:




