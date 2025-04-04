

# **API-INEGI**

Este proyecto está diseñado para apoyar en la consulta recurrente de información económica de INEGI. Para lograr esto, se usa la API que INEGI proporciona a los usuarios. Este programa está hecho en VBA porque pretendo que todo el proceso de recolección, tratamiento, análisis y presentación de datos se realice dentro de un ambiente conocido por la mayoría de usuarios: Excel.

---

<p align="center"> <img src="https://github.com/user-attachments/assets/e24fa025-ef66-49a3-8eb0-255a33d63e8e" alt="Logo INEGI" width="600"> </p>

---

## **Documentación**   :open_file_folder:
- Recomiendo revisar la página de INEGI que explica aspectos relevantes de su API, como parámetros de la consulta: https://www.inegi.org.mx/servicios/api_indicadores.html

- De igual forma, en el siguiente link directo se puede solicitar el **Token**: https://www.inegi.org.mx/app/desarrolladores/generatoken/Usuarios/token_Verify

- Por último, aclaro que personalmente prefiero usar la página del [Banco de Información Económica (BIE) de INEGI](https://www.inegi.org.mx/app/indicadores/default.aspx?tm=0) para buscar la información y obtener las claves (identificadores) de los datos. 

- Durante el desarrollo del código se utiliza el objeto MSXML2 para trabajar con los datos en formato XML. Para más información se puede consultar las siguientes páginas: [ServerXMLHTTP](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762278(v=vs.85)) y [DOMDocument](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757828(v=vs.85)).

<br>

**IMPORTANTE:** para poder trabajar con los objetos que se declaran en el código, es indispensable habilitar la referencia **Microsoft XML, v6.0** de la pestaña de herramientas (Tools) dentro de VBA.

<p align="center"> <img src="https://github.com/user-attachments/assets/72cf2a48-960e-49d3-90e2-7cd54db54320" alt="ejemplo2" width="650"> </p>
<p align="center"> <img src="https://github.com/user-attachments/assets/81b02e24-eb84-4602-be7f-b535f4ad26ec" alt="ejemplo2" width="400"> </p>


---


## **Uso de Module1**    :package:

En el archivo **Module1.bas** se puede encontrar el Módulo de VBA que contiene el código para utilizar la API de INEGI. 

Este código se puede dividir en tres secciones. 

- En la primera parte se declara el **Token** como una constante privada, lo que permite que esté disponible en todo el módulo.

      Private Const inegi_token As String = "[Token]"

El **Token** lo pueden generar ingresando en el segundo link en la documentación. Una vez que ingresen a la página deben introducir un correo electrónico en el que INEGI enviará el **Token**.
<p align="center"> <img src="https://github.com/user-attachments/assets/9525cb45-0346-4eaf-af80-f10cdaf3f81e" alt="ejemplo2" width="650"> </p>

<br>

- En la segunda parte se construye una función que realiza el procedimiento de consulta y almacenamiento de información. Esta función depende de una única variable (_"clave"_), la cual se define en la subrutina posterior.

      Function API_INEGI(clave)
<br>

- En la última sección se crea una subrutina donde se declara a la variable _"clave"_ y se utiliza como insumo en la función. Posteriormente se imprimen los datos en Excel.  

      Sub pib()

Para conocer la _"clave"_ de los datos pueden apoyarse del [Constructor de Consultas](https://www.inegi.org.mx/servicios/api_indicadores.html) de la página de la API de INEGI o pueden acceder a la pestaña de Metadados al revisar los indicadores directamente del [BIE](https://www.inegi.org.mx/app/indicadores/default.aspx?tm=0#D735904_1000024201150120):

<p align="center"> <img src="https://github.com/user-attachments/assets/c0a12687-a2c3-4dd3-9b3a-f911d94e1e81" alt="ejemplo2" width="650"> </p>

<br>

Por lo tanto, una vez declarado el **Token** y la _clave_ dentro del código, se puede consultar la información usando la API de INEGI. Recomiendo que asignen a un botón la subrutina (macro) para tener un control desde Excel.

<p align="center"> <img src="https://github.com/user-attachments/assets/159e8a53-5339-4cd5-b5c5-85c87081cd42" alt="ejemplo2" width="650"> </p>

La lógica del código permite aumentar el número de consultas al declarar una lista con claves que junto con un ciclo for permitirá realizar varias consultas de información.

También es posible crear varias subrutinas que se pueden ejecutar en distintas hojas de Excel. Se puede insertar un botón con la macro asignada y realizar las consultas repetidas veces.

<p align="center"> <img src="https://github.com/user-attachments/assets/76976700-d78c-4adb-879f-e04c27b9234a" alt="ejemplo1" width="400"> </p>
<br>

<p align="center"> <img src="https://github.com/user-attachments/assets/3b56d5e4-6913-4c93-881c-00911e303815" alt="ejemplo2" width="400"> </p>
<br>


Todo depende de las necesidades del proyecto o de las especificaciones de los usuarios para hacerlo más fácil de manejar. 

---

## Objetivos de este proyecto   :seedling:

- Optimizar y automatizar la consulta de información que publica periódicamente INEGI.
  
- Facilitar la comprensión de métodos y objetos en VBA para el uso de la API de INEGI.
  
- Construir un código que sirva como base para un proyecto que involucre analizar las condiciones económicas de México con información  disponible en INEGI.

---

## Video de apoyo   :computer:

Como una explicación adicional, he subido un video a YouTube en el que explico [cómo usar el código](https://youtu.be/NmETG6jiF0Y) para obtener información económica de INEGI a través  de su API. 
Espero que de esta forma puedan aclararse algunas dudas sobre cómo cargar el **Module1.bas** así como dónde se deben declarar las variables.

De igual forma, cualquier comentario, pregunta o sugerencia son bien recibidas para mejorar el contenido de este proyecto. 



