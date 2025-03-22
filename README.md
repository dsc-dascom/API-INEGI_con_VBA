

# **API-INEGI**

Este proyecto está diseñado para apoyar en la consulta recurrente de información económica de INEGI. Para lograr esto, se usa la API que INEGI proporciona a los usuarios. Este programa está hecho en VBA porque pretendo que todo el proceso de recolección, tratamiento, análisis y presentación de datos se realice dentro de un ambiente conocido por la mayoría de usuarios: Excel.

---

<p align="center"> <img src="https://github.com/user-attachments/assets/e24fa025-ef66-49a3-8eb0-255a33d63e8e" alt="Logo INEGI" width="600"> </p>

---

## **Documentación**
- Recomiendo revisar la página de INEGI que explica aspectos relevantes de su API, como parámetros de la consulta: https://www.inegi.org.mx/servicios/api_indicadores.html

- De igual forma, en el siguiente link directo se puede solicitar el **Token**: https://www.inegi.org.mx/app/desarrolladores/generatoken/Usuarios/token_Verify

- Por último, aclaro que personalmente prefiero usar la página del [BIE de INEGI](https://www.inegi.org.mx/app/indicadores/default.aspx?tm=0) para buscar la información y obtener las claves (identificadores) de los datos. 

- Durante el desarrollo del código se utiliza el objeto MSXML2 para trabajar con los datos en formato XML. Para más información se puede consultar las siguientes páginas: [ServerXMLHTTP](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762278(v=vs.85)) y [DOMDocument](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757828(v=vs.85)).
  
---


## **Uso de Module1**

En el archivo **Module1.bas** se puede encontrar el Módulo de VBA que contiene el código para utilizar la API de INEGI. 

Este código se puede dividir en tres secciones. En la primera se declara el **Token** como una constante.

    Private Const inegi_token As String = "[Token]"
<br>

En la segunda parte se construye una función que realiza el procedimiento de consulta y almacenamiento de información. Esta función depende de una única variable (_"clave"_), la cual se define en la subrutina.

    Function API_INEGI(clave)
<br>

En la última sección se crea una subrutina donde se declara a la variable _"clave"_ y se utiliza como insumo en la función. Posteriormente se imprimen los datos en Excel.  

    Sub pib()
<br>

La lógica del código permite aumentar el número de consultas al declarar una lista con claves que junto con un ciclo for permitirá realizar varias consultas de información.

También es posible crear varias subrutinas que se pueden ejecutar en distintas hojas de Excel. Se puede insertar un botón con la macro asignada y realizar las consultas repetidas veces.

<p align="center"> <img src="https://github.com/user-attachments/assets/76976700-d78c-4adb-879f-e04c27b9234a" alt="ejemplo1" width="400"> </p>
<br>

<p align="center"> <img src="https://github.com/user-attachments/assets/3b56d5e4-6913-4c93-881c-00911e303815" alt="ejemplo2" width="400"> </p>
<br>


Todo depende de las necesidades del proyecto o de las especificaciones de los usuarios para hacerlo más fácil de manejar. 

---

## Objetivos de este proyecto

- Optimizar y automatizar la consulta de información que publica periódicamente INEGI.
  
- Facilitar la comprensión de métodos y objetos en VBA para el uso de la API de INEGI.
  
- Construir un código que sirva como base para un proyecto que involucre analizar las condiciones económicas de México con información  disponible en INEGI.


