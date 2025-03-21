

# **API-INEGI**

Este proyecto está diseñado para apoyar en la consulta recurrente de información económica de INEGI. Para lograr esto, se usa la API que INEGI proporciona a los usuarios. Para mantener todo el proceso de tratamiento de la información en un ambiente conocido para la mayoría de personas, este programa está hecho en VBA. Con lo anterior intento mantener el proceso de consulta y análisis de datos dentro de Excel.



![INEGI-2020](https://github.com/user-attachments/assets/e24fa025-ef66-49a3-8eb0-255a33d63e8e)


---

## **Documentación**
- Recomiendo revisar la página de INEGI que explica aspectos releventes de su API, como parametros de la consulta de infomración: https://www.inegi.org.mx/servicios/api_indicadores.html

- De igual forma, en el siguiente link se puede solicitar el Token: https://www.inegi.org.mx/app/desarrolladores/generatoken/Usuarios/token_Verify

- Durante el desarrollo del código se utiliza el objeto MSXML2 para trabajar con los datos en formato XML. Para más información se puede consultar las siguientes páginas: [DOMDocument](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms766564(v=vs.85)) y [ServerXMLHTTP](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762278(v=vs.85)).
  
---


## **Uso de Module1**

En el archivo _Module1.bas_ se puede encontrar el Módulo de VBA que contiene el código para utilizar la API de INEGI. 

Para poder usar el código es necesario pegar el Token al inicio del código. Específicamente, en la línea donde se declara la variable _inegi_token_ se debe pegar el Token que se obtiene de la página del INEGI.

---


