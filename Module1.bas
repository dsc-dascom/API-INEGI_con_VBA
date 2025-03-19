Attribute VB_Name = "Module1"
Public Const inegi_token As String = "4102b645-1744-e402-be63-b06edff341ee"

Function API_INEGI(clave)

    'esta funcion ejecuta la peticion de data mediante la API

    'url de consulta, revisar los parametros en la pág. de INEGI
    Url = "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR/" & clave & "/es/0700/false/BIE/2.0/" & inegi_token & "?type=xml"

    'objeto para establecer conexion
    Set solicitud = CreateObject("MSXML2.ServerXMLHTTP")
    
    'enviar solicitud
    solicitud.Open "GET", Url, False
    solicitud.Send

    'guardamos la respuesta
    Set respuesta = CreateObject("MSXML2.DOMDocument")
    respuesta.LoadXML solicitud.responseText
    
    'verificamos los datos de repuesta
    'MsgBox solicitud.responseText

    'la funcion regresa la data
    Set API_INEGI = respuesta.getElementsByTagName("Observation")
    
    'borramos los datos guardadados en la solicitud y repuesta
    Set solicitud = Nothing
    Set respuesta = Nothing

End Function

Sub pib()

'esta macro usa la Clave "735904" de INEGI para obtener los datos de variacion % del pib
clave = 735904

'llamamos a la funcion para obtener la data
Set observaciones = API_INEGI(clave)

'escribimos los datos
i = 5   'iniciamos en la fila 5
For Each obs In observaciones
    Cells(i, 2).Value = obs.SelectSingleNode("TIME_PERIOD").Text
    Cells(i, 3).Value = obs.SelectSingleNode("OBS_VALUE").Text
    i = i + 1
Next obs


Range("I1").Select
MsgBox ("Consulta de datos del PIB exitosa")
    
End Sub
