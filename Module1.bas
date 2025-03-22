Attribute VB_Name = "Module1"

'declaramos el token como constante
Private Const inegi_token As String = "[Token]"

Function API_INEGI(clave)

    'esta funcion ejecuta la peticion de data mediante la API

    'url de consulta, revisar los parametros en la pág. de INEGI
    url = "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR/" & clave & "/es/0700/false/BIE/2.0/" & inegi_token & "?type=xml"

    'objeto para establecer conexion
    Set solicitud = CreateObject("MSXML2.ServerXMLHTTP")
    
    'establecemos el objeto para realizar la conexion
    solicitud.Open "GET", url, False
    solicitud.Send

    'guardamos la respuesta
    Set respuesta = CreateObject("MSXML2.DOMDocument")
    respuesta.LoadXML solicitud.responseText
    
    'verificamos los datos de repuesta
    'MsgBox solicitud.responseText

    'establecemos que la infomacion obtenida sea el resultado de la funcion
    Set API_INEGI = respuesta.getElementsByTagName("Observation")
    
    'borramos los datos guardadados en la solicitud y repuesta
    Set solicitud = Nothing
    Set respuesta = Nothing

End Function

Sub pib()

'esta macro usa la Clave "735904" de INEGI para obtener los datos de variacion % del pib
clave = 735904

'llamamos a la funcion
Set observaciones = API_INEGI(clave)

'escribimos los datos
i = 5   'iniciamos en la fila 5
For Each obs In observaciones
    Cells(i, 2).Value = obs.SelectSingleNode("TIME_PERIOD").Text
    Cells(i, 3).Value = obs.SelectSingleNode("OBS_VALUE").Text
    i = i + 1
Next obs


Range("A1").Select
MsgBox ("Consulta de datos del PIB exitosa")
    
End Sub
