Attribute VB_Name = "Module1"

'declaramos el token como constante
Private Const inegi_token As String = "[Token]"

Function API_INEGI(clave)

    'esta funcion ejecuta la peticion de data mediante la API

    'url de consulta, revisar los parametros en la pagina de API-INEGI
    url = "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR/" & clave & "/es/0700/false/BIE/2.0/" & inegi_token & "?type=xml"

    'declaramos el objeto para hacer la conexion
    Set solicitud = CreateObject("MSXML2.ServerXMLHTTP")
    
    'realizamos la peticion de la informacion
    solicitud.Open "GET", url, False
    solicitud.Send

    'guardamos la respuesta
    Set respuesta = CreateObject("MSXML2.DOMDocument")
    respuesta.LoadXML solicitud.responseText
    
    'verificamos el contenido de la respuesta
    'MsgBox solicitud.responseText

    'filtramos la respuesta y la guardamos como el resultado de la funcion 
    Set API_INEGI = respuesta.getElementsByTagName("Observation")
    
    'borramos los datos guardados en la solicitud y repuesta
    Set solicitud = Nothing
    Set respuesta = Nothing

End Function

Sub pib()

'esta macro usa la clave "735904" de INEGI para obtener los datos de variacion % del PIB
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
