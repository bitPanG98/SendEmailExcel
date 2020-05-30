Attribute VB_Name = "SendEmail"
Sub sendEmail()

    'Variables a utilizar
    'Dim olApp As Object
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    
    Dim hoja As Worksheet
    
    Dim indexFila As Integer
    Dim nombreTrabajador As String
    Dim emailTrabajador As String
    Dim subjectEmail As String
    Dim archivoEnviar As String
    Dim pathSeparator As String
    
    Dim listaDocumentosFO() As String
    Dim cantidadDocumentosFOTrabajador As Integer
    
    'Obtenemos la hoja de los registros que vamos a  trabajar
    Set hoja = ThisWorkbook.Worksheets("Datos")
    '-
    pathSeparator = Application.pathSeparator
    '-
    archivoEnviar = ThisWorkbook.Path & "\archivos\formatos.rar"
    archivoEnviar = Replace(archivoEnviar, "\", pathSeparator)
    
    With hoja
       
       'indexFila -> Inicializado en 2 para empezar desde el registro 2 que comienzan los nombres para no tomar los titulos
       '.Range -> Aqui recorremos todos los registros de la Columna A
        For indexFila = 2 To .Range("A" & Rows.Count).End(xlUp).Row
            
            listaDocumentosFO = getListaDocumentosTrabajadorFO(indexFila)
            cantidadDocumentosFOTrabajador = getCantidadDocumentosFOTrabajador(listaDocumentosFO)
            
            'Verificamos que solo los trabajadores que tengan documentos(Faltantes-Observados) se les envie el mensaje
            If (cantidadDocumentosFOTrabajador > 0) Then
                
                'Instanciamos el objeto para utilizar Outlook
                Set olApp = New Outlook.Application
                'Set olApp = CreateObject("Outlook.Application")
                Set olMail = olApp.CreateItem(olMailItem)
                
                nombreTrabajador = Cells(indexFila, 1).Value
                emailTrabajador = Cells(indexFila, 2).Value
                subjectEmail = "Documentos faltantes o pendientes de la actualizacion 2020"
                
                'Especificar los datos de Outlook para el mensaje
                With olMail
                    .To = emailTrabajador
                    .Subject = subjectEmail & " - " & nombreTrabajador
                    .Attachments.Add archivoEnviar
                    .HTMLBody = cuerpoMensaje(indexFila, nombreTrabajador)
                    .Display
                    .Send
                End With
                
                'Si ya no utilizamos, liberamos
                Set olMail = Nothing
                Set olApp = Nothing
                
            End If
        
        Next indexFila
    
    End With

End Sub


'Funcion para armar el cuerpo del mensaje a enviar
'   @paramIndexFila -> Indice de la fila de cada registro
'   @paramNombreTrabajador -> Nombre del trabajador obtenido de valor de la celda
Function cuerpoMensaje(ByVal paramIndexFila As Integer, _
    ByVal paramNombreTrabajador) As String

    'Variables a utilizar
    Dim hoja As Worksheet
    Dim separador As String
    Dim firma As String
    Dim vinetaDocumentos As String
    Dim body As String
    Dim nameTrabajador As String
    
    'Inicializacion de variables
    Set hoja = ThisWorkbook.Worksheets("Datos")
    pathSeparator = Application.pathSeparator
    '-
    firma = ThisWorkbook.Path & "\archivos\firma.png"
    firma = Replace(firma, "\", pathSeparator)
    '-
    vinetaDocumentos = getVinetaDocumentosFO(paramIndexFila)

    'Contruimos el cuerpo del mensaje
    body = "<body> <font size=3>"
    body = body & "" & getSeccionDia() & " estimado(a) " & paramNombreTrabajador & "."
    body = body & "<br><br>Sirva la presente para indicarle que a la fecha aun tiene formatos pendientes: <br>"
    body = body & "<br>" & vinetaDocumentos
    body = body & "<br>Plazo maximo para regularizar 07 de mayo del 2020."
    body = body & "<br>Enviar al correo: <a href='filepersonal@softpang.com'> filepersonal@softpang.com</a>"
    body = body & "<br><br>Atentamente."
    body = body & "<br> <img src='" & firma & "'> "
    body = body & "</font> </body>"
    
    
    cuerpoMensaje = body
End Function

'Funcion para formar la vineta que contendra el listado de documentos faltantes de un trabajador
Function getVinetaDocumentosFO(ByVal paramIndexFila As Integer) As String
    
    Dim vineta As String

    Dim listaDocumentos() As String
    
    listaDocumentos = getListaDocumentosTrabajadorFO(paramIndexFila)

    vineta = "<ul>"
    For i = LBound(listaDocumentos) To UBound(listaDocumentos)
        If (listaDocumentos(i) <> Empty) Then
                vineta = vineta & "<li>" & listaDocumentos(i) & "</li>"
        End If
    Next i
    vineta = vineta & "</ul>"
    
    getVinetaDocumentosFO = vineta
End Function

'Funcion para obtener el listado de documentos faltantes de un trabajador
Function getListaDocumentosTrabajadorFO(ByVal paramIndexFila As Integer) As String()

    Dim listaDocumentos() As String

    Dim indexColumna As Integer
    Dim valorCelda As String
    Dim codigoDocumento As String
    Dim descripcionDocumento As String
    Dim estadoDocumento As String
    Dim contenidoVineta As String
    
    'Obtenemos la hoja de los registros que vamos a  trabajar
    Set hoja = ThisWorkbook.Worksheets("Datos")
    
    ReDim listaDocumentos(getIndexUltimaFilaListaDocumentos() - 1)

    'Recorremos toda las columnas, iniciando desde la columna 4 hasta la 10 (Desde la columna 4 comienzan los nombres de documentos)
    For indexColumna = 4 To 10
        
        valorCelda = hoja.Cells(paramIndexFila, indexColumna)
        
        'Vamos comprobando cada registro, desde las columna(4-10) asignada verificamos el valor de celda de cada columna
        If ((valorCelda Like "*FALTA*") Or (valorCelda Like "*OBSERVADO*")) Then
            
            codigoDocumento = hoja.Cells(1, indexColumna).Value
            descripcionDocumento = getDescripcionDocumento(codigoDocumento)
            estadoDocumento = hoja.Cells(paramIndexFila, indexColumna).Value
            
            'Vamos contruyendo el cuerpo de la viñeta
            contenidoVineta = "<b>" & descripcionDocumento & " </b> ---> (" & estadoDocumento & ")"
            listaDocumentos(indexColumna - 3) = contenidoVineta
        End If
        
    Next indexColumna
    
    getListaDocumentosTrabajadorFO = listaDocumentos
End Function

'Funcion para obtener la descripcion de los documentos segun su codigo
Function getDescripcionDocumento(ByVal codigoDocumento As String) As String

    'Declarar un objeto que sera utilizado como diccionario(Key, Value)
    Dim diccionario As Object
    
    'Variables a utilizar
    Dim hoja As Worksheet
    Dim indexFila As Integer
    Dim getCodigo As String
    Dim getDescripcion As String
    Dim existeKey As Boolean
          
    'Creamos el diccionario
    Set diccionario = CreateObject("Scripting.Dictionary")
    'Obtenemos la hoja que vamos a utilizar
    Set hoja = ThisWorkbook.Worksheets("BD_Documentos")
    
    For indexFila = 2 To hoja.Range("A" & Rows.Count).End(xlUp).Row
        getCodigo = Trim(hoja.Cells(indexFila, 1).Value)
        getDescripcion = (hoja.Cells(indexFila, 2).Value)
        
        diccionario.Add Key:=getCodigo, Item:=getDescripcion
    Next
    
    existeKey = diccionario.Exists(codigoDocumento)
    If existeKey Then
        getDescripcion = diccionario(codigoDocumento)
    Else
        getDescripcion = "CODIGO DE DOCUMENTO NO REGISTRADO"
    End If
    
    getDescripcionDocumento = "[" & codigoDocumento & "]> " & getDescripcion
End Function

'Funcion para obtener el ultimo indice de filas(Registros) que existe en la hoja(BD_Documentos)
Function getIndexUltimaFilaListaDocumentos() As Integer
    
    Dim hoja As Worksheet
    Set hoja = ThisWorkbook.Worksheets("BD_Documentos")
    
    indexUltimaFila = hoja.Cells(Rows.Count, 1).End(xlUp).Row

    getIndexUltimaFilaListaDocumentos = indexUltimaFila
End Function

'Funcion para obtener la seccion del dia(Mañana-Tarde-Noche) segun la hora
Function getSeccionDia() As String
    
    Dim hora As Integer
    Dim seccionDia As String
    Dim formatoHora As String
    
    formatoHora = Format(Time, "h:m:s")
    
    hora = Hour(formatoHora)
    If (hora < 12 And hora >= 0) Then
        seccionDia = "Buen dia"
    ElseIf (hora <= 18 And hora >= 12) Then
        seccionDia = "Buena tarde"
    Else
        seccionDia = "Buena noche"
    End If
    
    getSeccionDia = seccionDia
End Function

'Funcion que nos permite obtener la cantidad de documentos que le faltan o estan observados de un trabajador
' @inputListaDocumentos() -> Array con todos los documentos
' Return @cantidadDocumentos -> Retorna cantidad de documentos >=0
Function getCantidadDocumentosFOTrabajador(ByRef inputListaDocumentos() As String) As Integer
        
        Dim cantidadDocumentos As Integer
        cantidadDocumentos = 0
        
        For i = LBound(inputListaDocumentos) To UBound(inputListaDocumentos)
        If (inputListaDocumentos(i) <> Empty) Then
                cantidadDocumentos = cantidadDocumentos + 1
        End If
    Next i

    getCantidadDocumentosFOTrabajador = cantidadDocumentos
End Function
