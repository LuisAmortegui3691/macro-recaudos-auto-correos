Attribute VB_Name = "Módulo1"
Sub macroCorreosAutonal()

    Dim tiempoInicio As Double
    Dim tiempoFinal As Double
    Dim duracion As Double

    Dim dia, mes, year
    Dim documentosEntrada As String, documentosSalida As String
    Dim plantilla
    Dim fecha
    Dim archiPagosGenalse As String
    Dim ultimaFilaPagosGenal
    Dim hoja As String
    Dim i As Long
    Dim valorCeldaPagosGenal
    Dim mesLetras As String
    Dim ultimaFilaPlantillaSoat As Long
    Dim valorTipo As String
    Dim placa As String, documento As String, nombreCliente As String, valor
    Dim ultimaFilaPlantiHoja1 As Long
    Dim plantilla_dos
    
    ' Registra el tiempo de inicio
    tiempoInicio = Timer
    
    fecha = ThisWorkbook.Sheets("main").Range("F2").Value
    hoja = ThisWorkbook.Sheets(1).Range("F3").Value
    
    dia = Split(fecha, "/")
    dia = dia(0)
    mes = Split(fecha, "/")
    mes = mes(1)
    year = Split(fecha, "/")
    year = year(2)
    
    documentosEntrada = ThisWorkbook.Sheets("main").Range("C2").Value
    documentosSalida = ThisWorkbook.Sheets("main").Range("C3").Value
    
    plantilla = documentosEntrada & "Plantilla\plantilla_correos_autonal_soat.xlsx"
    plantilla_dos = documentosEntrada & "Plantilla\plantilla_correos_autonal_polizas.xlsx"
    
    archiPagosGenalse = documentosEntrada & "Pagos Genalse\"
    archiPagosGenalse = Dir(archiPagosGenalse)
    
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=plantilla
    Application.DisplayAlerts = True
    
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=documentosEntrada & "Pagos Genalse\" & archiPagosGenalse
    Application.DisplayAlerts = True
    
    ' Seleccion de mes para convertirlo en letras
    Select Case mes
        Case "01"
            mesLetras = "ENERO"
        Case "02"
            mesLetras = "FEBRERO"
        Case "03"
            mesLetras = "MARZO"
        Case "04"
            mesLetras = "ABRIL"
        Case "05"
            mesLetras = "MAYO"
        Case "06"
            mesLetras = "JUNIO"
        Case "07"
            mesLetras = "JULIO"
        Case "08"
            mesLetras = "AGOSTO"
        Case "09"
            mesLetras = "SEPTIEMBRE"
        Case "10"
            mesLetras = "OCTUBRE"
        Case "11"
            mesLetras = "NOVIEMBRE"
        Case "12"
            mesLetras = "DICIEMBRE"
        Case Else
            MsgBox "Mes inválido. Ingrese un número de mes válido (01-12)."
    End Select
    

    Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets(1).Range("A4").Value = "PAGO " & dia & " DE " & mesLetras & year
    
    ultimaFilaPagosGenal = Workbooks(archiPagosGenalse).Sheets(hoja).Range("A" & Rows.Count).End(xlUp).Row
    
    
    ' Pegar valores de pagos generales a plantillla
    For i = 1 To 150
        valorCeldaPagosGenal = Workbooks(archiPagosGenalse).Sheets(hoja).Range("A" & i).Value
        
        If (valorCeldaPagosGenal <> "Area" And valorCeldaPagosGenal <> "") Then
            Workbooks(archiPagosGenalse).Sheets(hoja).Range("A" & i & ":J" & i).Copy
            Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("A" & i).PasteSpecial xlPasteValues
        End If
    Next i
    
    ' Definir la hoja de trabajo
    Set wsPlantilla = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales")
    ' Inicializar la variable de filas a eliminar
    Set filasAEliminar = Nothing
    
    ' Iterar sobre las filas para identificar las que deben eliminarse
    For i = 1 To wsPlantilla.Rows.Count
        If Trim(wsPlantilla.Range("A" & i).Value) = "" Then
            If filasAEliminar Is Nothing Then
                Set filasAEliminar = wsPlantilla.Rows(i)
            Else
                Set filasAEliminar = Union(filasAEliminar, wsPlantilla.Rows(i))
            End If
        End If
    Next i
    
    ' Eliminar todas las filas seleccionadas
    If Not filasAEliminar Is Nothing Then
        filasAEliminar.Delete
    End If
    
   ' Idntificar la celda con valor ORDENE DEVUELTAS para eliminar las filas hacia abajo incluyendola
    For i = 1 To 100
        valorCeldaPlantilla = Trim(Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("A" & i).Value)
        
        If valorCeldaPlantilla = "ORDENES DEVUELTAS" Then
            filaEliminar = i
            Exit For  ' Salir del bucle una vez encontrada la fila
        End If
    Next i
    
    ' Eliminar la fila y las subsiguientes
    If filaEliminar > 0 Then
        Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Rows(filaEliminar & ":" & wsPlantilla.Rows.Count).Delete
    End If
    
    Set wsPlantilla = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales")
    
    ' Encontrar la última fila con datos en la columna A
    ultimaFilaPlantilla = wsPlantilla.Cells(wsPlantilla.Rows.Count, "A").End(xlUp).Row
    
    
    ' Rango de datos para ordenar (asumiendo que los datos comienzan en la fila 1)
    Dim rangoDatos As Range
    Set rangoDatos = wsPlantilla.Range("A1:J" & ultimaFilaPlantilla)
    
    ' Ordenar por la columna "Tipo" (columna J)
    With wsPlantilla.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rangoDatos.Columns(10), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rangoDatos
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Validar ultima fila plantilla soat
    ultimaFilaPlantillaSoat = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("A" & Rows.Count).End(xlUp).Row
    
    ultimaFilaPlantiHoja1 = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Hacer recorrido y pegar valores
    For i = 1 To ultimaFilaPlantillaSoat
        valorTipo = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("J" & i).Value
        
        If valorTipo = "SOAT" Then
            placa = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("C" & i).Value
            documento = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("D" & i).Value
            nombreCliente = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("E" & i).Value
            valor = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("G" & i).Value
            valorTipo = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("J" & i).Value
        
            Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("Hoja1").Range("A" & i + 5).Value = placa
            Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("Hoja1").Range("B" & i + 5).Value = documento
            Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("Hoja1").Range("C" & i + 5).Value = nombreCliente
            Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("Hoja1").Range("D" & i + 5).Value = valor
            Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("Hoja1").Range("E" & i + 5).Value = valorTipo
        End If
    Next i
    
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=plantilla_dos
    Application.DisplayAlerts = True
    
    Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("A2:J" & ultimaFilaPlantillaSoat).Copy
    Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("datosTemporales").Range("A2").PasteSpecial xlPasteValues
    
    ' Validar ultima fila plantilla soat
    ultimaFilaPlantillaSoat = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("datosTemporales").Range("A" & Rows.Count).End(xlUp).Row
    
    ultimaFilaPlantiHoja1 = Workbooks("plantilla_correos_autonal_soat.xlsx").Sheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Hacer recorrido y pegar valores
    For i = 1 To ultimaFilaPlantillaSoat
        valorTipo = Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("datosTemporales").Range("J" & i).Value
        
        If valorTipo = "POLIZA" Then
            placa = Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("datosTemporales").Range("C" & i).Value
            documento = Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("datosTemporales").Range("D" & i).Value
            nombreCliente = Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("datosTemporales").Range("E" & i).Value
            valor = Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("datosTemporales").Range("G" & i).Value
            valorTipo = Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("datosTemporales").Range("J" & i).Value
        
            Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("Hoja1").Range("A" & i + 5).Value = placa
            Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("Hoja1").Range("B" & i + 5).Value = documento
            Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("Hoja1").Range("C" & i + 5).Value = nombreCliente
            Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("Hoja1").Range("D" & i + 5).Value = valor
            Workbooks("plantilla_correos_autonal_polizas.xlsx").Sheets("Hoja1").Range("E" & i + 5).Value = valorTipo
        End If
    Next i
    

    ' Registra el tiempo final
    tiempoFinal = Timer

    ' Calcula la duración en segundos
    duracion = tiempoFinal - tiempoInicio

    ' Muestra la duración en la ventana de inmediato
    Debug.Print "Tiempo de ejecución: " & duracion & " segundos"

    Debug.Print "Hola"
    
    MsgBox "Recuerda ingresar los valores de las pestañas de la hoja " & "NCS" & hoja
    

End Sub
