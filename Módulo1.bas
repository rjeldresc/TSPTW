Attribute VB_Name = "Módulo1"
Option Explicit
Public puntero As Integer

'no se que hace esto
'@param codigoPedido
'@return
Public Function actualizarContadorFormularioEliminar(codigoPedido As String) As Integer
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim encontrado As Boolean
    Dim ArrayPedido() As String
    Dim fechaActual As String
    Dim codigoCentro As String
    Dim numeroPedido As String
    Dim listaBorrado As String
    i = 0
    encontrado = False
    ArrayPedido() = Split(codigoPedido, " ")
    codigoCentro = ArrayPedido(0)
    numeroPedido = ArrayPedido(1)
    fechaActual = Format(Date, "yyyy-MM")
    'DesprotegerHoja ("TablaNumeracionPedidos")
    Do While Sheets("TablaNumeracionPedidos").Cells(i + 1, 1).Value <> Empty And encontrado = False
        If Sheets("TablaNumeracionPedidos").Cells(i + 1, 1).Value = fechaActual And Sheets("TablaNumeracionPedidos").Cells(i + 1, 2).Value = codigoCentro Then
            If Sheets("TablaNumeracionPedidos").Cells(i + 1, 3).Value <> Empty Then
                listaBorrado = Sheets("TablaNumeracionPedidos").Cells(i + 1, 3).Value
                listaBorrado = listaBorrado & "//" & numeroPedido
                Sheets("TablaNumeracionPedidos").Cells(i + 1, 3).Value = listaBorrado
                encontrado = True
            Else
                listaBorrado = numeroPedido
                Sheets("TablaNumeracionPedidos").Cells(i + 1, 3).Value = listaBorrado
                encontrado = True
            End If
        Else
            i = i + 1
        End If
    Loop
    'ProtegerHoja ("TablaNumeracionPedidos")
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error Ingresando Datos"
End Function

'no se que hace esto
'@param
'@return
Public Function ListIsIn(lst As ListBox, zString As String) As Boolean
    On Error Resume Next
    Dim i As Integer
    Dim ListIsIn As Boolean
    For i = 0 To lst.ListCount
        If lst.List(i) = zString Then ListIsIn = True: lst.ListIndex = i: GoTo grr
    Next i
    ListIsIn = False
grr:
End Function

'Bloquea entradas no numericas
'@param KeyAscii corresponde al valor numerico de la tecla presionada
'@return True, si la tecla es numerica, False si es una tecla no permitida
Public Function ValidadEntradaSoloNumeros(KeyAscii As Integer) As Boolean
    '48 al 57 es valido, que serian numeros
    If (KeyAscii >= 48 And KeyAscii <= 57) = False Then
        ValidadEntradaSoloNumeros = True
    Else
        ValidadEntradaSoloNumeros = False
    End If
End Function

'no se que hace esto
'@param varArray
'@return
Public Function GetUpper(varArray As Variant) As Integer
    Dim Upper As Integer
    On Error Resume Next
    Upper = UBound(varArray)
    If Err.Number Then
        If Err.Number = 9 Then
            Upper = 0
        Else
            With Err
                MsgBox "Error:" & .Number & "-" & .Description
            End With
            Exit Function
        End If
    Else
        Upper = UBound(varArray) + 1
    End If
    On Error GoTo 0
    GetUpper = Upper
End Function

'Genera los numero de pedido, si hay pedidos eliminados siempre retorna el primero de los eliminados,
'por ejemplo: pedidos eliminados 7//3//5, retorna 3
'@param fechaIngreso fecha actual para buscar pedidos del mes en curso
'@param codigoCentro es el codigo del centro
'@return numero de pedido
Public Function ContadorNumeroPedido(fechaIngreso As String, codigoCentro As String) As Integer
    On Error GoTo ErrorHandler 'Manejador de errores
    Dim i As Integer
    Dim encontrado As Boolean
    Dim ArrayListaBorrado() As String 'array para la lista de los pedidos borrados
    i = 0
    encontrado = False
    Do While Sheets("TablaNumeracionPedidos").Cells(i + 1, 1).Value <> Empty And encontrado = False 'recorre la hoja TablaNumeracionPedidos mientras existan datos
        If Sheets("TablaNumeracionPedidos").Cells(i + 1, 1).Value = fechaIngreso And Sheets("TablaNumeracionPedidos").Cells(i + 1, 2).Value = codigoCentro Then 'Si la fecha de ingreso y codigo corresponden a lo requerido
            If Sheets("TablaNumeracionPedidos").Cells(i + 1, 3).Value <> Empty Then 'Significa que hay pedidos eliminados en la columna 3
                ArrayListaBorrado() = Split(Sheets("TablaNumeracionPedidos").Cells(i + 1, 3).Value, "//") 'Crea un array con pedidos eliminados
                SortArray ArrayListaBorrado() 'Ordena el array, queda de menor a mayor
                ContadorNumeroPedido = ArrayListaBorrado(0) 'obtiene el menor pedido
                encontrado = True 'Se encontro el pedido
            Else
                ContadorNumeroPedido = Sheets("TablaNumeracionPedidos").Cells(i + 1, 4).Value + 1 'En caso de no haber pedidos eliminados, retorna el contador de pedidos, por ejemplo, si hay 3 pedidos para un centro, el siguiente es el 3+1
                encontrado = True 'Se encontro el pedido
            End If
        Else
            i = i + 1 'avanza en filas por el excel
        End If
    Loop
    If encontrado = False Then 'en caso de no haber nada, es decir es el primer ingreso de un pedido del centro, es el numero 1
      ContadorNumeroPedido = 1
    End If
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con una descripcion
End Function

'no se que hace esto
'@param
'@return
Public Function actualizarContador(fechaIngreso As String, codigoCentro As String) As Integer
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim encontrado As Boolean
    Dim ArrayListaBorrado() As String
    i = 0
    encontrado = False
    'DesprotegerHoja ("TablaNumeracionPedidos")
    Do While Sheets("TablaNumeracionPedidos").Cells(i + 1, 1).Value <> Empty And encontrado = False
        If Sheets("TablaNumeracionPedidos").Cells(i + 1, 1).Value = fechaIngreso And Sheets("TablaNumeracionPedidos").Cells(i + 1, 2).Value = codigoCentro Then
            If Sheets("TablaNumeracionPedidos").Cells(i + 1, 3).Value <> Empty Then
                ArrayListaBorrado() = Split(Sheets("TablaNumeracionPedidos").Cells(i + 1, 3).Value, "//")
                SortArray ArrayListaBorrado()
                actualizarContador = ArrayListaBorrado(0)
                encontrado = True
                DeleteArrayItem ArrayListaBorrado(), 0
                ActualizarListaBorradodeHojaTablaNumeracionPedidos ArrayListaBorrado(), i
            Else
                actualizarContador = Sheets("TablaNumeracionPedidos").Cells(i + 1, 4).Value
                Sheets("TablaNumeracionPedidos").Cells(i + 1, 4).Value = Sheets("TablaNumeracionPedidos").Cells(i + 1, 4).Value + 1
                encontrado = True
            End If
        Else
            i = i + 1
        End If
    Loop
    If encontrado = False Then
        Sheets("TablaNumeracionPedidos").Cells(i + 1, 1) = fechaIngreso
        Sheets("TablaNumeracionPedidos").Cells(i + 1, 2) = codigoCentro
        Sheets("TablaNumeracionPedidos").Cells(i + 1, 3) = ""
        Sheets("TablaNumeracionPedidos").Cells(i + 1, 4) = 1
        actualizarContador = 1
    End If
    'ProtegerHoja ("TablaNumeracionPedidos")
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error"
End Function

'no se que hace esto
'@param
'@return
Public Sub ActualizarListaBorradodeHojaTablaNumeracionPedidos(ArrayListaBorrado() As String, fila As Integer)
    On Error GoTo ErrorHandler
    Dim listaBorrados As String
    Dim i As Integer
    'DesprotegerHoja ("TablaNumeracionPedidos")
    If UBound(ArrayListaBorrado()) = 0 Then
        Sheets("TablaNumeracionPedidos").Cells(fila + 1, 3) = ""
    ElseIf UBound(ArrayListaBorrado()) = 1 Then
        Sheets("TablaNumeracionPedidos").Cells(fila + 1, 3) = ArrayListaBorrado(0)
    Else
        listaBorrados = ArrayListaBorrado(0)
        For i = 1 To UBound(ArrayListaBorrado()) - 2
            listaBorrados = listaBorrados & "//" & ArrayListaBorrado(i)
        Next
        listaBorrados = listaBorrados & "//" & ArrayListaBorrado(UBound(ArrayListaBorrado()) - 1)
        Sheets("TablaNumeracionPedidos").Cells(fila + 1, 3) = listaBorrados
    End If
    'ProtegerHoja ("TablaNumeracionPedidos")
Exit Sub
ErrorHandler:
    MsgBox Err.Description, 16, "Error"
End Sub

'Ordena un array
'@param TheArray es el array a ordenar
'@return array ordenado de menor a mayor
Public Function SortArray(ByRef TheArray As Variant) As Variant
    Dim sorted As Boolean
    Dim x As Integer
    Dim temp As Variant
    sorted = False
    Do While Not sorted
        sorted = True
    For x = 0 To UBound(TheArray) - 1
        If TheArray(x) > TheArray(x + 1) Then
            temp = TheArray(x + 1)
            TheArray(x + 1) = TheArray(x)
            TheArray(x) = temp
            sorted = False
        End If
    Next x
    Loop
End Function

'Borra un elemento de un array
'@param arr es el array a procesar
'@param index indica la posicion
'@return
Public Function DeleteArrayItem(arr As Variant, index As Integer) As Variant
    On Error GoTo ErrorHandler
    Dim i As Long
    For i = index To UBound(arr) - 1
    arr(i) = arr(i + 1)
    Next
    ' VB will convert this to 0 or to an empty string.
    arr(UBound(arr)) = Empty
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error"
End Function

'Le agrega proteccion a la ficha indicada
'@param nombreHoja string con el nombre de la ficha a bloquear
'@return
Sub ProtegerHoja(nombreHoja As String)
    Sheets(nombreHoja).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="1111"
End Sub

'le quita la proteccion a la ficha indicada
'@param nombreHoja string con el nombre de la ficha a desbloquear
'@return
Sub DesprotegerHoja(nombreHoja As String)
    Sheets(nombreHoja).Unprotect Password:="1111"
End Sub

'no se que hace esto
'@param
'@return
Sub actualizarColumnaNumeroPedido(ContadorPedidos As Integer)
    Dim i As Integer
    For i = 1 To ContadorPedidos
        Cells(10 + i, 1) = i
    Next i
End Sub
Sub GuardarExcel()
    ActiveWorkbook.Save
End Sub

'Consulta una hoja del excel, y permite obtener una lista de resultados, similar a lo retornado
'por select [parametro] from nombreTabla
'@param nombreTabla corresponde al nombre de la hoja a ser consultada
'@return array con resultados
Public Function Select_from(nombreTabla As String) As String()
    Dim i As Integer                'Posicion en el arreglo
    Dim posicionLectura As Integer  'Posicion en el excel
    Dim resultado(1000) As String 'array donde se almacenan los resultados
    i = 0
    posicionLectura = 1 'Corresponde al valor desde que empiezan los pedidos
    Do While Sheets(nombreTabla).Cells(posicionLectura, 1).Value <> Empty 'Se lee toda la hoja, mientras existan datos
        resultado(i) = Sheets(nombreTabla).Cells(posicionLectura, 1).Value 'Lee el valor de la celda (i,j), y lo almacena en resultado
        i = i + 1 'posicion de array
        posicionLectura = posicionLectura + 1 'avanza a la siguiente posicion de lectura en el excel
    Loop
    Select_from = resultado 'retorna el resultado
End Function

'Cuenta la cantidad de elementos que hay en una hoja dentro del excel, similar a select count(*) from nombreHoja
'@param nombreHoja es el nombre de la hoja donde se contaran los elementos
'@param posicionFila desde donde comienza a contar elementos
'@param posicionColumna fija la columna por la que se va avanzado al contar pedidos
'@return Cantidad de elementos que tiene la hoja
Public Function select_count_from(nombreHoja As String, posicionFila As Integer, posicionColumna As Integer) As Integer
    On Error GoTo ErrorHandler
    Dim i As Integer
    i = 0
    Do Until Sheets(nombreHoja).Cells(posicionFila, posicionColumna).Value = Empty 'Se lee hasta fin de los datos
        i = i + 1 'incrementa en 1
        posicionFila = posicionFila + 1 'pasa a la siguiente posicion-Fila de lectura
    Loop
    select_count_from = i 'retorna el resultado
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Select codigo from [nombreTabla] where codigo=dato, obtiene el codigo del centro
'@param nombreHoja es el nombre de la hoja donde se contaran los elementos
'@param dato Nombre de lo buscado
'@param columnaBuscarDato valor de la columna donde se encuentra el dato que se esta buscando
'@param columnaRetorno valor de la columna del dato que tiene que ser retornado
'@return depende de quien llame la funcion, es un string
Public Function Select_from_where(nombreTabla As String, dato As String, columnaBuscarDato As Integer, columnaRetorno As Integer) As String
    Dim i As Integer                'Posicion en el arreglo
    Dim posicionLectura As Integer  'Posicion en el excel
    Dim resultado As String
    Dim encontrado As Boolean
    i = 0
    posicionLectura = 1 'Corresponde al valor desde que empiezan los pedidos
    encontrado = False
    Do Until encontrado Or Sheets(nombreTabla).Cells(posicionLectura, columnaBuscarDato).Value = Empty 'Se lee mientras existan datos
        If Sheets(nombreTabla).Cells(posicionLectura, columnaBuscarDato).Value = dato Then
            resultado = Sheets(nombreTabla).Cells(posicionLectura, columnaRetorno).Value
            encontrado = True
        End If
        i = i + 1
        posicionLectura = posicionLectura + 1
    Loop
    Select_from_where = resultado
End Function

'mmm no se que hace esto
'@param
'@return
Public Function ActualizarTablaNumeracionPedidos()
    Dim fechaActual As String
    Dim numeroFilas As Integer
    Dim puntero As Integer
    Dim encontrado As Boolean
    Dim columnaEliminar As String
    Dim i As Integer
    fechaActual = Format(Date, "yyyy-MM")
    numeroFilas = select_count_from("TablaNumeracionPedidos", 1, 1)
    puntero = 1
    encontrado = False
    Do Until encontrado
        If Sheets("TablaNumeracionPedidos").Cells(puntero, 1).Value = fechaActual Then
            encontrado = True
        Else
            puntero = puntero + 1
        End If
    Loop
    For i = puntero To numeroFilas
        Sheets("TablaNumeracionPedidos").Cells(i, 1) = ""
        Sheets("TablaNumeracionPedidos").Cells(i, 2) = ""
        Sheets("TablaNumeracionPedidos").Cells(i, 3) = ""
        Sheets("TablaNumeracionPedidos").Cells(i, 4) = ""
    Next
End Function

'Genera un mensaje en pantalla, de acuerdo a los parametros que se le dan
'@param titulo string que indica lo que se mostrara en el titulo de la ventana
'@param texto string que indica lo que se mostrara en el texto de la ventana
'@param icono string que indica cual sera el icono a mostrar
'@return
Public Function Mensaje(titulo As String, texto As String, icono As String)
    Dim numeroIcono As Integer
    Select Case icono
        Case "informacion"
            numeroIcono = 64
        Case "warning"
            numeroIcono = 48
        Case "pregunta"
            numeroIcono = 32
        Case "error"
            numeroIcono = 16
    End Select
    MsgBox texto, numeroIcono, titulo
End Function

'Limpia completamente la Ficha Stop
'@param
'@return
Public Function LimpiarFichaStop()
    On Error GoTo ErrorHandler
    Sheets("stop").Select
    Range("A1:A4").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Fase1").Select
Exit Function
ErrorHandler:
    Select Case Err.Number
      Case 1004 'La ficha esta vacia
        Resume Next
      Case Else 'Cualquier otro error no esperado
        Mensaje "Error", Err.Description, "error"
      End Select
End Function

'Genera la ruta del archivo, concatenado con el nombre de salida.txt
'@param
'@return la ruta completa del archivo salida.txt
Public Function RutaArchivo() As String
    RutaArchivo = Application.ThisWorkbook.Path & "\salida.txt"
End Function

'Optimiza los movimientos en el excel, elimina el parpadeo que aparece durante los movimientos
'@param mode si es on : indica que se activa la optimizacion, off: la desactiva
'@return
Public Function Optimizado(mode As String)
    Select Case mode
        Case "on"
            Application.ScreenUpdating = False 'Apagar el parpadeo de pantalla
            ActiveSheet.DisplayPageBreaks = False 'Apagar visualización de saltos de página
        Case "off"
            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            Application.CutCopyMode = False 'Borrar contenido de portapapeles
    End Select
    
End Function

