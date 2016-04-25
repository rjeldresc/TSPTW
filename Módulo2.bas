Attribute VB_Name = "Módulo4"
Option Explicit

'Algorithm 4: VND. Corresponde a la segunda parte del algoritmo.
'@param
'@return valor de la FO
Public Function VND() As Double
    On Error GoTo ErrorHandler
    Dim intentos  As Integer 'numero de intentos que lleva el do while
    Dim Movimiento2Opt As Boolean 'Usado para saber si se realizo el Movimiento2Opt
    Dim i As Integer 'Valor de la posicion-Fila, va incrementando por cada interacion que tiene el algoritmo,
    Dim parar As Boolean 'Indica si el Do While se debe detener
    Dim temp As Boolean
    Dim LineaTemp As String
    Const Linea1 = "LINEA 1" 'Constante, define el nombre de la linea
    Const Linea2 = "LINEA 2" 'Constante, define el nombre de la linea
    Const maxIntentos = 10 'Constante, numero maximo de intentos

    GuardarHoraUnixInicial 'Ingresa la hora actual a la ficha Valores
    i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
    
    If Sheets("Fase1").Cells(9, 5).Value = Sheets("Fase1").Cells(10, 5).Value Then 'paso 3: Si E9=E10
        Paso4 'Evalua FuncionObjetivo("Nueva_FO1") <= FuncionObjetivo("Antigua_FO1")
        If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then
            'Paso 6, la FO mejoro
            LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
            CopiarMejorSecuenciaATemporal 'copia la secuencia a la ficha "MejorSecuencia"
        Else
            Paso4 'Se hace en el Paso6, cuando vuelve al Paso4
        End If

        'GuardarFuncionesObjetivo 'Paso 7 , funcion por definir
        i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
        i = i + 1
        GuardarI i
        If Sheets("Fase1").Cells(11, 12).Value = Linea1 Then 'Paso 8, primer If, L11=LINEA 1
            If Sheets("Fase1").Cells(11, 5).Value = Sheets("Fase1").Cells(9, 5).Value Then 'E11 = E9
                'Paso9
Paso9a:          If Sheets("Fase1").Cells(i + 1, 12).Value = Linea1 Then 'L(i+1) = LINEA 1
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea1, i + 1) Then 'Evalua el calibre E(i+1) con el calibre del ultimo pedido de la LINEA X
                        i = i + 1
                        GoTo Paso9a 'un goto D:, regresa al Paso9a
                    Else
                        Paso10
                        'Paso 12
                        If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then
                            LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                            CopiarMejorSecuenciaATemporal 'Guarda la secuencia en la Ficha "MejorSecuencia"
                        Else
                            'Vuelve a ejecutar el paso 10
                            Paso10
                        End If
                    End If
                 End If
                 GuardarI i 'Guarda el valor de i en la Ficha Valores
            Else
                Paso10
                'Paso 12
                If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then
                    LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                    CopiarMejorSecuenciaATemporal 'Guarda la secuencia en la Ficha "MejorSecuencia"
                Else
                    Paso10
                End If
            End If
        End If

        GuardarHoraUnixFinal
        If TiempoStop Then
            GoTo CCCComboBreaker
        End If

        i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
        i = i + 1
        GuardarI i
        If Sheets("Fase1").Cells(11, 12).Value = Linea2 Then 'Paso 8, Segundo If, L11=LINEA 2
            If Sheets("Fase1").Cells(11, 5).Value = Sheets("Fase1").Cells(10, 5).Value Then 'E11 = E10
                'Paso9
Paso9b:         If Sheets("Fase1").Cells(i + 1, 12).Value = Linea2 Then 'L(i+1) = LINEA X
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea2, i + 1) Then 'Evalua el calibre E(i+1) con el calibre del ultimo pedido de la LINEA X
                        i = i + 1
                        GoTo Paso9b 'un goto D:, regresa al Paso9
                    Else
                        Paso10
                        'Paso 12
                        If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then
                            LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                            CopiarMejorSecuenciaATemporal 'Guarda la secuencia en la Ficha "MejorSecuencia"
                        Else
                            Paso10 'Vuelve a ejecutar el paso 10
                        End If
                    End If
                    GuardarI i 'Guarda el valor de i en la Ficha Valores
                End If
            Else
                Paso10
                'Paso 12
                If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then
                    LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                    CopiarMejorSecuenciaATemporal 'Guarda la secuencia en la Ficha "MejorSecuencia"
                Else
                    Paso10 'Vuelve a ejecutar el paso 10
                End If
            End If
        End If
    
    'GuardarFuncionesObjetivo 'Paso 13

    Else
        GuardarHoraUnixFinal
        If TiempoStop Then
            GoTo CCCComboBreaker
        End If
        'Else correspondiente al If del paso 3
        'Paso 14: Primer IF
        i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
        i = i + 1
        GuardarI i
        If Sheets("Fase1").Cells(11, 12).Value = Linea1 Then 'L11=LINEA 1
            If Sheets("Fase1").Cells(11, 5).Value = Sheets("Fase1").Cells(9, 5).Value Then 'E11=E9
Paso15aa:
                If Sheets("Fase1").Cells(i + 1, 12).Value = Linea1 Then
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea1, i + 1) Then
                        i = i + 1
                        GoTo Paso15aa 'un goto D:
                    Else
                        Paso16 ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
                    End If
                End If
Paso15ba:
                If Sheets("Fase1").Cells(i + 1, 12).Value = Linea2 Then
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea2, i + 1) Then
                        i = i + 1
                        GoTo Paso15ba 'un goto D:
                    Else
                        Paso16 ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
                    End If
                End If
                GuardarI i 'Guarda el valor de i en la Ficha Valores
            Else
                Paso16 ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
            End If
            Paso18 ValorI
        End If

        GuardarHoraUnixFinal
        If TiempoStop Then
            GoTo CCCComboBreaker
        End If
        'Paso 14: Segundo If
        i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
        i = i + 1
        GuardarI i
        If Sheets("Fase1").Cells(11, 12).Value = Linea2 Then 'L11= LINEA 1
            If Sheets("Fase1").Cells(11, 5).Value = Sheets("Fase1").Cells(10, 5).Value Then 'E11=E10
'Paso15
Paso15ab:
                If Sheets("Fase1").Cells(i + 1, 12).Value = Linea1 Then
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea1, i + 1) Then
                        i = i + 1
                        GoTo Paso15ab 'un goto D:
                    Else
                        Paso16 ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
                    End If
                End If
Paso15bb:
                If Sheets("Fase1").Cells(i + 1, 12).Value = Linea2 Then
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea2, i + 1) Then
                        i = i + 1
                        GoTo Paso15bb 'un goto D:
                    Else
                        Paso16 ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
                    End If
                End If
                GuardarI i 'Guarda el valor de i en la Ficha Valores
            Else
                Paso16 ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
            End If
            Paso18 ValorI
        End If

        GuardarHoraUnixFinal
        If TiempoStop Then
            GoTo CCCComboBreaker
        End If
        'GuardarFuncionesObjetivo 'Paso 19
        'Paso 20 Primer If
        i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
        i = i + 1
        GuardarI i
        If Sheets("Fase1").Cells(12, 12) = Linea1 Then 'L12 = LINEA 1
            If Sheets("Fase1").Cells(12, 5) = ultimoCalibre(Linea1, 12) Then 'E12=Ultimo Calibre
'Paso21
Paso21aa:
                If Sheets("Fase1").Cells(i + 1, 12).Value = Linea1 Then
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea1, i + 1) Then
                        i = i + 1
                        GoTo Paso21aa 'un goto D:
                    Else
                        Paso22
                    End If
                End If
Paso21ba:
                If Sheets("Fase1").Cells(i + 1, 12).Value = Linea2 Then 'L12= LINEA 2
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea2, i + 1) Then 'E12=Ultimo Calibre
                        i = i + 1
                        GoTo Paso21ba 'un goto D:
                    Else
                        Paso22
                    End If
                End If
                GuardarI i 'Guarda el valor de i en la Ficha Valores
                Paso24
            Else
                Paso22
                Paso24
            End If
        End If

        GuardarHoraUnixFinal
        If TiempoStop Then
            GoTo CCCComboBreaker
        End If
        'Paso 20 Segundo If
        i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
        i = i + 1
        GuardarI i
        If Sheets("Fase1").Cells(12, 12) = Linea1 Then 'L12= LINEA 2
            If Sheets("Fase1").Cells(12, 5) = ultimoCalibre(Linea2, 12) Then 'E12=Ultimo Calibre
'Paso21
Paso21ab:
                If Sheets("Fase1").Cells(i + 1, 12).Value = Linea1 Then
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea1, i + 1) Then
                        i = i + 1
                        GoTo Paso21ab 'un goto D:
                    Else
                        Paso22
                    End If
                End If
Paso21bb:
                If Sheets("Fase1").Cells(i + 1, 12).Value = Linea2 Then 'L12= LINEA 2
                    If Sheets("Fase1").Cells(i + 1, 5).Value = ultimoCalibre(Linea2, i + 1) Then 'E12=Ultimo Calibre
                        i = i + 1
                        GoTo Paso21bb 'un goto D:
                    Else
                        Paso22
                    End If
                End If
                GuardarI i 'Guarda el valor de i en la Ficha Valores
                Paso24
            Else
                Paso22
                Paso24
            End If
        End If
        'GuardarFuncionesObjetivo  'Paso 25
    End If
CCCComboBreaker:
    GuardarI i
    VND = FuncionObjetivoGlobal
	
Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 6 'En caso de desbordamiento
            Sheets("Valores").Cells(11, 2).Value = "11" 'En la ficha Valores, ingresa el valor por defecto 11
    End Select
    'Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

Public Function GuardarHoraUnixInicial()
    Sheets("valores").Cells(23, 2) = Date & " " & Time
End Function

Public Function GuardarHoraUnixFinal()
    Sheets("valores").Cells(24, 2) = Date & " " & Time
End Function

Public Function TiempoStop() As Boolean
    If Sheets("Valores").Cells(23, 4).Value = 1 Then
        TiempoStop = True
    Else
        TiempoStop = False
    End If
End Function

'Retorna el valor de i almacenado en la ficha Valores
'@param
'@return el valor de i
Public Function ValorI() As Integer
    On Error GoTo ErrorHandler 'Manejador de errores
    ValorI = Sheets("Valores").Cells(11, 2).Value 'Obtiene el valor de i que esta almacenado en la ficha Valores
Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 6 'En caso de desbordamiento
            Sheets("Valores").Cells(11, 2).Value = "11" 'En la ficha Valores, ingresa el valor por defecto 11
    End Select
    'Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Guarda el valor de i en la ficha Valores
'@param i valor que se usa para ir evaluando los calibres, en los pasos 9, 15 y 21
'@return
Public Function GuardarI(i As Integer)
    On Error GoTo ErrorHandler 'Manejador de errores
    Sheets("Valores").Cells(11, 2).Value = i
Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 6 'En caso de desbordamiento
            Sheets("Valores").Cells(11, 2).Value = "11" 'En la ficha Valores, ingresa el valor por defecto 11
    End Select
    'Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Debe encontrar un calibre que sea distinto a id_pedidoVNS, y con la fecha de entrega mas cercana
'y debe verificar que el movimiento sea de al menos 2 posiciones
'@param dato entrada a ser guardada
'@return
Public Function Local1Shift_Regla1(id_pedidoVNS As Integer)
    Dim contador  As Integer 'Va indicando la posicion-Fila en la ficha PedidosOrdenados
    Dim posicionPedidoMover As Integer 'posicion-Fila del pedido a mover, corresponde al pedido de la ficha PedidosOrdenados, pero segun posicion de la ficha Fase1
    Dim cantidadPedidosOrdenados As Integer
    Dim ultimoPedido As Integer
    cantidadPedidosOrdenados = select_count_from("PedidosOrdenados", 1, 2) 'Total de pedidos en la Ficha Fase1
    ultimoPedido = Sheets("Valores").Cells(2, 2).Value + 8 'Es la posicion-Fila del ultimo pedido de la ficha Fase1
    If cantidadPedidosOrdenados >= 10 Then 'Exige al menos 10 pedidos en la ficha PedidosOrdenados para hacer el movimiento
        contador = 1 'La lista empieza de la fila 1
        Do Until Sheets("Fase1").Cells(id_pedidoVNS, 5).Value <> _
                 Sheets("PedidosOrdenados").Cells(contador, 3).Value And _
                 movimientoAvanza2Posiciones(id_pedidoVNS, contador) 'Busca un calibre distinto y que el movimiento sea de al menos 2 posiciones
            contador = contador + 1 'Si no lo encuentra, avanza 1 posicion en la lista de pedidos de la ficha PedidosOrdenados
        Loop
        If id_pedidoVNS + 2 < ultimoPedido Then
            posicionPedidoMover = posicionRelativaDePedidosOrdenadosEnFase1(contador) 'Genera la posicion-Fila de la ficha Fase1 que se tiene que mover
            LeeryGuardarPedido id_pedidoVNS, posicionPedidoMover 'Realiza el movimiento
            borrarPedidoFichaPedidosOrdenados contador 'Borra el pedido que fue movido, desde la ficha PedidosOrdenados
        End If
    End If
End Function

'Permite saber si el pedido cumple la condicion
'@param id_pedidoVNS posicion-Fila del pedido destino
'@param contadorFichaPedidosOrdenados posicion-Fila del pedido a mover de la ficha PedidosOrdenados
'@return True, en caso de porder realizar el movimiento, False en caso de no poder realizar el movimiento
Public Function movimientoAvanza2Posiciones(id_pedidoVNS As Integer, posicionFichaPedidosOrdenados As Integer) As Boolean
    Dim posicion As Integer
    posicion = posicionRelativaDePedidosOrdenadosEnFase1(posicionFichaPedidosOrdenados)
    If Abs(posicion - id_pedidoVNS) >= 2 Then 'Es Abs, ya que pedidos anteriores pueden ser movidos, y dara como resultado un valor negativo, pero lo que interesa es la cantidad de movimientos
        movimientoAvanza2Posiciones = True
    Else
        movimientoAvanza2Posiciones = False
    End If
End Function

'Obtiene la posicion relativa de un pedido de la ficha Pedidos ordenados, respecto a la ficha Fase1
'Ejemplo: si un pedido AA1 esta en la posicion 5 de la ficha PedidosOrdenados, este pedido se busca en la ficha Fase1,  se
'busca su posicion
'@param posicionFichaPedidosOrdenados posicion-Fila del pedido destino
'@return
Public Function posicionRelativaDePedidosOrdenadosEnFase1(posicionFichaPedidosOrdenados As Integer) As Integer
    On Error GoTo ErrorHandler 'Manejador de errores
    Dim idPedidodsOrdenados As String
    Dim contador As Integer
    contador = 9 'Empieza buscando desde la posicion-Fila 9 de la ficha Fase1
    idPedidodsOrdenados = Sheets("PedidosOrdenados").Cells(posicionFichaPedidosOrdenados, 1).Value 'Obtiene el id del pedido en la ficha PedidosOrdenados
    Do Until idPedidodsOrdenados = Sheets("Fase1").Cells(contador, 3).Value 'Itera hasta que encuentre el pedido en la ficha Fase1
        contador = contador + 1 'Avanza una posicion en la ficha Fase1
    Loop
    posicionRelativaDePedidosOrdenadosEnFase1 = contador 'Retorna el valor
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Borra la fila de la ficha PedidosOrdenados donde esta el pedido que se movio
'@param posicion es la posicion-Fila del pedido que fue movido, es el pedido a borrar
'@return
Public Function borrarPedidoFichaPedidosOrdenados(posicion As Integer)
    On Error GoTo ErrorHandler 'Manejador de errores
    Sheets("PedidosOrdenados").Select 'Selecciona la ficha PedidosOrdenados
    Range("A1").Select 'Selecciona la primera celda
    Rows(posicion & ":" & posicion).Select 'Selecciona el rango donde esta el pedido a eliminar
    Selection.Delete Shift:=xlUp 'Elimina el pedido
    Sheets("Fase1").Select 'Selecciona la ficha Fase1
    Range("A1").Select 'Selecciona la primera celda
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Genera una lista con los pedidos ordenados, en la ficha PedidosOrdenados
'Pedidos ordenados por calibre, luego por fecha de entrega
'@param
'@return
Public Function OrdenarPedidosCalibreFechaEntrega()
    On Error GoTo ErrorHandler
    Dim pedidosVNS As Integer
    Sheets("Fase1").Select 'Selecciona la ficha Fase1
    pedidosVNS = select_count_from("Fase1", 9, 3) 'Total de pedidos en la Ficha Fase1
    Range("C9:J" & pedidosVNS + 8).Select 'Selecciona todos los pedidos
    Selection.Copy 'Los copia, es un CTRL+C
    Sheets("PedidosOrdenados").Select 'Selecciona la ficha PedidosOrdenados
    Range("A1").Select 'Selecciona la primera casilla de la ficha, es para que desde esa posicion se peguen los datos
    ActiveSheet.Paste 'Pegado normal, osea pega todo con formulas
    Range("A1:H" & pedidosVNS).Select
    Application.CutCopyMode = False 'Deshabilita el modo Cortar y Copiar.
    ActiveWorkbook.Worksheets("PedidosOrdenados").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PedidosOrdenados").Sort.SortFields.Add Key:=Range( _
        "C1:C" & pedidosVNS), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal 'Ordena los pedidos por calibre
    ActiveWorkbook.Worksheets("PedidosOrdenados").Sort.SortFields.Add Key:=Range( _
        "H1:H" & pedidosVNS), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal 'Ordena los pedidos por fecha de entrega
    With ActiveWorkbook.Worksheets("PedidosOrdenados").Sort 'Aplica el orden al rango seleccionado
        .SetRange Range("A1:H" & pedidosVNS)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Fase1").Select
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Retorna valores de FO, dependiendo del parametro de entrada
'@param valor parametro que indica que valor de las 4 FO es el que debe ser retornado
'@return Double, tiene el valor de la funcion objetivo
Public Function FuncionObjetivo(valor As String) As Double
error:
    On Error GoTo ErrorHandler 'manejador de errores
    Select Case valor
        Case "Nueva_FO1"
            FuncionObjetivo = Sheets("Valores").Cells(7, 1).Value 'Retorna la Nueva FO1
        Case "Antigua_FO1"
            FuncionObjetivo = Sheets("Valores").Cells(7, 2).Value 'Retorna la Antigua FO1
        Case "Nueva_FO2"
            FuncionObjetivo = Sheets("Valores").Cells(9, 1).Value 'Retorna la Nueva FO2
        Case "Antigua_FO2"
            FuncionObjetivo = Sheets("Valores").Cells(9, 2).Value 'Retorna la Antigua FO2
    End Select
Exit Function
ErrorHandler:
    'Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
    Select Case Err.Number
        Case 13 'En caso de un mal movimiento , da error numero 13, la FO entrega un resultado erroneo
            LimpiarFichaEnCasoDeError 'Borra todo de la ficha fase1
            CopiarTemporalAFase1 'Copia la secuencia de pedidos de la ficha MejorSecuencia a Fase1
            GenerarBatch
            ColumnaHolgura
            GenerarFormulasLinea1Linea2 11 'ingresa otra vez las formulas a la ficha fase1
            GoTo error
    End Select
End Function

'Realiza la Regla2, escoge un calibre distinto al de la posicion Ei, y escoge la menor fecha de entrega, luego verifica
'si ese movimiento es hacia atras o hacia adelante
'@param posicionDestino posicion-Fila del pedido destino, en la ficha Fase1
'@return True, si se puede realizar el movimiento, False en caso de no poder hacer el movimiento
Public Function Local2Opt_Regla2(posicionDestino As Integer) As Boolean
    Dim contador  As Integer 'Va indicando la posicion-Fila en la ficha PedidosOrdenados
    Dim posicionPedidoMover As Integer 'posicion-Fila del pedido a mover, corresponde al pedido de la ficha PedidosOrdenados, pero segun posicion de la ficha Fase1
    Dim cantidadPedidosOrdenados As Integer
    Dim ultimoPedido As Integer
    Dim auxiliar As Boolean
    auxiliar = False
    cantidadPedidosOrdenados = select_count_from("PedidosOrdenados", 1, 2) 'Total de pedidos en la Ficha PedidosOrdenados
    ultimoPedido = Sheets("Valores").Cells(2, 2).Value + 8 'Es la posicion-Fila del ultimo pedido de la ficha Fase1
    
    If cantidadPedidosOrdenados >= 10 Then 'Exige al menos 10 pedidos en la ficha PedidosOrdenados para hacer el movimiento
        contador = 1 'La lista de los PedidosOrdenados comienza de la fila 1
        
        'Busca un calibre distinto y que el movimiento sea de hacia atras o hacia adelante, en ciertos casos es posible que Do Until
        'no encuentre un pedido, en ese caso se cumple Value = Empty, es decir que el contador llego al final de la ficha
        Do Until (Sheets("Fase1").Cells(posicionDestino, 5).Value <> _
                 Sheets("PedidosOrdenados").Cells(contador, 3).Value And _
                 movimientoAvanza1Posicion(posicionDestino, contador)) Or _
                 Sheets("PedidosOrdenados").Cells(contador, 3).Value = Empty
            contador = contador + 1 'Si no lo encuentra, avanza 1 posicion en la lista de pedidos de la ficha PedidosOrdenados
        Loop

        'Una vez sale del Do Until, se necesita saber si realmente se encontro el pedido a mover,
        'por lo que se usa la misma condicion anterior
        If Sheets("Fase1").Cells(posicionDestino, 5).Value <> _
                 Sheets("PedidosOrdenados").Cells(contador, 3).Value And _
                 movimientoAvanza1Posicion(posicionDestino, contador) Then
                 
            If posicionDestino + 1 < ultimoPedido Then
                posicionPedidoMover = posicionRelativaDePedidosOrdenadosEnFase1(contador) 'Genera la posicion que se tiene que mover
                LeeryGuardarPedido posicionDestino, posicionPedidoMover 'Realiza el movimiento
                borrarPedidoFichaPedidosOrdenados contador 'Borra el pedido que fue movido, desde la ficha PedidosOrdenados
                auxiliar = True 'Se realizo el movimiento
            Else
                auxiliar = False 'No se hizo el movimiento
            End If
        Else
            auxiliar = False 'No se hizo el movimiento
        End If
    End If
    Local2Opt_Regla2 = auxiliar 'Retorna el valor
End Function

'Permite saber si el pedido cumple la condicion
'@param id_pedidoVNS posicion-Fila del pedido destino
'@param contadorFichaPedidosOrdenados posicion-Fila del pedido a mover de la ficha PedidosOrdenados
'@return
Public Function movimientoAvanza1Posicion(id_pedidoVNS As Integer, posicionFichaPedidosOrdenados As Integer) As Boolean
    Dim posicion As Integer
    posicion = posicionRelativaDePedidosOrdenadosEnFase1(posicionFichaPedidosOrdenados)
    If Abs(posicion - id_pedidoVNS) = 1 Then
        movimientoAvanza1Posicion = True
    Else
        movimientoAvanza1Posicion = False
    End If
End Function

'Realiza un movimiento al azar, es decir, si la posicion a determinar el la i , los posibles movimientos
'se realizan desde i+1 hasta el ultimo pedido de la ficha Fase1
'@param posicionDestino posicion-Fila del pedido destino en la ficha Fase1
'@return
Public Function Regla3(posicionDestino As Integer)
    Dim ultimoPedido As Integer
    Dim PedidoMover As Integer
    ultimoPedido = Sheets("Valores").Cells(2, 2).Value + 8 'Es la posicion-Fila del ultimo pedido de la ficha Fase1
    If posicionDestino + 1 < ultimoPedido Then
        PedidoMover = Aleatorio(posicionDestino + 1, ultimoPedido) 'Obtiene un numero random, entra la posicion i+1 a ultimoPedido
        LeeryGuardarPedido posicionDestino, PedidoMover 'Realiza el movimiento
    End If
End Function

'Busca el ultimo pedido que entro en una determinada Linea
'@param linea Linea de la que se debe buscar el calibre
'@param pedido posicion-Fila desde la que hay que buscar el pedido
'@return ultimo calibre que entro en la Linea
Public Function ultimoCalibre(linea As String, pedido As Integer) As Integer
    Dim i As Integer
    i = 1
    Do Until Sheets("Fase1").Cells(pedido - i, 12).Value = linea 'evalua la columna L
        i = i + 1
    Loop
    ultimoCalibre = Sheets("Fase1").Cells(pedido - i, 5).Value 'retorna el calibre
End Function

'Ejecuta el Paso4
'@param
'@return
Public Function Paso4()
    Dim intentos  As Integer 'numero de intentos que lleva el do while
    'Dim maxIntentos As Integer 'numero maximo de intentos para el d while
    Dim parar As Boolean 'Indica si el Do While se debe detener
    Const maxIntentos = 10 'Constante, numero maximo de intentos
    intentos = 0
    parar = False
    'Do While Itera hasta el maximo de maxIntentos, o hasta que encuentre una FO menor, lo que ocurra primero
    Do While intentos <= maxIntentos And Not parar
        Local1Shift_Regla1 10 'Se tiene que mover un pedido a la posicion-Fila 10
        If FuncionObjetivo("Nueva_FO1") <= FuncionObjetivo("Antigua_FO1") Then 'compara la nueva FO con la antigua
            LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
            CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
            parar = True 'Detiene la iteracion del Do While
        Else
            CopiarTemporalAFase1 'Borra la secuencia, y copia la mejor a Fase1
            GeneraTiempoPreparacion 11
            'Turbo 11
            intentos = intentos + 1 'Incrementa el valor de los intentos
        End If
    Loop
End Function

'Ejecuta el Paso10, tambian hace un Local2Opt_Regla2 , y dependiendo si lo pudo hacer, realiza la Regla3
'@param
'@return
Public Function Paso10()
    Dim intentos As Integer
    Dim parar As Boolean
    Dim Movimiento2Opt As Boolean
    Dim i As Integer
    Const maxIntentos = 10 'Constante, numero maximo de intentos
    intentos = 0
    parar = False
    Do While intentos <= maxIntentos And Not parar 'itera un maximo de 10 veces
        Movimiento2Opt = Local2Opt_Regla2(12) 'posicion-Fila 12, pero es el pedido 4
        If Movimiento2Opt Then 'Si se puede realizar Local2Opt_Regla2
            If FuncionObjetivo("Nueva_FO1") <= FuncionObjetivo("Antigua_FO1") Then 'Evalua FO1<= FO1
                LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
                parar = True 'Detiene la iteracion del Do While
            Else 'La FO no mejoro
                CopiarTemporalAFase1 'Vuelve a la secuencia anterior
                GeneraTiempoPreparacion 11
                'Turbo 11
                intentos = intentos + 1 'Incrementa en 1 intento
            End If
        Else 'En caso de no poder hacer la Regla2, se hace la Regla3
            i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
            Regla3 i
            If FuncionObjetivo("Nueva_FO1") <= FuncionObjetivo("Antigua_FO1") Then
                LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
                parar = True 'Detiene la iteracion del Do While
            Else
                CopiarTemporalAFase1 'Vuelve a la secuencia anterior
                GeneraTiempoPreparacion 11
                'Turbo 11
                intentos = intentos + 1 'Incrementa en 1 intento
            End If
        End If
    Loop
End Function

'Ejecuta el Paso16, con un Local1Shift_Regla1
'@param
'@return
Public Function Paso16(i As Integer)
    Dim intentos As Integer
    Dim parar As Boolean
    intentos = 0
    parar = False
    Const maxIntentos = 10 'Constante, numero maximo de intentos
    Do While intentos <= maxIntentos And Not parar
        Local1Shift_Regla1 i 'Mueve un pedido para que el calibre i sea distinto al calibre i-1
        If FuncionObjetivo("Nueva_FO1") <= FuncionObjetivo("Antigua_FO1") Then
            LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
            CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
            GuardarExcel
            parar = True 'Detiene la iteracion del Do While
        Else
            CopiarTemporalAFase1
            GeneraTiempoPreparacion 11
            'Turbo 11
            intentos = intentos + 1
        End If
    Loop
End Function

'Borra el contenido de la ficha MejorSecuencia, usado antes de copiar la mejor secuencia a la ficha MejorSecuencia
'@param
'@return
Public Function LimpiarFichaMejorSecuencia()
    On Error GoTo ErrorHandler
    Dim TotalPedidos As Integer
    Sheets("MejorSecuencia").Select
    TotalPedidos = select_count_from("MejorSecuencia", 1, 1) 'Obtiene el total de pedidos, desde la ficha Valores'Borra el contenido de la ficha MejorSecuencia
    Range("A1:H" & TotalPedidos).Select
    Application.CutCopyMode = False
    Selection.ClearContents
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

'Limpia por completo la ficha de PedidosOrdenados
'@param
'@return
Public Function LimpiarFichaPedidosOrdenados()
    On Error GoTo ErrorHandler
    Dim TotalPedidos As Integer
    Sheets("PedidosOrdenados").Select
    TotalPedidos = select_count_from("PedidosOrdenados", 1, 1)
    Range("A1:H" & TotalPedidos).Select
    Selection.ClearContents
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

'Limpia por completo la ficha Fase1
'@param
'@return
Public Function LimpiarFichaEnCasoDeError()
    On Error GoTo ErrorHandler
    Dim TotalPedidos As Integer
    Sheets("Fase1").Select
    TotalPedidos = Sheets("Valores").Cells(2, 2).Value + 20
    Range("C9:AQ" & TotalPedidos).Select
    Selection.ClearContents
Exit Function
ErrorHandler:
    Select Case Err.Number
      Case 1004 'La ficha esta vacia
        Resume Next
      Case Else 'Cualquier otro error no esperado
        Mensaje "Error", Err.Description, "error"
      End Select
End Function

'Limpia por completo la ficha Fase1
'@param
'@return
Public Function LimpiarFichaFase1()
    On Error GoTo ErrorHandler
    Dim TotalPedidos As Integer
    Sheets("Fase1").Select
    TotalPedidos = select_count_from("Fase1", 9, 3) + 8
    Range("C9:AQ" & TotalPedidos).Select
    Selection.ClearContents
Exit Function
ErrorHandler:
    Select Case Err.Number
      Case 1004 'La ficha esta vacia
        Resume Next
      Case Else 'Cualquier otro error no esperado
        Mensaje "Error", Err.Description, "error"
      End Select
End Function

'Permite Borrar una ficha
'@param nombreFicha Corresponde a la ficha que se quiere borrar
'@param rango1 es la casilla desde donde comienza la seleccion de las casillas de la ficha
'@param rango2 es la casilla donde termina la seleccion, pero solo es la columna
'@param fila es la posicion-Fila desde donde empieza a contar los elementos
'@param columna es la posicion-Columna desde donde empieza a contar los elementos
'@param puntero para el caso de que los elementos no empiezen en la posicion 1,1
'@return
Public Function LimpiarFicha(nombreFicha As String, rango1 As String, rango2 As String, fila As Integer, columna As Integer, puntero As Integer)
    Dim TotalPedidos As Integer
    Sheets(nombreFicha).Select
    TotalPedidos = select_count_from(nombreFicha, fila, columna) + puntero
    Range(rango1 & ":" & rango2 & TotalPedidos).Select
    Selection.ClearContents
    Sheets("Fase1").Select
End Function

'Genera las formular BATCH de la ficha Fase1
'@param
'@return
Public Function GenerarBatch()
    On Error GoTo ErrorHandler
    Dim TotalPedidos As Integer
    TotalPedidos = select_count_from("Fase1", 9, 3) + 8
    Sheets("Fase1").Cells(9, 7).FormulaLocal = "=F9/BUSCARV(D9;'Tiempos Procesos y Preparacion'!$G$5:$J$116;2;FALSO)"
    Range("G9").Select
    Selection.AutoFill Destination:=Range("G9:G" & TotalPedidos)
    Range("G9:G" & TotalPedidos).Select
    Range("C8").Select
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error"
End Function

'Ejecuta el Paso18, Local1Shift_Regla1
'@param
'@return
Public Function Paso18(i As Integer)
    Dim intentos As Integer
    Dim parar As Boolean
    Const maxIntentos = 10 'Constante, numero maximo de intentos
    If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then
        LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
        CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
    Else
        intentos = 0
        parar = False
        Do While intentos <= maxIntentos And Not parar
            Local1Shift_Regla1 i 'Funcion por definir
            If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then
                LimpiarFichaMejorSecuencia
                CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
                parar = True 'Detiene la iteracion del Do While
            Else
                CopiarTemporalAFase1
                GeneraTiempoPreparacion 11
                'Turbo 11
                intentos = intentos + 1
            End If
        Loop
    End If
End Function

'Ejecuta el Paso22, tambian hace un Local2Opt_Regla2 , y dependiendo si lo pudo hacer, realiza la Regla3
'@param
'@return
Public Function Paso22()
    Dim intentos As Integer
    Dim parar As Boolean
    Dim Movimiento2Opt As Boolean
    Dim i As Integer
    Const maxIntentos = 10 'Constante, numero maximo de intentos
    intentos = 0
    parar = False
    i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
    Do While intentos <= maxIntentos And Not parar 'Paso 22
        Movimiento2Opt = Local2Opt_Regla2(i)  'Posicion destino
        If Movimiento2Opt Then
            If FuncionObjetivo("Nueva_FO1") <= FuncionObjetivo("Antigua_FO1") Then
                LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
                parar = True 'Detiene la iteracion del Do While
            Else
                CopiarTemporalAFase1
                GeneraTiempoPreparacion 11
                'Turbo 11
                intentos = intentos + 1
            End If
        Else
            i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
            Regla3 i
            If FuncionObjetivo("Nueva_FO1") <= FuncionObjetivo("Antigua_FO1") Then
                LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
                parar = True 'Detiene la iteracion del Do While
            Else
                CopiarTemporalAFase1
                GeneraTiempoPreparacion 11
                'Turbo 11
                intentos = intentos + 1
            End If
        End If
    Loop
End Function

'Ejecuta el Paso24, compara Nueva_FO2 con Antigua_FO2, y ejecuta Local2Opt_Regla2
'@param
'@return
Public Function Paso24()
    Dim intentos As Integer
    Dim parar As Integer
    Dim i As Integer
    Const maxIntentos = 10 'Constante, numero maximo de intentos
    
    If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then 'paso 12
        LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
        CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
    Else
        intentos = 0
        parar = False
        Do While intentos <= maxIntentos And Not parar
            i = ValorI 'Obtiene el valor de i que esta almacenado en la ficha Valores
            Local2Opt_Regla2 i 'Funcion por definir, se le pasa una posicion
            If FuncionObjetivo("Nueva_FO2") <= FuncionObjetivo("Antigua_FO2") Then
                LimpiarFichaMejorSecuencia 'Borra todos los datos antiguos de la ficha MejorSecuencia
                CopiarMejorSecuenciaATemporal 'Copia los pedidos de la Ficha Fase1 a MejorSecuencia
                parar = True 'Detiene la iteracion del Do While
            Else
                CopiarTemporalAFase1 'Copia la mejor secuancia a la ficha Fase1
                GeneraTiempoPreparacion 11 'Actualiza los tiempos de preparacion de ambas lineas
                'Turbo 11
                intentos = intentos + 1 'Incrementa en +1 los intentos
            End If
        Loop
    End If
End Function
