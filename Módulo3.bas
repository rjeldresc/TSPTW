Attribute VB_Name = "Módulo3"
Option Explicit

'Algorithm 1: Two Phase Heuristic
'@param iterMax: indica numero maximo de iteraciones
'@return Mejor Funcion Objetivo
Public Function Algorithm1(iterMax As Integer) 'As Double
    On Error GoTo ErrorHandler
    Dim x As Double
    Dim x1 As Double
    Dim iter As Integer
    iter = 0
    Do While iter < iterMax ' And x_infactible 'La condicion original era iter < iterMax
        x1 = Sheets("Valores").Cells(7, 2).Value 'Es la mejor solucion hasta este momento
        x = BuildFeasibleSolution 'see algorithm 2 'Algorithm 2: VNS - Constructive phase
        x = GVNS(8) 'GVNS
        x1 = better(x, x1) 'x1 es en realidad x* dentro del algoritmo que esta en pdf
        iter = iter + 1 'Incrementa en 1 iteracion
    Loop
    'Algorithm1 = x1 'Retorna el valor x*
Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 13 'En caso de un mal movimiento , da error numero 13, la FO entrega un resultado erroneo
            LimpiarFichaEnCasoDeError 'Borra todo de la ficha fase1
            CopiarTemporalAFase1 'Copia la secuencia de pedidos de la ficha MejorSecuencia a Fase1
            GenerarBatch
            ColumnaHolgura
            GenerarFormulasLinea1Linea2 11 'ingresa otra vez las formulas a la ficha fase1
            Resume Next 'Sigue con la siguiente linea del codigo
    End Select
End Function

'Algorithm 2: VNS - Constructive phase , Genera una solucion factible para un determinado numero de pedidos
'@param
'@return Valor de la Funcion Objetivo
Public Function BuildFeasibleSolution() As Double
    On Error GoTo ErrorHandler
    Dim level As Integer
    Dim levelMax As Integer
    Dim x As Double
    Dim x1 As Double
    Dim maxIteraciones As Integer
    Dim error As Boolean
    Dim Criterio As Integer 'Define el criterio que se usara para detener a BuildFeasibleSolution
    x = 0
    x1 = 0
    level = 0
    levelMax = CInt(Sheets("Fase1").Cells(2, 7).Value) 'Lee levelMax desde la ficha Fase1
    maxIteraciones = 0
    'DesprotegerHoja "Fase1"
    'x = RandomSolution 'Ordena los pedidos por fecha de entrega, desde el mas cercano al mas lejano
    Criterio = Sheets("stop").Cells(1, 1).Value 'Lee el criterio a usar para detener el While
    Select Case Criterio 'De acuerdo al criterio, selecciona el codigo a ejecutar
        Case 1 'Para que se detenga despues de n-Iteraciones
            Do Until x_factible Or Not IteracionesMax(maxIteraciones) 'Se ejecuta hasta que la solucion es factible
                level = 1
                x = Local1Shift 'Efectua 4 movimientos
                Do While x_infactible And level < levelMax And IteracionesMax(maxIteraciones)  'Se ejecuta mientras la solucion sea infactible y level sea menor a levelMax
                    x1 = Perturbation(level) 'Realiza movimientos random Level-veces
                    x1 = Local1Shift 'Efectua 4 movimientos
                    x = better(x, x1) 'Compara la solución anterior con la actual, y retorna la mejor solución
                    If x = x1 Then
                        level = 1
                    Else
                        level = level + 1
                    End If
                    maxIteraciones = maxIteraciones + 1
                    copiadoDePedidos 'Verifica si la FO mejoro
                Loop
            Loop
        Case 3
            Do Until x_factible 'Se ejecuta hasta que la solucion es factible
                level = 1
                x = Local1Shift 'Efectua 4 movimientos
                Do While x_infactible And level < levelMax 'Se ejecuta mientras la solucion sea infactible y level sea menor a levelMax
                    x1 = Perturbation(level) 'Realiza movimientos random Level-veces
                    x1 = Local1Shift 'Efectua 4 movimientos
                    x = better(x, x1) 'Compara la solución anterior con la actual, y retorna la mejor solución
                    If x = x1 Then
                        level = 1
                    Else
                        level = level + 1
                    End If
                    copiadoDePedidos 'Verifica si la FO mejoro
                Loop
            Loop
    End Select
    BuildFeasibleSolution = x 'Termina la iteracion del algoritmo, y retorna el resultado a Algorithm 1: Two Phase Heuristic
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Algorithm 3: GVNS
'@param levelMax numero maximo de iteraciones que hace el Do While
'@return
Public Function GVNS(levelMax As Integer) As Double
'ACA ESCRIBIR NUEVA FUNCION
End Function


'Algorithm 3: GVNS_2 Funcion Antigua, se reemplazo por GVNS
'@param levelMax numero maximo de iteraciones que hace el Do While
'@return
Public Function GVNS_2(levelMax As Integer) As Double
    On Error GoTo ErrorHandler
    Dim level As Integer
    Dim x As Double
    Dim x1 As Double
    'If x_infactible Then
    level = 1
    x = VND
    'GuardarHoraUnixInicialGVNS
    Do While level <= levelMax 'And x_infactible And TiempoStopGVNS
        x1 = Perturbation(level)
        x1 = VND
        x = better(x, x1)
        If x = x1 Then
            level = 1
        Else
            level = level + 1
        End If
        'GuardarHoraUnixFinalGVNS
    Loop
    '    GVNS = x
    'Else
    '    GVNS = FuncionObjetivo
    'End If
    GVNS = FuncionObjetivoGlobal
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

Public Function GuardarHoraUnixInicialGVNS()
    Sheets("valores").Cells(27, 2) = Date & " " & Time
End Function

Public Function GuardarHoraUnixFinalGVNS()
    Sheets("valores").Cells(28, 2) = Date & " " & Time
End Function

Public Function TiempoStopGVNS() As Boolean
    If Sheets("Valores").Cells(27, 4).Value = 1 Then
        TiempoStopGVNS = False
    Else
        TiempoStopGVNS = True
    End If
End Function

'Detiene a BuildFeasibleSolution al llegar a un limite maximo de iteraciones
'@param iteraciones cantidad de iteraciones actual que lleva el algoritmo
'@return True cuando aún no alcanza a iteracionesMaximas, False cuando alcanzo el maximo de iteraciones
Public Function IteracionesMax(iteraciones As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim iteracionesMaximas As Integer
    iteracionesMaximas = Sheets("stop").Cells(2, 1).Value 'Lee las iteraciones maximas desde la Ficha "stop"
    If iteraciones <= iteracionesMaximas Then
        IteracionesMax = True
    ElseIf iteraciones > iteracionesMaximas Then
        IteracionesMax = False
    End If
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Detiene a BuildFeasibleSolution al llegar a un tiempo maximo.
'La funcion timer tiene la limitacion de entregar la hora en segundos despues de la media noche del dia en curso, asi que
'existe un problema al ejecutar el programa unos cuantos minutos antes de la media noche.
'@param
'@return True cuando aún no alcanza el tiempo Maximo, False cuando alcanzo el tiempo maximo
Public Function tiempoMaximo() As Boolean
    On Error GoTo ErrorHandler
    Dim TiempoStop As Single
    Dim tiempoActual As Single
    TiempoStop = Sheets("stop").Cells(2, 1).Value * 60 'Pasa el tiempo a segundos
    'Select Case Sheets("stop").Cells(3, 1).Value
        'Case "Minutos"
            'tiempoStop = tiempoStop * 60 'Tiempo lo pasa a segundos
        
        'Case "Horas"
            'tiempoStop = tiempoStop * 3600 'Tiempo lo pasa a segundos
    'End Select
    tiempoActual = Timer - Sheets("stop").Cells(4, 1).Value
    If tiempoActual < TiempoStop Then
        tiempoMaximo = True
    ElseIf tiempoActual < TiempoStop Then
        tiempoMaximo = False
    End If
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Permite obtener la mejor Funcion Objetivo
'@param a FO de una iteracion anterior
'@param b FO de una iteracion actual
'@return la mayor FO
Public Function better(a As Double, b As Double) As Double
    On Error GoTo ErrorHandler
    If a > b Then
        better = b
    ElseIf b > a Then
        better = a
    ElseIf a = b Then
        better = a
    End If
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Se obtiene la posicion-Fila en la ficha Pedidos para los pedidos del mes en curso
'@param
'@return puntero indica la posicion-Fila desde donde comienzan los pedidos para el mes en curso, en la ficha Pedidos
Public Function punteroActualizado() As Integer
    Dim encontrado As Boolean
    Dim fechaActual As String
    Dim fechaRecepcion As String
    fechaActual = Format(Date, "yyyy-MM") 'Obtiene la fecha actual , formato año-mes, por ejemplo 2012-08
    encontrado = False 'Usado para encontrar los pedidos del mes en curso
    puntero = 11 'Indica la posicion-Fila de la Hoja Pedidos desde donde empiezan los pedidos
    
    'Do Until necesario para leer solo los pedidos del mes en curso
    Do Until encontrado 'Itera hasta que encontrado = true
        fechaRecepcion = Format(Cells(puntero, 4).Value, "yyyy-MM") 'Lectura de la fecha de recepcion en la ficha Pedidos, formato año-mes
        If fechaRecepcion <> fechaActual And Sheets("Pedidos").Cells(puntero, 1).Value <> "" Then 'Si la fecha de recepcion es distinta a la fecha actual, sigue iterando hasta que encuentre las fechas actuales
            puntero = puntero + 1 'Incrementa posicion
        Else
            encontrado = True 'Encontro los pedidos actuales
        End If
    Loop
    punteroActualizado = puntero
End Function

'Copia los datos (Pedido, tipo, calibre, fecha de ingreso, fecha de entrega, etc) de la ficha pedidos hacia la ficha VNS
'@param puntero indica la posicion-Fila desde donde comienzan los pedidos en la ficha Pedidos
'@return
Public Function CopiarDatosdeFichaPedidosaVNS(puntero As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim CalibreTemp As String
    Dim Calibre As Integer
    Dim ArrayCalibre() As String 'Usado para separar el AD del numero del calibre
    'Copiando datos de columna Pedido
    j = 0 'Incrementa la posicion-Fila de la ficha Fase1
    For i = puntero - 11 To Sheets("Valores").Cells(2, 2).Value - 1 'Itera n-veces segun la cantidad n de pedidos; i indica posicion-Fila para la ficha Pedidos
        Sheets("Fase1").Cells(9 + j, 3) = Sheets("Pedidos").Cells(i + 11, 2) 'Pedido
        Sheets("Fase1").Cells(9 + j, 4) = Sheets("Pedidos").Cells(i + 11, 3) 'tipo
        CalibreTemp = Sheets("Pedidos").Cells(i + 11, 8) 'Lee celda de calibre, tiene el AD, asi que lo lee como string
        ArrayCalibre() = Split(CalibreTemp, "AD") 'Separa CalibreTemp, y retorna la "primera parte" que contiene el calibre, como dato numerico
        Calibre = CInt(ArrayCalibre(0)) 'calibre de string lo transforma a dato Integer
        Sheets("Fase1").Cells(9 + j, 5) = Calibre 'Calibre
        Sheets("Fase1").Cells(9 + j, 6) = Sheets("Pedidos").Cells(i + 11, 7) 'kgs
        Sheets("Fase1").Cells(9 + j, 7).FormulaLocal = "=F" & 9 + j & "/BUSCARV(D" & 9 + j & ";'Tiempos Procesos y Preparacion'!$G$5:$J$116;2;FALSO)" 'BATCH
        Sheets("Fase1").Cells(9 + j, 8) = Sheets("Pedidos").Cells(i + 11, 12) 'Envase
        Sheets("Fase1").Cells(9 + j, 9) = Sheets("Pedidos").Cells(i + 11, 4) 'Fecha recepcion
        Sheets("Fase1").Cells(9 + j, 10) = Sheets("Pedidos").Cells(i + 11, 13) 'Fecha entrega
        'Sheets("Fase1").Cells(9 + j, 11).FormulaLocal = "=J" & 9 + j & "-$G$1" 'Holgura
        j = j + 1 'Avanza 1 fila en la ficha Fase1
    Next
End Function

'Ingresa las formulas correspondientes a los tiempos de comienzo, preparacion, finalizacion para las LINEA 1 y LINEA 2
'@param puntero indica la posicion-Fila desde donde comienzan los pedidos en la ficha Pedidos
'@return
Public Function GenerarFormulasLinea1Linea2(puntero As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim ContadorPedidos As Integer
    '"Primera Fila"
    Sheets("Fase1").Cells(9, 12).FormulaLocal = "LINEA 1" 'LINEA PROD
    Sheets("Fase1").Cells(9, 13).FormulaLocal = "=$B$6" 'Linea 1: T. Comienzo Maq 1
    Sheets("Fase1").Cells(9, 14).FormulaLocal = "=BUSCARV(D9;'Tiempos Procesos y Preparacion'!$G$5:$J$116;3;FALSO)*G9" 'Linea 1: T. Proceso Maq 1
    Sheets("Fase1").Cells(9, 15).FormulaLocal = "=N9" 'Linea 1: T. Finalización Maq 1
    Sheets("Fase1").Cells(9, 16).FormulaLocal = "=0" 'Linea 1: Tiempo Preparación Maq 2
    Sheets("Fase1").Cells(9, 17).FormulaLocal = "=P9+O9" 'Linea 1: T. Comienzo Maq 2
    Sheets("Fase1").Cells(9, 18).FormulaLocal = "=BUSCARV(D9;'Tiempos Procesos y Preparacion'!$G$5:$J$116;4;FALSO)*G9/3"
    Sheets("Fase1").Cells(9, 19).FormulaLocal = "=Q9+R9" 'Linea 1: 'T. Finalización Maq 2
    Sheets("Fase1").Cells(9, 20).FormulaLocal = "=S9" 'Linea 1: T. comienzo Maq 3
    Sheets("Fase1").Cells(9, 21).FormulaLocal = "=R9" 'Linea 1: T. Proceso Maq 3
    Sheets("Fase1").Cells(9, 22).FormulaLocal = "=T9+U9" 'Linea 1: T. Finalización Maq 3
    Sheets("Fase1").Cells(9, 23).FormulaLocal = "=V9" 'Linea 1: T. comienzo Maq 4
    Sheets("Fase1").Cells(9, 24).FormulaLocal = "=U9" 'Linea 1: T. Proceso Maq 4
    Sheets("Fase1").Cells(9, 25).FormulaLocal = "=W9+X9" 'Linea 1: T. Finalización Maq 4
    Sheets("Fase1").Cells(9, 26).FormulaLocal = "=M7" 'Linea 2: T. Comienzo Maq 1
    Sheets("Fase1").Cells(9, 27).FormulaLocal = "=M7" 'Linea 2: T. Proceso Maq 1
    Sheets("Fase1").Cells(9, 28).FormulaLocal = "=M7" 'Linea 2: T. Finalización Maq 1
    Sheets("Fase1").Cells(9, 29).FormulaLocal = "=M7" 'Linea 2: Tiempo Preparación Maq 2
    Sheets("Fase1").Cells(9, 30).FormulaLocal = "=M7" 'Linea 2: T. Comienzo Maq 2
    Sheets("Fase1").Cells(9, 31).FormulaLocal = "=M7" 'Linea 2: T. Proceso Maq 2
    Sheets("Fase1").Cells(9, 32).FormulaLocal = "=M7" 'Linea 2: T. Finalización Maq 2
    Sheets("Fase1").Cells(9, 33).FormulaLocal = "=M7" 'Linea 2: T. Comienzo Maq 3
    Sheets("Fase1").Cells(9, 34).FormulaLocal = "=M7" 'Linea 2: T. Proceso Maq 3
    Sheets("Fase1").Cells(9, 35).FormulaLocal = "=M7" 'Linea 2: T. Finalización Maq 3
    Sheets("Fase1").Cells(9, 36).FormulaLocal = "=M7" 'Linea 2: T. Comienzo Maq 4
    Sheets("Fase1").Cells(9, 37).FormulaLocal = "=M7" 'Linea 2: T. Proceso Maq 4
    Sheets("Fase1").Cells(9, 38).FormulaLocal = "=M7" 'Linea 2: T. Finalización Maq 4
    Sheets("Fase1").Cells(9, 39).FormulaLocal = "=$B$5" 'T. traslado
    Sheets("Fase1").Cells(9, 40).FormulaLocal = "=SI(M10=$Z$7;AL10/16+AM10/24;Y10/16+AM10/24)" 'Día de Llegada
    Sheets("Fase1").Cells(9, 41).FormulaLocal = "=AN9-K9" 'Días Restantes
    Sheets("Fase1").Cells(9, 42).FormulaLocal = "=MAX(0;AO9)" 'Atrasos
    Sheets("Fase1").Cells(9, 43).FormulaLocal = "=SI(L9=""LINEA 1"";P9;AC9)" 'SOLUCION GVNS: Tiempos de Preparación
    
    ' "Segunda Fila"
    Sheets("Fase1").Cells(10, 12).FormulaLocal = "LINEA 2" 'LINEA PROD
    Sheets("Fase1").Cells(10, 13).FormulaLocal = "=$Z$7" 'Linea 1: T. Comienzo Maq 1
    Sheets("Fase1").Cells(10, 14).FormulaLocal = "=$M$10" 'Linea 1: T. Proceso Maq 1
    Sheets("Fase1").Cells(10, 15).FormulaLocal = "=$N$10" 'Linea 1: T. Finalización Maq 1
    Sheets("Fase1").Cells(10, 16).FormulaLocal = "=$O$10" 'Linea 1: Tiempo Preparación Maq 2
    Sheets("Fase1").Cells(10, 17).FormulaLocal = "=$P$10" 'Linea 1: T. Comienzo Maq 2
    Sheets("Fase1").Cells(10, 18).FormulaLocal = "=$Q$10" 'Linea 1: T. Proceso Maq 2
    Sheets("Fase1").Cells(10, 19).FormulaLocal = "=$R$10" 'Linea 1: T. Finalización Maq 2
    Sheets("Fase1").Cells(10, 20).FormulaLocal = "=$Z$7" 'Linea 1: T. comienzo Maq 3
    Sheets("Fase1").Cells(10, 21).FormulaLocal = "=$Z$7" 'Linea 1: T. Proceso Maq 3
    Sheets("Fase1").Cells(10, 22).FormulaLocal = "=$Z$7" 'Linea 1: T. Finalización Maq 3
    Sheets("Fase1").Cells(10, 23).FormulaLocal = "=$Z$7" 'Linea 1: T. comienzo Maq 4
    Sheets("Fase1").Cells(10, 24).FormulaLocal = "=$Z$7" 'Linea 1: T. Proceso Maq 4
    Sheets("Fase1").Cells(10, 25).FormulaLocal = "=$Z$7" 'Linea 1: T. Finalización Maq 4
    Sheets("Fase1").Cells(10, 26).FormulaLocal = "=$B$6" 'Linea 2: T. Comienzo Maq 1
    Sheets("Fase1").Cells(10, 27).FormulaLocal = "=BUSCARV(D10;'Tiempos Procesos y Preparacion'!$G$5:$J$116;3;FALSO)*G10" 'Linea 2: T. Proceso Maq 1
    Sheets("Fase1").Cells(10, 28).FormulaLocal = "=AA10" 'Linea 2: T. Finalización Maq 1
    Sheets("Fase1").Cells(10, 29).FormulaLocal = "=0" 'Linea 2: Tiempo Preparación Maq 2
    Sheets("Fase1").Cells(10, 30).FormulaLocal = "=AC10+AB10" 'Linea 2: T. Comienzo Maq 2
    Sheets("Fase1").Cells(10, 31).FormulaLocal = "=BUSCARV(D10;'Tiempos Procesos y Preparacion'!$G$5:$J$116;4;FALSO)*G10/3" 'Linea 2: T. Proceso Maq 2
    Sheets("Fase1").Cells(10, 32).FormulaLocal = "=AD10+AE10" 'Linea 2: T. Finalización Maq 2
    Sheets("Fase1").Cells(10, 33).FormulaLocal = "=AF10" 'Linea 2: T. Comienzo Maq 3
    Sheets("Fase1").Cells(10, 34).FormulaLocal = "=AE10" 'Linea 2: T. Proceso Maq 3
    Sheets("Fase1").Cells(10, 35).FormulaLocal = "=AG10+AH10" 'Linea 2: T. Finalización Maq 3
    Sheets("Fase1").Cells(10, 36).FormulaLocal = "=AI10" 'Linea 2: T. Comienzo Maq 4
    Sheets("Fase1").Cells(10, 37).FormulaLocal = "=AI10" 'Linea 2: T. Proceso Maq 4
    Sheets("Fase1").Cells(10, 38).FormulaLocal = "=AJ10+AK10" 'Linea 2: T. Finalización Maq 4
    Sheets("Fase1").Cells(10, 39).FormulaLocal = "=$B$5" 'T. traslado
    Sheets("Fase1").Cells(10, 40).FormulaLocal = "=SI(M10=$Z$7;AL10/16+AM10/24;Y10/16+AM10/24)" 'Día de Llegada
    Sheets("Fase1").Cells(10, 41).FormulaLocal = "=AN10-K10" 'Dias restantes
    Sheets("Fase1").Cells(10, 42).FormulaLocal = "=MAX(0;AO10)" 'F.O
    Sheets("Fase1").Cells(10, 43).FormulaLocal = "=SI(L10=""LINEA 1"";P10;AC10)" 'SOLUCION GVNS: Tiempos de Preparación
    
    ContadorPedidos = Sheets("Valores").Cells(2, 2).Value
    'Generando datos desde el pedido 3 en adelante
    j = 1
    For i = puntero - 11 To ContadorPedidos - 3 'Itera n-3 veces; i indica posicion-Fila para la ficha Pedidos
        Sheets("Fase1").Cells(10 + j, 12).FormulaLocal = "=SI(M" & 10 + j & "=""LINEA 2"";""LINEA 2"";""LINEA 1"")"
        Sheets("Fase1").Cells(10 + j, 13).FormulaLocal = "=SI(MAX($Q$9:Q" & 9 + j & ")<MAX($AD$9:AD" & 9 + j & ");MAX($Q$9:Q" & 9 + j & ");$Z$7)" 'Linea 1: T. Comienzo Maq 1
        Sheets("Fase1").Cells(10 + j, 14).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;$Z$7;BUSCARV(D" & 10 + j & ";'Tiempos Procesos y Preparacion'!$G$5:$J$116;3;FALSO)*G" & 10 + j & ")" 'Linea 1: T. Proceso Maq 1
        Sheets("Fase1").Cells(10 + j, 15).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;$Z$7;M" & 10 + j & "+N" & 10 + j & ")" 'Linea 1: T. Finalización Maq 1
        Sheets("Fase1").Cells(10 + j, 17).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;$Z$7;P" & 10 + j & "+O" & 10 + j & ")" 'Linea 1: T. Comienzo Maq 2
        Sheets("Fase1").Cells(10 + j, 18).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;$Z$7;(BUSCARV(D" & 10 + j & ";'Tiempos Procesos y Preparacion'!$G$5:$J$116;4;FALSO)*G" & 10 + j & ")/3)" 'Linea 1: T. Proceso Maq 2
        Sheets("Fase1").Cells(10 + j, 19).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;$Z$7;Q" & 10 + j & "+R" & 10 + j & ")" 'Linea 1: 'T. Finalización Maq 2
        Sheets("Fase1").Cells(10 + j, 20).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;$Z$7;MAX(MAX($T$9:T" & 9 + j & ");S" & 10 + j & "))" 'Linea 1: T. comienzo Maq 3
        Sheets("Fase1").Cells(10 + j, 21).FormulaLocal = "=R" & 10 + j 'Linea 1: T. Proceso Maq 3
        Sheets("Fase1").Cells(10 + j, 22).FormulaLocal = "=SI(T" & 10 + j & "=$Z$7;$Z$7;T" & 10 + j & "+U" & 10 + j & ")" 'Linea 1: T. Finalización Maq 3
        Sheets("Fase1").Cells(10 + j, 23).FormulaLocal = "=SI(T" & 10 + j & "=$Z$7;$Z$7;MAX(MAX($W$9:W" & 9 + j & ");V" & 10 + j & "))" 'Linea 1: T. comienzo Maq 4
        Sheets("Fase1").Cells(10 + j, 24).FormulaLocal = "=U" & 10 + j 'Linea 1: T. Proceso Maq 4
        Sheets("Fase1").Cells(10 + j, 25).FormulaLocal = "=SI(T" & 10 + j & "=$Z$7;$Z$7;W" & 10 + j & "+X" & 10 + j & ")" 'Linea 1: T. Finalización Maq 4
        Sheets("Fase1").Cells(10 + j, 26).FormulaLocal = "=SI(MAX($AD$9:AD" & 9 + j & ")<MAX($Q$9:Q" & 9 + j & ");MAX($AD$9:AD" & 9 + j & ");$M$7)" 'Linea 2: T. Comienzo Maq 1
        Sheets("Fase1").Cells(10 + j, 27).FormulaLocal = "=SI(Z" & 10 + j & "=$M$7;$M$7;BUSCARV(D" & 10 + j & ";'Tiempos Procesos y Preparacion'!$G$5:$J$116;3;FALSO)*G" & 10 + j & ")" 'Linea 2: T. Proceso Maq 1
        Sheets("Fase1").Cells(10 + j, 28).FormulaLocal = "=SI(Z" & 10 + j & "=$M$7;$M$7;Z" & 10 + j & "+AA" & 10 + j & ")" 'Linea 2: T. Finalización Maq 1
        Sheets("Fase1").Cells(10 + j, 30).FormulaLocal = "=SI(Z" & 10 + j & "=$M$7;$M$7;AB" & 10 + j & "+AC" & 10 + j & ")"  'Linea 2: T. Comienzo Maq 2
        Sheets("Fase1").Cells(10 + j, 31).FormulaLocal = "=SI(Z" & 10 + j & "=$M$7;$M$7;(BUSCARV(D" & 10 + j & ";'Tiempos Procesos y Preparacion'!$G$5:$J$116;4;FALSO)*G" & 10 + j & ")/3)" 'Linea 2: T. Proceso Maq 2
        Sheets("Fase1").Cells(10 + j, 32).FormulaLocal = "=SI(Z" & 10 + j & "=$M$7;$M$7;AD" & 10 + j & "+AE" & 10 + j & ")" 'Linea 2: T. Finalización Maq 2
        Sheets("Fase1").Cells(10 + j, 33).FormulaLocal = "=SI(L" & 10 + j & "=$M$7;$M$7;MAX(MAX($AG$9:AG" & 9 + j & ");AF" & 10 + j & "))" 'Linea 2: T. Comienzo Maq 3
        Sheets("Fase1").Cells(10 + j, 34).FormulaLocal = "=AE" & 10 + j & "" 'Linea 2: T. Proceso Maq 3
        Sheets("Fase1").Cells(10 + j, 35).FormulaLocal = "=SI(AG" & 10 + j & "=$M$7;$M$7;AG" & 10 + j & "+AH" & 10 + j & ")" 'Linea 2: T. Finalización Maq 3
        Sheets("Fase1").Cells(10 + j, 36).FormulaLocal = "=SI(L" & 10 + j & "=$M$7;$M$7;MAX(MAX($AJ$9:AJ" & 9 + j & ");AI" & 10 + j & "))" 'Linea 2: T. Comienzo Maq 4
        Sheets("Fase1").Cells(10 + j, 37).FormulaLocal = "=AH" & 10 + j & "" 'Linea 2: T. Proceso Maq 4
        Sheets("Fase1").Cells(10 + j, 38).FormulaLocal = "=SI(AG" & 10 + j & "=$M$7;$M$7;AJ" & 10 + j & "+AK" & 10 + j & ")" 'Linea 2: T. Finalización Maq 4
        Sheets("Fase1").Cells(10 + j, 39).FormulaLocal = "=$B$5" 'Tiempo de traslado
        Sheets("Fase1").Cells(10 + j, 40).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;AL" & 10 + j & "/16+AM" & 10 + j & "/24;Y" & 10 + j & "/16+AM" & 10 + j & "/24)" 'Día de Llegada
        Sheets("Fase1").Cells(10 + j, 41).FormulaLocal = "=AN" & 10 + j & "-K" & 10 + j & "" 'Dias restantes
        Sheets("Fase1").Cells(10 + j, 42).FormulaLocal = "=MAX(0;AO" & 10 + j & ")" 'Atrasos
        Sheets("Fase1").Cells(10 + j, 43).FormulaLocal = "=SI(L" & 10 + j & "=""LINEA 1"";P" & 10 + j & ";AC" & 10 + j & ")" 'SOLUCION GVNS: Tiempos de Preparación
        j = j + 1
    Next
    GeneraTiempoPreparacion puntero 'Genera la lista de los tiempos de preparacion de la maquina 2, para la Linea 1 y 2
    CuadraditoResumenConResultados
    OrdenarPedidosCalibreFechaEntrega 'Funcion que toma los pedidos de la ficha Fase1, y los deja ordenados en la ficha PedidosOrdenados
End Function

'Genera los tiempos de preparacion, en caso de ser 0 o 0,35
'@param puntero Indica la posicion-Fila donde comienzan los pedidos en la Ficha Pedidos
'@return
Public Function GeneraTiempoPreparacion(puntero As Integer)
    On Error GoTo ErrorHandler
    Dim x As Integer
    Dim i As Integer
    Dim j As Integer
    Dim z As Integer
    Dim contador As Integer
    For z = 1 To 3 Step 1 'Fuerza la actualizacion de las columnas
        For x = 13 To 0 Step -13 'Avanza en las columnas, primero llena las de la LINEA 2, luego LINEA 1
        j = 1 'Indicador de la fila de lectura para la ficha Fase1
            For i = puntero - 11 To Sheets("Valores").Cells(2, 2).Value - 3 'Recorre desde el tercer al ultimo pedido
                If Sheets("Fase1").Cells(10 + j, 13 + x) = "LINEA 1" Then 'Si el pedido esta en la Linea 1
                   Sheets("Fase1").Cells(10 + j, 16 + x).FormulaLocal = "=SI(Z" & 10 + j & "=$M$7;$M$7;SI(E" & 10 + j & "=E10;0;0,35))" 'Ingresa en la celda LINEA 1
                ElseIf Sheets("Fase1").Cells(10 + j, 13 + x) = "LINEA 2" Then 'Si el pedido esta en la linea 2
                   Sheets("Fase1").Cells(10 + j, 16 + x).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;$Z$7;SI(E" & 10 + j & "=E9;0;0,35))" 'Ingresa en la celda LINEA 2
                Else 'Caso en el que tenga que buscar un pedido anterior
                    contador = 1 'Usado para ubicar la posicion del pedido anterior en la ficha Fase1
                    Do While Sheets("Fase1").Cells(10 + j - contador, 16 + x).Value = "LINEA 1" Or Sheets("Fase1").Cells(10 + j - contador, 16 + x).Value = "LINEA 2" 'Va retrocediendo, hasta que encuentre un pedido
                        contador = contador + 1 'Incrementa la posicion
                    Loop
                    If x = 0 Then 'x=0, indica la posicion de la columna tiempo de preparacion para la linea 1
                        'Ingresa la formula en la posicion actual, y agrega la posicion del pedido anterior en la formula
                        Sheets("Fase1").Cells(10 + j, 16 + x).FormulaLocal = "=SI(M" & 10 + j & "=$Z$7;$Z$7;SI(E" & 10 + j & "=E" & 10 + j - contador & ";0;0,35))"
                    ElseIf x = 13 Then 'x=13, indica la posicion de la columna tiempo de preparacion para la linea 2
                         'Ingresa la formula en la posicion actual, y agrega la posicion dentro de la formula de la posicion del pedido anterior
                        Sheets("Fase1").Cells(10 + j, 16 + x).FormulaLocal = "=SI(Z" & 10 + j & "=$M$7;$M$7;SI(E" & 10 + j & "=E" & 10 + j - contador & ";0;0,35))"
                    End If
                    contador = 1
                End If
                j = j + 1 'Avanzando por Fila
            Next
        Next
    Next
Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 13 'En caso de un mal movimiento , da error numero 13, la FO entrega un resultado erroneo
            LimpiarFichaEnCasoDeError 'Borra todo de la ficha fase1
            CopiarTemporalAFase1 'Copia la secuencia de pedidos de la ficha MejorSecuencia a Fase1
            GenerarBatch
            ColumnaHolgura
            GenerarFormulasLinea1Linea2 11 'ingresa otra vez las formulas a la ficha fase1
            'Resume Next 'Sigue con la siguiente linea del codigo
    End Select
End Function

Public Function CuadraditoResumenConResultados()
    Dim TotalPedidos As Integer
    TotalPedidos = select_count_from("Fase1", 9, 3) + 8 'cuenta el total de pedidos de la ficha Fase1, el +8 le agrega la posicion desde donde empiezan los pedidos
    Sheets("Fase1").Cells(7, 46).FormulaLocal = "=SUMA(AP9:AP" & TotalPedidos & ")" 'Total FO1 (VNS)
    Sheets("Fase1").Cells(8, 46).FormulaLocal = "=SUMA(AQ9:AQ" & TotalPedidos & ")" 'Total FO2 (GVNS)
    Sheets("Fase1").Cells(9, 46).FormulaLocal = "=MAX(AN9:AN" & TotalPedidos & ")" 'Makespan
    Sheets("Fase1").Cells(10, 46).FormulaLocal = "=SUMA('Tiempo muerto'!E4:E79)" 'Tiempo muerto L1
    Sheets("Fase1").Cells(11, 46).FormulaLocal = "=SUMA('Tiempo muerto'!J4:J96)" 'Tiempo muerto L2
    ':)
End Function

'Hace movimientos de pedidos level-veces
'@param level Numero de movimientos a realizar
'@return retorna el valor de la funcion objetivo
Public Function Perturbation(level As Integer) As Double
    On Error GoTo ErrorHandler
    Dim movimiento As Integer
    Dim i As Integer
    If x_infactible Then 'Se ejecuta si FO no es factible
        For i = 1 To level
            movimiento = Aleatorio(1, 4) 'Selecciona un movimiento random, de 1 a 4
            Select Case movimiento
                Case 1
                        moverPedido ("forwardAtrasado")
                Case 2
                        moverPedido ("backwardNoAtrasado")
                Case 3
                        moverPedido ("forwardNoAtrasado")
                Case 4
                        moverPedido ("backwardAtrasado")
            End Select
        Next i
        GeneraTiempoPreparacion 11 'Actualiza los tiempos de preparacion para las 2 lineas
    End If
    Perturbation = FuncionObjetivoGlobal
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Indica si una solucion es factible
'@param
'@return True si la FO es factible, False en caso contrario
Public Function x_factible() As Boolean
    On Error GoTo ErrorHandler
    If FuncionObjetivoGlobal = 0 Then
        x_factible = True
    Else
        x_factible = False
    End If
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Indica si la solucion es infactible
'@param
'@return True si la FO es infactible, False en caso contrario
Public Function x_infactible() As Boolean
    On Error GoTo ErrorHandler
    If FuncionObjetivoGlobal = 0 Then
        x_infactible = False
    Else
        x_infactible = True
    End If
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Ordena los pedidos por fecha de entrega, desde la mas cercana a la más lejana
'@param
'@return Valor de la Funcion Objetivo
Public Function RandomSolution() 'As Double
    On Error GoTo ErrorHandler
    Dim rango As String
    Dim RangoOrdenar As String
    Dim i As Integer
    Dim suma As Integer
    Dim pedidosVNS As Integer
    Sheets("Fase1").Select
    pedidosVNS = select_count_from("Fase1", 9, 3)
    suma = 0
    rango = "C9:J" & pedidosVNS + 8 'Es 8 porque esos espacios estan en blanco, asi que los pedidos empiezan de la fila 9
    RangoOrdenar = "J9:J" & pedidosVNS + 8 'Es 8 porque esos espacios estan en blanco, asi que los pedidos empiezan de la fila 9
    Range(rango).Select
    ActiveWorkbook.Worksheets("Fase1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Fase1").Sort.SortFields.Add Key:=Range(RangoOrdenar), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Fase1").Sort
        .SetRange Range(rango)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'RandomSolution = FuncionObjetivo
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'La suma de las funciones objetivo
'@param
'@return FO, corresponde a la suma de FO individuales
Public Function FuncionObjetivoGlobal() As Double
    On Error GoTo ErrorHandler
    FuncionObjetivoGlobal = Sheets("Fase1").Cells(7, 46).Value 'Corresponde al valor que esta en Total FO1 (VNS)
Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 13 'En caso de desbordamiento
            Sheets("Valores").Cells(11, 2).Value = "11" 'En la ficha Valores, ingresa el valor por defecto 11
    End Select
End Function

'Realiza 4 movimientos, en orden forwardAtrasado, backwardNoAtrasado, forwardNoAtrasado, backwardAtrasado
'@param
'@return valor de la Función Objetivo
Public Function Local1Shift() As Double
    On Error GoTo ErrorHandler
    If x_infactible Then 'Se ejecuta si FO no es factible
        moverPedido "forwardAtrasado"    '1º movimiento
        moverPedido "backwardNoAtrasado" '2º movimiento
        moverPedido "forwardNoAtrasado"  '3º movimiento
        moverPedido "backwardAtrasado"   '4º movimiento
        GeneraTiempoPreparacion 11
    End If
    Local1Shift = FuncionObjetivoGlobal 'Retorna el valor de FO
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Obtiene la lista de Pedidos Atrasados, indica la posicion dentro del excel donde se encuentran
'@param
'@return arreglo con las posiciones del excel
Public Function ListaPedidosAtrasados() As Integer()
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim posicionLectura As Integer
    Dim posicionPedidoAtrasado As Integer
    Dim resultado(1000) As Integer
    i = 0
    posicionPedidoAtrasado = 42 'Columna donde evalua si esta atrasado, corresponde a la DIAS DE ATRASO
    posicionLectura = 9 'Posicion Fila donde empiezan los pedidos en la ficha Fase1
    Do While Sheets("Fase1").Cells(posicionLectura, 3).Value <> "" 'Se lee mientras existan datos
        If Sheets("Fase1").Cells(posicionLectura, posicionPedidoAtrasado).Value > 0 Then 'Si es un pedido atrasado, lo agrega a la lista
            resultado(i) = posicionLectura 'Guarda el valor en el arreglo
            i = i + 1 'Avanzando en las posiciones del arreglo
        End If
        posicionLectura = posicionLectura + 1 'Avanzando en las posiciones del Excel
    Loop
    ListaPedidosAtrasados = resultado 'Retornando la lista de pedidos atrasados
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Obtiene la lista de Pedidos No Atrasados, indica la posicion dentro del excel donde se encuentran
'@param
'@return arreglo con las posiciones del excel
Public Function ListaPedidosNoAtrasados() As Integer() 'Retorna la lista de los pedidos No Atrasados
    On Error GoTo ErrorHandler
    Dim i As Integer                'Posicion en el arreglo
    Dim posicionLectura As Integer  'Posicion en el excel
    Dim posicionPedidoAtrasado As Integer 'es la columna del excel donde estan
    Dim resultado(1000) As Integer
    i = 0
    posicionPedidoAtrasado = 42 'Columna donde evalua si esta atrasado, corresponde a la DIAS DE ATRASO
    posicionLectura = 9 'Posicion Fila donde empiezan los pedidos en la ficha Fase1
    Do While Sheets("Fase1").Cells(posicionLectura, 3).Value <> "" 'Se lee mientras existan datos
        If Sheets("Fase1").Cells(posicionLectura, posicionPedidoAtrasado).Value <= 0 Then
            resultado(i) = posicionLectura
            i = i + 1
        End If
        posicionLectura = posicionLectura + 1
    Loop
    ListaPedidosNoAtrasados = resultado
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Realiza movimientos de pedidos, que pueden ser forwardAtrasado, backwardNoAtrasado, forwardNoAtrasado, backwardAtrasado
'@param movimiento Indica el tipo de movimiento a efectuar
'@return arreglo con las posiciones del excel
Public Sub moverPedido(movimiento As String)
    On Error GoTo ErrorHandler
    Dim primerPedido As Integer
    Dim PedidoMover As Integer
    Dim nuevaPosicion As Integer
    Dim PedidosAtrasados() As Integer
    Dim PedidosNoAtrasados() As Integer
    Dim pedidosVNS As Integer
    
    pedidosVNS = select_count_from("Fase1", 9, 3) 'Total de pedidos en la Ficha Fase1
    primerPedido = 9 'Posicion Fila del primer pedido en la Ficha Fase1
    
    Select Case movimiento
        Case "forwardAtrasado"
            PedidosAtrasados = ListaPedidosAtrasados() 'Obtiene una lista de pedidos atrasados
            If contarElementosArreglo(PedidosAtrasados) > 0 Then 'Si hay pedidos atrasados, hace el movimiento
                PedidoMover = PedidoAleatorio(PedidosAtrasados()) 'De ListaPedidosAtrasados toma un pedido aleatorio; PedidoMover es la posicion del excel que será movida
                nuevaPosicion = PedidoMover - 1 'Es la nueva posicion, 1 arriba
                If nuevaPosicion < primerPedido Then 'En caso de que el movimiento se salga del rango de pedidos
                    nuevaPosicion = primerPedido 'Le asigna su misma posición, por ejemplo, si el pedido 9 se mueve a la posición 8, su nueva posición sera la 9
                End If
                LeeryGuardarPedido nuevaPosicion, PedidoMover  'Se encarga de cambiar la posicion de los pedidos en el Excel
            End If
        Case "backwardNoAtrasado"
            PedidosNoAtrasados = ListaPedidosNoAtrasados() 'Obtiene una lista de pedidos No atrasados
            If contarElementosArreglo(PedidosNoAtrasados) > 0 Then
                PedidoMover = PedidoAleatorio(PedidosNoAtrasados()) 'De esa lista, toma un pedido aleatorio
                nuevaPosicion = PedidoMover + 1 'Es la nueva posicion, 1 abajo
                If nuevaPosicion > pedidosVNS + 8 Then 'En caso de salirse del limite del rango de pedidos
                    nuevaPosicion = pedidosVNS + 8
                End If
                LeeryGuardarPedido nuevaPosicion, PedidoMover  'Se encarga de cambiar la posicion de los pedidos en el Excel
            End If
        Case "forwardNoAtrasado"
            PedidosNoAtrasados = ListaPedidosNoAtrasados() 'Obtiene una lista de pedidos No atrasados
            If contarElementosArreglo(PedidosNoAtrasados) > 0 Then
                PedidoMover = PedidoAleatorio(PedidosNoAtrasados()) 'De esa lista, toma un pedido aleatorio
                nuevaPosicion = PedidoMover - 1 'Es la nueva posicion, 1 arriba
                If nuevaPosicion < primerPedido Then
                    nuevaPosicion = primerPedido
                End If
                LeeryGuardarPedido nuevaPosicion, PedidoMover  'Se encarga de cambiar la posicion de los pedidos en el Excel
            End If
        Case "backwardAtrasado"
            PedidosAtrasados = ListaPedidosAtrasados() 'Obtiene una lista de pedidos atrasados
            If contarElementosArreglo(PedidosAtrasados) > 0 Then
                PedidoMover = PedidoAleatorio(PedidosAtrasados()) 'De esa lista, toma un pedido aleatorio
                nuevaPosicion = PedidoMover + 1 'Es la nueva posicion, 1 abajo
                If nuevaPosicion > pedidosVNS + 8 Then
                    nuevaPosicion = pedidosVNS + 8
                End If
                LeeryGuardarPedido nuevaPosicion, PedidoMover  'Se encarga de cambiar la posicion de los pedidos en el Excel
            End If
    End Select
Exit Sub
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Sub

'Hace el movimiento de pedidos, desde una posicion a su nueva posicion
'@param nuevaPosicion numero que indica la nueva posicion dentro del excel
'@param PedidoMover numero que indica el pedido a mover dentro del excel
'@return
Public Function LeeryGuardarPedido(nuevaPosicion As Integer, PedidoMover As Integer)
    On Error GoTo ErrorHandler
    Dim rangePedidotemp As String
    Dim rangePedidoMover As String
    Dim posicionPedidoTemp As Integer
    Dim pedidosVNS As Integer
    
    Sheets("Fase1").Select
    pedidosVNS = select_count_from("Fase1", 9, 3) 'obtiene la cantidad de pedidos en la ficha VNS
    posicionPedidoTemp = pedidosVNS + 14 'la ultima fila donde se guarda el pedido temporal

    '1º Selecciono el pedido destino
    rangePedidotemp = "C" & nuevaPosicion & ":J" & nuevaPosicion 'Se define el rango a mover
    Range(rangePedidotemp).Select 'Selecciona la fila
    Selection.Copy 'Copia la seleccion, deja el espacio libre para los nuevos datos
    Range("C" & posicionPedidoTemp).Select 'selecciona la fila donde pegar los datos de manera temporal
    ActiveSheet.Paste 'Pega los datos en un espacion temporal
    
    '2º Selecciono el pedido a mover
    rangePedidoMover = "C" & PedidoMover & ":J" & PedidoMover
    Range(rangePedidoMover).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C" & nuevaPosicion).Select
    ActiveSheet.Paste
    
    '3º Muevo la fila temporal a su posicion final
    Range("C" & posicionPedidoTemp & ":J" & posicionPedidoTemp).Select 'selecciona la fila donde pegar los datos de manera temporal
    Application.CutCopyMode = False
    Selection.Copy
    Range("C" & PedidoMover).Select
    ActiveSheet.Paste
    
    '4º Borrar el Pedido Temporal
    Range("C" & posicionPedidoTemp & ":J" & posicionPedidoTemp).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("C8").Select

    'ActualizarColumnaTiempoFinalizacion pedidosVNS 'Actualizando la columna de los tiempos de finalizacion
    GeneraTiempoPreparacion 11 'posicion-Fila 11, desde donde empiezan los pedidos del mes en curso en la ficha Pedidos
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Dada una lista de pedidos (atrasados o no atrasados) elige uno random
'@param listaPedidos() lista con los pedidos (atrasados o no atrasados)
'@return posicion del pedido dentro del excel
Public Function PedidoAleatorio(listaPedidos() As Integer) As Integer
    On Error GoTo ErrorHandler
    Dim contador As Integer
    Dim indicePedido As Integer
    contador = 0
    contador = contarElementosArreglo(listaPedidos)
    'Do While listaPedidos(contador) <> 0
    '    contador = contador + 1
    'Loop
    indicePedido = Aleatorio(0, contador - 1) 'Del array, toma un indice random
    PedidoAleatorio = listaPedidos(indicePedido) 'Obtiene el numero del pedido, que corresponde a la posicion dentro de la Ficha Fase1
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Genera numeros aleatorios dentro de un rango
'@param Minimo menor valor a generar inclusive
'@param Maximo mayor valor a generar inclusive
'@return numero aleatorio
Public Function Aleatorio(Minimo As Integer, Maximo As Integer) As Integer
    On Error GoTo ErrorHandler
    Randomize ' inicializar la semilla
    Aleatorio = CInt((Minimo - Maximo) * Rnd + Maximo)
    'Dim value As Integer = CInt(Int((6 * Rnd()) + 1))
    'Aleatorio = CInt((Maximo - Minimo + 1) * Rnd + Minimo)
Exit Function
ErrorHandler:
    Mensaje "Error", Err.Description, "error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Cuenta la cantidad de elementos que tiene un arreglo
'@param listaPedidos() arreglo con la lista de pedidos, pueden ser atrasados, o no atrasados
'@return Cantidad de elementos que tiene el arreglo
Public Function contarElementosArreglo(listaPedidos() As Integer) As Integer
    On Error GoTo ErrorHandler
    Dim contador As Integer
    contador = 0
    Do While listaPedidos(contador) <> 0
        contador = contador + 1
    Loop
    contarElementosArreglo = contador
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Funcion que permite guardar datos en un archivo txt plano.
'@param dato entrada a ser guardada
'@return
Public Function Guardar(dato As Variant)
    On Error GoTo ErrorHandler
    Open Application.ThisWorkbook.Path & "\salida.txt" For Append As #1
    Print #1, dato
    Close #1
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error Guardando Datos" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Borra el archivo salida.txt
'@param
'@return
Public Function BorrarArchivo()
    On Error GoTo ErrorHandler
    Kill Application.ThisWorkbook.Path & "\salida.txt"
Exit Function
ErrorHandler:
    Select Case Err.Number
      Case 53 'El archivo no existe
        Resume Next
      Case 75 'El archivo existe pero es solo lectura
        Mensaje "Error", "El archivo salida.txt es de solo lectura", "error"
      Case Else 'Cualquier otro error no esperado
        Mensaje "Error", Err.Description, "error"
      End Select
End Function

'Si la FO mejora, entonces la secuencia que esta en la Ficha Fase1 se copia a la Ficha MejorSecuencia
'@param
'@return
Public Function CopiarMejorSecuenciaATemporal()
    On Error GoTo ErrorHandler
    Dim rango As Integer
    Dim TotalPedidos As Integer
    TotalPedidos = Sheets("Valores").Cells(2, 2).Value 'Obtiene el total de pedidos, desde la ficha Valores'Borra el contenido de la ficha MejorSecuencia
    Sheets("Fase1").Select
    rango = TotalPedidos + 8
    Range("C9:J" & rango).Select
    Selection.Copy
    Sheets("MejorSecuencia").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Fase1").Select
    Range("AT7").Select 'AU7 es la suma de FO en la Ficha Fase1
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Valores").Select
    Range("B7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False 'Copia el AS9  a la Ficha Valores
    
    Sheets("Fase1").Select
    Range("AT8").Select 'Copia el valor de la suma de los tiempos de preparacion, lo copia a la ficha Valores
    Selection.Copy
    Sheets("Valores").Select
    Range("B9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Fase1").Select
Exit Function
ErrorHandler:
    MsgBox Err.Description, 16, "Error" 'En caso de error, muestra un mensaje con la descripcion
End Function

'Si la FO empeora, entonces la secuencia que esta en la Ficha MejorSecuencia se copia a la Ficha Fase1
'@param
'@return
Public Function CopiarTemporalAFase1()
    Dim rango As Integer
    rango = Sheets("Valores").Cells(2, 2).Value
    Sheets("MejorSecuencia").Select
    Range("A1:H" & rango).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Fase1").Select
    Range("C9").Select
    ActiveSheet.Paste
End Function

'Compara FO nueva con la antigua, y decide si mejoro la FO
'@param
'@return True, en caso de que la FO haya mejorado, False la FO no mejoro
Public Function MejoroFuncion() As Boolean
    'If Sheets("Fase1").Cells(7, 47).Value <= Sheets("Valores").Cells(7, 2).Value Then 'Compara la nueva FO  con la antigua FO
    If FuncionObjetivoGlobal <= Sheets("Valores").Cells(7, 2).Value Then 'Compara la nueva FO  con la antigua FO
        MejoroFuncion = True 'La FO mejoro, retorna True
    Else
        MejoroFuncion = False 'La FO no mejoro, retorna False
    End If
End Function

'Realiza el copiado de pedidos dependiendo si:
'Caso1: La FO mejoro, copia los pedidos de la Ficha Fase1 a la Ficha MejorSecuencia
'Caso2: La FO no mejoro, copia los pedidos de la Ficha MejorSecuencia a la Ficha Fase1
'@param
'@return
Public Function copiadoDePedidos()
    If MejoroFuncion Then
        CopiarMejorSecuenciaATemporal
    Else
        CopiarTemporalAFase1
        GenerarFormulasLinea1Linea2 11
        'Turbo 11
    End If
End Function

Public Function ColumnaHolgura()
    Dim TotalPedidos As Integer
    Dim rango As Integer
    TotalPedidos = Sheets("Valores").Cells(2, 2).Value 'Obtiene el total de pedidos, desde la ficha Valores'Borra el contenido de la ficha MejorSecuencia
    Sheets("Fase1").Select
    rango = TotalPedidos + 8
    Range("K9").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-2]"
    Range("K9").Select
    Selection.AutoFill Destination:=Range("K9:K" & rango)
    Range("K9:K" & rango).Select
End Function
