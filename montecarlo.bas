'Beta: Valor "V" de la variable dividido por su valor esperado VE.
        
'Función para calcular la probabilidad en la que se cambia de función beta (área izquierda o derecha).
Function Pcamb(bmin As Double, bmax As Double) As Double
    Pcamb = (1 - bmin) / (bmax - bmin)
    'bmin: beta mínimo
    'bmax: beta máximo
End Function

'Función para hallar beta cuando el área es a la izquierda
Function beta1(Pac, bmin, bmax)
    beta1 = bmin + (Pac * (bmax - bmin) * (1 - bmin)) ^ 0.5 'Pac: Probalidad acumulada. Número random con el que se calcularán los valores de la variable.
End Function

'Función para hallar beta cuando el área es a la derecha
Function beta2(Pac, bmin, bmax)
    beta2 = bmax - ((1 - Pac) * (bmax - bmin) * (bmax - 1)) ^ 0.5
End Function

Sub MontecarloTriangular() '(Para mejor entendimiento leer copias economía (PI 510.pdf) páginas 453-466.)

'Número de iteraciones
it = 100

'Número de variables
NumVar = Cells(37, 2)

'Declaración de variables de funciones
Dim bmin As Double
Dim bmax As Double
Dim Vesp As Double


For i = 1 To it
'Detener cálculo automático de Excel de la hoja de datos para agilizar programa
Application.Calculation = xlManual

    'Loop para calcular beta y el valor aleatorio de cada variable
    For j = 1 To NumVar
        'Beta mínimo y máximo
        bmin = Cells(j + 39, 3)
        bmax = Cells(j + 39, 4)
        
        'Valor esperado
        Vesp = Cells(j + 39, 5)
        
        'Número entre 0 y 1 al azar
        Randomize
        P = Rnd
    
        'Cálculo de beta con probabilidad P
        If P <= Pcamb(bmin, bmax) Then
        beta = beta1(P, bmin, bmax)
        Else
        beta = beta2(P, bmin, bmax)
        End If
        
        'Cálculo y reemplazo de nuevo valor
        Cells(j + 39, 6) = beta * Vesp
    Next j

'Recalcular con los nuevos valores
Application.Calculation = xlAutomatic

'Valor de indicador en tabla
Cells(i + 399, 2) = i / 100
Cells(i + 399, 3) = Cells(395, 3)
Cells(i + 399, 4) = Cells(396, 3)

Next i


'Ordenar valores de VAN de menor a mayor
    Range("C400:C499").Select
    ActiveWorkbook.Worksheets("Var + VAN (Test 2)").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Var + VAN (Test 2)").Sort.SortFields.Add Key:= _
        Range("C400"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Var + VAN (Test 2)").Sort
        .SetRange Range("C400:C499")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D400:D499").Select
    ActiveWorkbook.Worksheets("Var + VAN (Test 2)").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Var + VAN (Test 2)").Sort.SortFields.Add Key:= _
        Range("D400"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Var + VAN (Test 2)").Sort
        .SetRange Range("D400:D499")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub
'Ed1123

'Programa para calcular los estados finacieros con los valores esperados.
Sub ValorEsperado()

'Número de variables
NumVar = Cells(37, 2)

'Loop para reemplzarar variables el valor esperado de cada variable
For j = 1 To NumVar
    'Valor esperado
    Vesp = Cells(j + 39, 5)
   
    'Reemplazo del valor esperado
    Cells(j + 39, 6) = Vesp
Next j

End Sub

