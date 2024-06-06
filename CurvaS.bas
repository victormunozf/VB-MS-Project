Attribute VB_Name = "CurvaS"
Sub CurvaS()
    ' Definición de Variables
    Dim N, M, K, T As Long
    Dim i, j As Integer
    Dim Baseline As TimeScaleValues
    Dim Timenow As TimeScaleValues
    Dim myBCWS() As TimeScaleValues
    Dim myACWP() As TimeScaleValues
    Dim suma, sumabis, sumatris As Double
    Dim cumtaskBCWS() As Double
    Dim cumtaskBCWP() As Double
    Dim cumtaskACWP() As Double
    Dim auxBCWP() As Double
    Dim auxACWP() As Double
    Dim cumBCWS() As Double
    Dim cumBCWP() As Double
    Dim cumACWP() As Double
    Dim BAC As Double
    Dim Mes() As Integer
    Dim Fecha() As String
    Dim Tnow() As String
    Dim Dif As Integer
    Dim Dur As Long
    Dim BSY, BSM, BSD, AfinB, MfinB, DfinB, _
    Ainicio, Minicio, Dinicio, Afin, Mfin, Dfin, Aestado, Mestado, Destado As Integer
    Dim TimeScale As PjTimescaleUnit
    
    TimeScale = pjTimescaleMonths
    
    ' Verificar línea base y fecha de estado
    If ActiveProject.BaselineStart = "NA" Then
        MsgBox "No hay línea base " & vbCrLf & "Establezca una ", vbExclamation
        Exit Sub
    End If
    
    If ActiveProject.StatusDate = "NA" Then
        MsgBox "Debe asignar una fecha de estado " & vbCrLf & "antes de continuar ", vbExclamation
        Exit Sub
    End If
    
    ' Asignación de fechas
    BSY = Year(ActiveProject.BaselineStart)
    BSM = Month(ActiveProject.BaselineStart)
    BSD = Day(ActiveProject.BaselineStart)
    AfinB = Year(ActiveProject.BaselineFinish)
    MfinB = Month(ActiveProject.BaselineFinish)
    DfinB = Day(ActiveProject.BaselineFinish)
    Ainicio = Year(ActiveProject.ProjectSummaryTask.Start)
    Minicio = Month(ActiveProject.ProjectSummaryTask.Start)
    Dinicio = Day(ActiveProject.ProjectSummaryTask.Start)
    Afin = Year(ActiveProject.ProjectFinish)
    Mfin = Month(ActiveProject.ProjectFinish)
    Dfin = Day(ActiveProject.ProjectFinish)
    Aestado = Year(ActiveProject.StatusDate)
    Mestado = Month(ActiveProject.StatusDate)
    Destado = Day(ActiveProject.StatusDate)
    Dif = Minicio - BSM
    
    ' Trabajo planificado y actual
    Set Baseline = ActiveProject.ProjectSummaryTask.TimeScaleData(StartDate:=ActiveProject.BaselineStart, _
    EndDate:=ActiveProject.BaselineFinish, _
    Type:=pjTaskTimescaledBaselineCost, _
    TimescaleUnit:=TimeScale, _
    count:=1)
    
    Set Timenow = ActiveProject.ProjectSummaryTask.TimeScaleData(StartDate:=ActiveProject.ProjectSummaryTask.Start, _
    EndDate:=ActiveProject.StatusDate, _
    Type:=pjTaskTimescaledCost, _
    TimescaleUnit:=TimeScale, _
    count:=1)
    
    N = Baseline.count
    M = Timenow.count
    Dur = M + Dif
    K = ActiveProject.Tasks.count
    ReDim myBCWS(1 To K) As TimeScaleValues
    ReDim myBCWP(1 To K, 1 To M) As Double
    ReDim myACWP(1 To K) As TimeScaleValues
    ReDim cumtaskBCWS(1 To K, 1 To N) As Double
    ReDim cumtaskBCWP(1 To K, 1 To M) As Double
    ReDim cumtaskACWP(1 To K, 1 To M) As Double
    ReDim auxBCWP(1 To M) As Double
    ReDim auxACWP(1 To M) As Double
    ReDim cumBCWS(1 To N) As Double
    ReDim cumBCWP(1 To M) As Double
    ReDim cumACWP(1 To M) As Double
    ReDim Tnow(1 To M) As String
    
    ' Cálculos
    sumabis = 0
    sumatris = 0
    For T = 1 To K
        If Not ActiveProject.Tasks(T).Summary Then
            ' Obtener costos de línea base y costos actuales
            Set myBCWS(T) = ActiveProject.Tasks(T).TimeScaleData(StartDate:=ActiveProject.BaselineStart, _
            EndDate:=ActiveProject.BaselineFinish, _
            Type:=pjTaskTimescaledBaselineCost, _
            TimescaleUnit:=TimeScale, _
            count:=1)
            Set myACWP(T) = ActiveProject.Tasks(T).TimeScaleData(StartDate:=ActiveProject.ProjectSummaryTask.Start, _
            EndDate:=ActiveProject.StatusDate, _
            Type:=pjTaskTimescaledActualCost, _
            TimescaleUnit:=TimeScale, _
            count:=1)
            For i = 1 To N
                suma = 0
                For j = 1 To i
                    If myBCWS(T)(j) = "" Then
                        suma = suma
                    Else
                        suma = suma + myBCWS(T)(j)
                    End If
                Next j
                cumtaskBCWS(T, i) = suma
            Next i
            sumabis = sumabis + ActiveProject.Tasks(T).PercentComplete * ActiveProject.Tasks(T).BaselineCost / 100
            sumatris = sumatris + ActiveProject.Tasks(T).ActualCost
        End If
    Next T
    auxBCWP(M) = sumabis
    auxACWP(M) = sumatris
    
    For i = 1 To N
        suma = 0
        For T = 1 To K
            suma = suma + cumtaskBCWS(T, i)
        Next T
        cumBCWS(i) = suma
    Next i
    
    ActiveProject.Tasks(M).Number19 = auxBCWP(M)
    ActiveProject.Tasks(M).Number20 = auxACWP(M)
    ActiveProject.Tasks(M).Text20 = DateSerial(Aestado, Mestado, Destado)
    
    For i = 1 To M
        cumBCWP(i) = ActiveProject.Tasks(i).Number19
        cumACWP(i) = ActiveProject.Tasks(i).Number20
        Tnow(i) = ActiveProject.Tasks(i).Text20
    Next i
    
    BAC = cumBCWS(N)
    
    ' Cálculo del valor ganado
    Dim count() As Integer
    Dim ecount() As Double
    Dim Num() As Double
    Dim Den() As Long
    Dim interp() As Double
    Dim AT() As Double
    Dim EScum() As Double
    Dim auxAT() As Double
    Dim auxEScum() As Double
    Dim SVcum() As Double
    Dim PFMF, PLMF, ALMF, SLMF As Double
    
    ' PFMF := Plan First Month Fraction (Start date)
    ' PLMF := Plan Last Month Fraction (Finish date)
    ' ALMF := Actual Last Month Fraction (Finish date)
    ' SLMF := Status Last Month Fraction (Status date)
    
    PFMF = (ActiveProject.Calendar.Years(BSY).Months(BSM).Days.count - BSD + 1) / ActiveProject.Calendar.Years(BSY).Months(BSM).Days.count
    PLMF = DfinB / ActiveProject.Calendar.Years(AfinB).Months(MfinB).Days.count
    ALMF = Dfin / ActiveProject.Calendar.Years(Afin).Months(Mfin).Days.count
    SLMF = Destado / ActiveProject.Calendar.Years(Aestado).Months(Mestado).Days.count
    
    ReDim count(1 To M) As Integer
    ReDim ecount(1 To M) As Double
    ReDim Num(1 To M) As Double
    ReDim Den(1 To M) As Long
    ReDim interp(1 To M) As Double
    ReDim AT(0 To M) As Double
    ReDim EScum(1 To M) As Double
    ReDim auxAT(1 To M) As Double
    ReDim auxEScum(1 To M) As Double
    ReDim SVcum(1 To M) As Double
    ReDim SPIcum(1 To M) As Double
    ReDim TSPI(1 To M) As Double
    
    For j = 1 To M
        suma = 0
        i = 1
        Do While cumBCWS(i) <= cumBCWP(j)
            If i = N Then
                suma = suma + 1
                GoTo fuera:
            End If
            suma = suma + 1
            i = i + 1
        Loop
fuera:
        count(j) = suma
    Next j
    
    AT(0) = 0
    
    For i = 1 To M - 1
        ecount(i) = count(i) - (1 - PFMF)
        
        If ecount(i) < 1 Then
            auxEScum(i) = (i - 1) + ecount(i)
            EScum(i) = auxEScum(i) - 1 + PFMF
        Else
            auxEScum(i) = ecount(i)
            EScum(i) = auxEScum(i)
        End If
        auxAT(i) = i - 1 + PFMF
        AT(i) = auxAT(i)
        SVcum(i) = EScum(i) - AT(i)
        SPIcum(i) = EScum(i) / AT(i)
        TSPI(i) = (BAC - cumBCWS(i)) / (BAC - cumBCWP(i))
    Next i
    
    ecount(M) = count(M) - (1 - PFMF)
    
    If ecount(M) < 1 Then
        auxEScum(M) = (M - 1) + ecount(M)
        EScum(M) = auxEScum(M) - 1 + PFMF
    Else
        auxEScum(M) = ecount(M)
        EScum(M) = auxEScum(M)
    End If
    auxAT(M) = M - 1 + PFMF
    AT(M) = auxAT(M)
    SVcum(M) = EScum(M) - AT(M)
    SPIcum(M) = EScum(M) / AT(M)
    TSPI(M) = (BAC - cumBCWS(M)) / (BAC - cumBCWP(M))
    
    ' Obtener la primera hoja del libro
    Dim Workbook As Object
    Dim Sheet As Object
    
    On Error Resume Next
    Set Workbook = GetObject(, "Excel.Application")
    If Workbook Is Nothing Then
        Set Workbook = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
    
    Workbook.Visible = True
    Workbook.Workbooks.Add
    Set Sheet = Workbook.ActiveSheet
    
    ' Escribir datos a Excel
    With Sheet
        .Cells(1, 1).Value = "Fecha de Línea Base"
        .Cells(1, 2).Value = ActiveProject.BaselineFinish
        .Cells(2, 1).Value = "Fecha de Finalización"
        .Cells(2, 2).Value = ActiveProject.ProjectFinish
        .Cells(3, 1).Value = "Fecha de Estado"
        .Cells(3, 2).Value = ActiveProject.StatusDate
        
        .Cells(5, 1).Value = "CPTP"
        .Cells(5, 2).Value = "CPTR"
        .Cells(5, 3).Value = "CRTR"
        .Cells(5, 4).Value = "Fecha"
        
        For i = 1 To M
            .Cells(i + 5, 1).Value = cumBCWS(i)
            .Cells(i + 5, 2).Value = cumBCWP(i)
            .Cells(i + 5, 3).Value = cumACWP(i)
            .Cells(i + 5, 4).Value = Tnow(i)
        Next i
        
        .Cells(5, 6).Value = "Tiempo Real"
        .Cells(5, 7).Value = "Índice de Rendimiento de la Programación acumulado(IRP)"
        .Cells(5, 8).Value = "Variación de Cronograma acumulada"
        .Cells(5, 9).Value = "Índice de Rendimiento del Costo acumulado(IRC)"
        .Cells(5, 10).Value = "Índice de Rendimiento del Cronograma Puntual"
        
        For i = 1 To M
            .Cells(i + 5, 6).Value = AT(i)
            .Cells(i + 5, 7).Value = EScum(i)
            .Cells(i + 5, 8).Value = SVcum(i)
            .Cells(i + 5, 9).Value = SPIcum(i)
            .Cells(i + 5, 10).Value = TSPI(i)
        Next i
    End With
    
    ' Crear Tablas
    Dim Chart As Object
    Set Chart = Workbook.Charts.Add
    With Chart
        .ChartType = xlLine
        .SetSourceData Source:=Sheet.Range("A6:B" & 6 + M - 1)
        .HasTitle = True
        .ChartTitle.Text = "Avance Físico"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Período"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Gasto Acumulado"
        
              ' Escribir series de nombres
        .SeriesCollection(1).Name = "CPTP"
        .SeriesCollection(1).Border.Color = RGB(0, 112, 192) ' Azul
        .SeriesCollection(2).Name = "CPTR"
        .SeriesCollection(2).Border.Color = RGB(112, 173, 71) ' Verde
        
    End With
    
    Set Chart = Workbook.Charts.Add
    With Chart
        .ChartType = xlLine
        .SetSourceData Source:=Sheet.Range("B6:C" & 6 + M - 1)
        .HasTitle = True
        .ChartTitle.Text = "Avance Financiero"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Período"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Gasto Acumulado"
        
              ' Escribir series de nombres
        .SeriesCollection(1).Name = "CPTR"
        .SeriesCollection(1).Border.Color = RGB(112, 173, 71) ' Verde
        .SeriesCollection(2).Name = "CRTR"
        .SeriesCollection(2).Border.Color = RGB(192, 0, 0) ' Rojo
        
    End With
    
    Set Chart = Workbook.Charts.Add
    With Chart
        .ChartType = xlLine
        .SetSourceData Source:=Sheet.Range("A6:C" & 6 + M - 1)
        .HasTitle = True
        .ChartTitle.Text = "Curva S"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Período"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Gasto Acumulado"
        
              ' Escribir series de nombres
        .SeriesCollection(1).Name = "CPTP"
        .SeriesCollection(1).Border.Color = RGB(0, 112, 192) ' Azul
        .SeriesCollection(2).Name = "CPTR"
        .SeriesCollection(2).Border.Color = RGB(112, 173, 71) ' Verde
        .SeriesCollection(3).Name = "CRTR"
        .SeriesCollection(3).Border.Color = RGB(192, 0, 0) ' Rojo
                
    End With
    
    
    ' Mostrar mensaje
    MsgBox "Resultados exportados a Excel correctamente", vbInformation
    
End Sub

