Attribute VB_Name = "Recursos"
Sub Recursos()
    ' Variables para Project
    Dim Proj As Project
    Dim Recurso As Resource
    Dim Asignacion As Assignment
    Dim Tarea As Task
    
    ' Variables para Excel
    Dim ExcelApp As Object
    Dim ExcelLibro As Object
    Dim ExcelHoja As Object
    Dim ExcelRango As Object
    
    ' Crear instancia de Excel
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = True
    Set ExcelLibro = ExcelApp.Workbooks.Add
    Set ExcelHoja = ExcelLibro.Sheets(1)
    
    ' Agregar título
    Dim Titulo As String
    Titulo = "Lista de Recursos - " & ActiveProject.Name & " - " & Format(Date, "dd/mm/yyyy")
    ExcelHoja.Cells(1, 1).Value = Titulo
    ExcelHoja.Range("A1:F1").Merge
    ExcelHoja.Range("A1:F1").Font.Bold = True
    ExcelHoja.Range("A1:F1").Font.Size = 14
    ExcelHoja.Rows("1:1").HorizontalAlignment = xlCenter
    
    ' Encabezados de la tabla
    With ExcelHoja
        .Cells(3, 1).Value = "Nombre del Recurso"
        .Cells(3, 2).Value = "Tipo de Recurso"
        .Cells(3, 3).Value = "Etiqueta del Material"
        .Cells(3, 4).Value = "Unidades de Asignación"
        .Cells(3, 5).Value = "Valor"
        .Cells(3, 6).Value = "Nombre de la Tarea"
    End With
    
    ' Rellenar datos de los recursos
    Dim Fila As Long
    Fila = 4
    
    For Each Recurso In ActiveProject.Resources
        If Not Recurso Is Nothing Then
            For Each Asignacion In Recurso.Assignments
                ExcelHoja.Cells(Fila, 1).Value = Recurso.Name
                
                ' Tipo de recurso
                Select Case Recurso.Type
                    Case pjResourceTypeMaterial
                        ExcelHoja.Cells(Fila, 2).Value = "Material"
                    Case pjResourceTypeWork
                        ExcelHoja.Cells(Fila, 2).Value = "Trabajo"
                    Case pjResourceTypeCost
                        ExcelHoja.Cells(Fila, 2).Value = "Costo"
                End Select
                
                ExcelHoja.Cells(Fila, 3).Value = Recurso.MaterialLabel
                ExcelHoja.Cells(Fila, 4).Value = Asignacion.Units
                ExcelHoja.Cells(Fila, 5).Value = Asignacion.Cost
                ExcelHoja.Cells(Fila, 6).Value = Asignacion.Task.Name
                Fila = Fila + 1
            Next Asignacion
        End If
    Next Recurso
    
    ' Crear el rango de datos
    Set ExcelRango = ExcelHoja.Range("A3:F" & Fila - 1)
    
    ' Establecer filtros en la fila de encabezados
    ExcelHoja.Range("A3:F3").AutoFilter
    
    ' Liberar objetos
    Set ExcelRango = Nothing
    Set ExcelHoja = Nothing
    Set ExcelLibro = Nothing
    Set ExcelApp = Nothing
    
    ' Mensaje de éxito
    MsgBox "La lista de recursos se ha creado correctamente en Excel.", vbInformation
End Sub


