Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel
Public Class Form1
    Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
    Dim Libro = ExcelApp.Workbooks.Open("C:\Users\USER-XPS\Desktop\Prueba1.xlsx")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Fila As Integer = 2
        Dim Columna As Integer = 1
        Dim RowCount = DataGridView1.Rows.Count - 2
        Dim ColumnCount = DataGridView1.Columns.Count - 1

        For nColumna As Integer = 0 To ColumnCount

            Libro.Worksheets("Hoja1").Cells(1, Columna) = DataGridView1.Columns(nColumna).HeaderText
            Libro.Worksheets("Hoja1").Cells(1, Columna).Font.Bold = True

            For nFila As Integer = 0 To RowCount
                Libro.Worksheets("Hoja1").Cells(Fila, Columna) = DataGridView1.Rows(nFila).Cells(nColumna).Value
                Fila = Fila + 1
            Next
            Columna = Columna + 1
            Fila = 2
        Next

        MsgBox("Los registros se exportaron satisfactoriamente")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Libro.Save()
        MsgBox("Los cambios se han guardado en " & Libro.Name)
        ExcelApp.Quit()
        Libro = Nothing
        ExcelApp = Nothing
        End
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Libro.saved() = False Then
            Dim Respuesta As MsgBoxResult = MsgBox("Desea guardar los cambios en el libro " & Libro.Name & vbExclamation + vbYesNo, "Microsoft Excel")

            Select Case Respuesta
                Case MsgBoxResult.Yes
                    Libro.Save()
                    MsgBox("Los cambios se han guardado en " & Libro.Name)
                    ExcelApp.Quit()
                    Libro = Nothing
                    ExcelApp = Nothing
                Case MsgBoxResult.No
                    Libro.saved() = True
                    ExcelApp.Quit()
                    Libro = Nothing
                    ExcelApp = Nothing
            End Select

        Else
            ExcelApp.Quit()
            Libro = Nothing
            ExcelApp = Nothing
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i As Integer = 1 To 100
            DataGridView1.Rows.Add(New String() {"Cliente " & i, "correo" & i & "@mail.com"})
        Next



    End Sub
End Class