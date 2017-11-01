Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel

Public Class Form2
    Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
    Dim Libro = ExcelApp.Workbooks.Open("C:\Users\USER-XPS\Desktop\Prueba1.xlsx")
    Dim Hoja1 = Libro.Worksheets("Hoja1")
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim Fila As Integer
        Dim Final As Integer
        Final = ExcelApp.Cells(1, 1).End(XlDirection.xlDown).Row
        Final = Final + 1
        For Fila = 2 To Final
            If Hoja1.Cells(Fila, 1).Text = Me.ComboBox1.Text Then
                Me.TextBox1.Text = Hoja1.Cells(Fila, 2).Text
                Exit For
            End If
        Next
    End Sub

    Private Sub ComboBox1_Enter(sender As Object, e As EventArgs) Handles ComboBox1.Enter
        Dim Fila As Integer
        Dim Final As Integer
        Dim Lista As String
        Me.ComboBox1.Items.Clear()

        Final = ExcelApp.Cells(1, 1).End(XlDirection.xlDown).Row
        Final = Final + 1

        For Fila = 2 To Final
            Lista = Hoja1.cells(Fila, 1).text
            Me.ComboBox1.Items.Add(Lista)
        Next


    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        If Me.ComboBox1.Text = "" Then
            Me.TextBox1.Text = ""
        End If
    End Sub

    Private Sub Form2_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ExcelApp.Quit()
        Libro = Nothing
        ExcelApp = Nothing
    End Sub
End Class