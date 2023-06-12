Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports HtmlAgilityPack
Imports OfficeOpenXml
Imports System.Text


Public Class Form1

    Dim connection As New SqlConnection("Data Source=192.168.10.114\APPSERVER;Initial Catalog=Application;User ID=sa;Password=XXXXXXX")

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.No Then
            e.Cancel = True ' Cancel the closing event
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'System.Diagnostics.Process.Start("http://ad00/infor/working.aspx")
        DataGridView2.VirtualMode = True
        AddHandler DataGridView2.CellValueNeeded, AddressOf DataGridView2_CellValueNeeded
        PopulateDataTable()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'ApplicationDataSet1.EmployeeCheckIn' table. You can move, or remove it, as needed.
        Me.EmployeeCheckInTableAdapter.Fill(Me.ApplicationDataSet1.EmployeeCheckIn)
        FilterData("")

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        ' Create a new Excel package
        Using package As New ExcelPackage()
            ' Add a new worksheet to the Excel package
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Data")

            ' Get the filtered rows from the DataGridView and export to Excel
            Dim filteredRows As DataGridViewRowCollection = DataGridView1.Rows

            ' Set the column headers in Excel
            For columnIndex As Integer = 0 To DataGridView1.Columns.Count - 1
                worksheet.Cells(1, columnIndex + 1).Value = DataGridView1.Columns(columnIndex).HeaderText
            Next

            ' Set the row data in Excel
            For rowIndex As Integer = 0 To filteredRows.Count - 1
                For columnIndex As Integer = 0 To DataGridView1.Columns.Count - 1
                    If columnIndex = 4 Then ' Column 5 (zero-based index)
                        Dim originalValue As Date = CDate(filteredRows(rowIndex).Cells(columnIndex).Value)
                        Dim formattedValue As String = originalValue.ToString("dd/MM/yy HH:mm:ss")
                        worksheet.Cells(rowIndex + 2, columnIndex + 1).Value = formattedValue
                    Else
                        worksheet.Cells(rowIndex + 2, columnIndex + 1).Value = filteredRows(rowIndex).Cells(columnIndex).Value
                    End If
                Next
            Next

            ' Save the Excel package to a file
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            saveFileDialog.FileName = "SQLData " & TextBox1.Text & ".xlsx"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath As String = saveFileDialog.FileName
                Dim file As New FileInfo(filePath)
                package.SaveAs(file)
                MessageBox.Show("Data exported to Excel successfully.")
            End If
        End Using
    End Sub

    Public Sub FilterData(valueToSearch As String)


        Dim searchQuery As String = "SELECT * FROM EmployeeCheckIn WHERE Date between '" & valueToSearch & " 00:00:00' and '" & valueToSearch & " 23:59:59'"

        Dim command As New SqlCommand(searchQuery, connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()

        adapter.Fill(table)
        DataGridView1.DataSource = table

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        FilterData(TextBox1.Text)
    End Sub

    Private dataTable As New DataTable()

    Private Sub PopulateDataTable()
        ' Specify the URL of the webpage to retrieve data from
        Dim url As String = "http://ad00/infor/working.aspx"

        ' Create a WebClient instance to download the HTML content
        Dim client As New WebClient()
        client.Encoding = Encoding.UTF8
        Dim htmlContent As String = client.DownloadString(url)

        ' Load the HTML content into an HtmlDocument
        Dim htmlDocument As New HtmlAgilityPack.HtmlDocument()
        htmlDocument.LoadHtml(htmlContent)

        ' Select all <tr> elements within the table
        Dim trNodes As HtmlNodeCollection = htmlDocument.DocumentNode.SelectNodes("//table//tr")

        ' Add columns to the DataTable with desired column names
        dataTable.Columns.Add("No")
        dataTable.Columns.Add("Date")
        dataTable.Columns.Add("Parent_code")
        dataTable.Columns.Add("Parent_desc")
        dataTable.Columns.Add("Child_code")
        dataTable.Columns.Add("Child_desc")
        dataTable.Columns.Add("SSN")
        dataTable.Columns.Add("Title")
        dataTable.Columns.Add("Name")
        dataTable.Columns.Add("Surname")
        dataTable.Columns.Add("Phone")
        dataTable.Columns.Add("SF_code")
        dataTable.Columns.Add("SF_desc")
        dataTable.Columns.Add("Stamp_1")
        dataTable.Columns.Add("Stamp_2")
        dataTable.Columns.Add("Stamp_3")
        dataTable.Columns.Add("Stamp_4")

        ' Iterate over the <tr> elements and extract the cell values
        For rowIndex As Integer = 1 To trNodes.Count - 1 ' Start from index 1 to skip the header row
            ' Select all <td> elements within the current <tr>
            Dim tdNodes As HtmlNodeCollection = trNodes(rowIndex).SelectNodes("td")

            ' Check if the number of <td> elements matches the number of columns
            If tdNodes IsNot Nothing AndAlso tdNodes.Count = dataTable.Columns.Count Then
                ' Create a new row in the DataTable
                Dim row As DataRow = dataTable.NewRow()

                ' Extract the cell values and add them to the DataTable
                For columnIndex As Integer = 0 To tdNodes.Count - 1
                    row(columnIndex) = tdNodes(columnIndex).InnerText
                Next

                ' Add the row to the DataTable
                dataTable.Rows.Add(row)
            End If
        Next

        ' Set the DataTable as the data source for the DataGridView
        DataGridView2.DataSource = dataTable
    End Sub

    Private Sub DataGridView2_CellValueNeeded(sender As Object, e As DataGridViewCellValueEventArgs)
        ' Check if the requested column index is valid
        If e.ColumnIndex >= 0 AndAlso e.ColumnIndex < dataTable.Columns.Count Then
            ' Check if the requested row index is valid
            If e.RowIndex >= 0 AndAlso e.RowIndex < dataTable.Rows.Count Then
                ' Retrieve the value for the specified cell in the DataTable
                e.Value = dataTable.Rows(e.RowIndex)(e.ColumnIndex)
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        ' Set the license context for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        ' Create a new Excel package
        Using package As New ExcelPackage()
            ' Add a new worksheet to the Excel package
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Data")

            ' Get the filtered rows from the DataGridView and export to Excel
            Dim filteredRows As DataGridViewRowCollection = DataGridView2.Rows

            ' Set the column headers in Excel
            For columnIndex As Integer = 0 To DataGridView2.Columns.Count - 1
                worksheet.Cells(1, columnIndex + 1).Value = DataGridView2.Columns(columnIndex).HeaderText
            Next

            ' Set the row data in Excel
            For rowIndex As Integer = 0 To filteredRows.Count - 1
                For columnIndex As Integer = 0 To DataGridView2.Columns.Count - 1
                    worksheet.Cells(rowIndex + 2, columnIndex + 1).Value = filteredRows(rowIndex).Cells(columnIndex).Value
                Next
            Next

            ' Save the Excel package to a file
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            Dim dt As Date = Date.Today
            saveFileDialog.FileName = "BplusData " & dt.ToString("dd-MM-yyyy") & ".xlsx"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath As String = saveFileDialog.FileName
                Dim file As New FileInfo(filePath)
                package.SaveAs(file)
                MessageBox.Show("Data exported to Excel successfully.")
            End If
        End Using

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Load the Excel file using EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Using package As New ExcelPackage(New FileInfo("D:\Report ScanEmployee\SCAN2023 QR-Bplus.xlsx"))
            ' Get the desired worksheet by name
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets("BPLUS4001")

            ' Check if the worksheet exists
            If worksheet IsNot Nothing Then
                ' Clear the existing data in the worksheet
                worksheet.Cells.Clear()

                ' Set the column headers in the first row of the worksheet
                For columnIndex As Integer = 0 To DataGridView2.Columns.Count - 1
                    worksheet.Cells(1, columnIndex + 1).Value = DataGridView2.Columns(columnIndex).HeaderText
                Next

                ' Get the DataTable underlying the DataGridView
                Dim dataTable As DataTable = DirectCast(DataGridView2.DataSource, DataTable)

                ' Copy the data from the DataTable to the worksheet
                For rowIndex As Integer = 0 To dataTable.Rows.Count - 1
                    For columnIndex As Integer = 0 To dataTable.Columns.Count - 1
                        worksheet.Cells(rowIndex + 2, columnIndex + 1).Value = dataTable.Rows(rowIndex)(columnIndex)
                    Next
                Next

                ' Save the changes in the Excel file
                package.Save()

            Else
                ' Handle the case when the worksheet is not found
                MessageBox.Show("Worksheet 'BPLUS4001' not found.")
            End If
        End Using

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Using package As New ExcelPackage(New FileInfo("D:\Report ScanEmployee\SCAN2023 QR-Bplus.xlsx"))
            ' Get the desired worksheet by name
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets("SQL(QR)3001")

            ' Check if the worksheet exists
            If worksheet IsNot Nothing Then
                ' Clear the existing data in the worksheet
                worksheet.Cells.Clear()

                ' Set the column headers in the first row of the worksheet
                For columnIndex As Integer = 0 To DataGridView1.Columns.Count - 1
                    worksheet.Cells(1, columnIndex + 1).Value = DataGridView1.Columns(columnIndex).HeaderText
                Next

                ' Get the DataTable underlying the DataGridView
                Dim dataTable As DataTable = DirectCast(DataGridView1.DataSource, DataTable)

                ' Copy the data from the DataTable to the worksheet
                For rowIndex As Integer = 0 To dataTable.Rows.Count - 1
                    For columnIndex As Integer = 0 To dataTable.Columns.Count - 1
                        worksheet.Cells(rowIndex + 2, columnIndex + 1).Value = dataTable.Rows(rowIndex)(columnIndex)
                    Next
                Next

                ' Save the changes in the Excel file
                package.Save()
                MessageBox.Show("Worksheet 'BPLUS4001 & SQL(QR)3001' Success")
            Else
                ' Handle the case when the worksheet is not found
                MessageBox.Show("Worksheet 'SQL(QR)3001' not found.")
            End If
        End Using
    End Sub

End Class
