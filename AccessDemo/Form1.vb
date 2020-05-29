Imports AccessDemo.Classes
Imports AccessDemo.HelperClasses

Public Class Form1
    Private customerBindingSource As New BindingSource
    Private Sub CopyDatabaseButton_Click(sender As Object, e As EventArgs) _
        Handles CopyDatabaseButton.Click


        DataGridView1.DataSource = Nothing

        FileHelper.CopyDatabase(True)

        If Not FileHelper.IsSuccessFul Then
            If FileHelper.DatbaseInuse Then
                MessageBox.Show("Please close the database and try again")
            Else
                MessageBox.Show(FileHelper.LastExceptionMessage)
            End If
        Else

            If CreateTableCheckBox.Checked Then
                If Not DataOperations.CreateTable() Then
                    MessageBox.Show($"{DataOperations.LastExceptionMessage}")
                End If
            End If

            customerBindingSource.DataSource = DataOperations.LoadCustomers()
            DataGridView1.DataSource = customerBindingSource
        End If

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        If Not FileHelper.Exists Then
            FileHelper.CopyDatabase(True)
        End If

        customerBindingSource.DataSource = DataOperations.LoadCustomers()
        DataGridView1.DataSource = customerBindingSource
    End Sub
End Class