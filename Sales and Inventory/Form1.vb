Imports System.Data.OleDb ' The OleDB Class
Public Class frmMain

    ' Set the connection string
    Dim conString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath() + "\data.accdb"
    Dim con As New OleDbConnection

    ' Create new datasets to handle data
    Dim ds As New DataSet
    Dim da As New OleDbDataAdapter
    Dim table As DataTable
    Dim row As DataRow
    Dim cmd As OleDbCommand

    Dim i As Integer = 0
    Dim sql

    Private Sub Populate()
        ds.Clear()
        sql = "SELECT * FROM items" ' This is your SQL SELECT Statement
        da = New OleDbDataAdapter(sql, con) ' Run the SQL SELECT STATEMENT (sql) using the connection we opened (con)
        da.Fill(ds) ' Fill the data set with the result set we got from the DataAdapter
        table = ds.Tables(0) ' For shortcuts, instead of writing ds.Tables(0).Rows we could use table.Rows
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Start a connection to the Access database
        con.ConnectionString = conString
        con.Open()

        ' Run the Populate sub we created above to populate the data set with the INITIAL entries
        Populate()

        ' Count the items in the data set. If no items are found, tell the user.
        If table.Rows.Count = 0 Then
            MessageBox.Show("There are no items yet. Click on Add to add a new item.", "No Items", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ' There are items. Show the first item as an initial view.
            txtID.Text = table.Rows(0).Item(0)
            txtName.Text = table.Rows(0).Item(1)
            txtDesc.Text = table.Rows(0).Item(2)
            numQty.Value = table.Rows(0).Item(3)
            txtPrice.Text = table.Rows(0).Item(4)
            i = 0 ' Set the current index to 0
        End If
    End Sub

    Private Sub handlerBtn_Tick(sender As Object, e As EventArgs) Handles handlerBtn.Tick
        Dim numRows = table.Rows.Count
        ' If there are no entries (or if there's only one entry), disable all nav buttons
        If numRows <= 1 Then
            btnFirst.Enabled = False
            btnPrevious.Enabled = False
            btnNext.Enabled = False
            btnLast.Enabled = False
        ElseIf i + 1 = 1 Then
            ' If we're on the first entry, disable the First and Previous buttons
            btnFirst.Enabled = False
            btnPrevious.Enabled = False
            btnNext.Enabled = True
            btnLast.Enabled = True
        ElseIf i + 1 = numRows Then
            ' If we're on the last entry, disable the Next and Last buttons
            btnFirst.Enabled = True
            btnPrevious.Enabled = True
            btnNext.Enabled = False
            btnLast.Enabled = False
        Else
            ' Enable everything
            btnFirst.Enabled = True
            btnPrevious.Enabled = True
            btnNext.Enabled = True
            btnLast.Enabled = True
        End If

    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        ' Display the first entry
        txtID.Text = table.Rows(0).Item(0)
        txtName.Text = table.Rows(0).Item(1)
        txtDesc.Text = table.Rows(0).Item(2)
        numQty.Value = table.Rows(0).Item(3)
        txtPrice.Text = table.Rows(0).Item(4)
        i = 0 ' Set the current index to 0
    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        If i <> 0 Then
            ' If we're not on the first entry, deduct 1 from the current index.
            i -= 1
        End If
        txtID.Text = table.Rows(i).Item(0)
        txtName.Text = table.Rows(i).Item(1)
        txtDesc.Text = table.Rows(i).Item(2)
        numQty.Value = table.Rows(i).Item(3)
        txtPrice.Text = table.Rows(i).Item(4)
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If i + 1 <> table.Rows.Count Then
            ' If we're not on the last entry, add 1 from the current index.
            i += 1
        End If
        txtID.Text = table.Rows(i).Item(0)
        txtName.Text = table.Rows(i).Item(1)
        txtDesc.Text = table.Rows(i).Item(2)
        numQty.Value = table.Rows(i).Item(3)
        txtPrice.Text = table.Rows(i).Item(4)
    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        i = table.Rows.Count - 1 ' Set index to last entry
        txtID.Text = table.Rows(i).Item(0)
        txtName.Text = table.Rows(i).Item(1)
        txtDesc.Text = table.Rows(i).Item(2)
        numQty.Value = table.Rows(i).Item(3)
        txtPrice.Text = table.Rows(i).Item(4)
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If btnAdd.Text = "Add" Then
            ' Make me a Cancel button, then enable the Save button. Disable every other button too. Temporarily stop the button nav handler
            handlerBtn.Stop()
            btnAdd.Text = "Cancel"
            btnSave.Enabled = True
            btnEdit.Enabled = False
            btnDelete.Enabled = False
            btnFirst.Enabled = False
            btnPrevious.Enabled = False
            btnNext.Enabled = False
            btnLast.Enabled = False

            ' Clear the form for new entry
            txtID.Clear()
            txtName.Clear()
            txtDesc.Clear()
            txtPrice.Clear()
            numQty.Value = 1
            txtID.Text = "Automatically assigned"
            txtID.Enabled = False

            ' Make things editable
            txtName.ReadOnly = False
            txtDesc.ReadOnly = False
            txtPrice.ReadOnly = False
            numQty.ReadOnly = False
        Else
            ' Return to normal. We don't need to re-enable the nav buttons, the handler handles it automatically.
            handlerBtn.Start()
            btnAdd.Text = "Add"
            btnSave.Enabled = False
            btnEdit.Enabled = True
            btnDelete.Enabled = True
            txtID.Text = table.Rows(i).Item(0)
            txtName.Text = table.Rows(i).Item(1)
            txtDesc.Text = table.Rows(i).Item(2)
            numQty.Value = table.Rows(i).Item(3)
            txtPrice.Text = table.Rows(i).Item(4)

            ' Make things non-editable
            txtName.ReadOnly = True
            txtDesc.ReadOnly = True
            txtPrice.ReadOnly = True
            numQty.ReadOnly = True
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If btnAdd.Text = "Cancel" Then
            ' We have to save an ADDITIONAL entry.
            ' Verify if price is numeric
            If IsNumeric(txtPrice.Text) = False Then
                MessageBox.Show("Price must be a number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            sql = "INSERT INTO items ([ItemName], [ItemDescription], [ItemQuantity], [Price]) VALUES ('" + txtName.Text + "', '" + txtDesc.Text + "', " + numQty.Value.ToString + ", " + txtPrice.Text + ")"

            ' Run Query
            cmd = New OleDbCommand(sql, con)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Entry saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Return to normal. We don't need to re-enable the nav buttons, the handler handles it automatically. Refresh the data set, then set i to the latest entry.
            Populate()
            i = table.Rows.Count - 1
            handlerBtn.Start()
            btnAdd.Text = "Add"
            btnSave.Enabled = False
            btnEdit.Enabled = True
            btnDelete.Enabled = True
            txtID.Text = table.Rows(i).Item(0)
            txtName.Text = table.Rows(i).Item(1)
            txtDesc.Text = table.Rows(i).Item(2)
            numQty.Value = table.Rows(i).Item(3)
            txtPrice.Text = table.Rows(i).Item(4)
        Else
            ' We have to save an UPDATE to an existing entry.
            If IsNumeric(txtPrice.Text) = False Then
                MessageBox.Show("Price must be a number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            sql = "UPDATE items SET ItemName = '" + txtName.Text + "', ItemDescription = '" + txtDesc.Text + "', ItemQuantity = '" + numQty.Value.ToString + "', Price = '" + txtPrice.Text + "' WHERE ID = " + txtID.Text

            ' Run Query
            cmd = New OleDbCommand(sql, con)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Entry updated.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ' Return to normal. We don't need to re-enable the nav buttons, the handler handles it automatically. Refresh the data set.
            Populate()
            handlerBtn.Start()
            btnEdit.Text = "Edit"
            btnSave.Enabled = False
            btnAdd.Enabled = True
            btnDelete.Enabled = True
            txtID.Text = table.Rows(i).Item(0)
            txtName.Text = table.Rows(i).Item(1)
            txtDesc.Text = table.Rows(i).Item(2)
            numQty.Value = table.Rows(i).Item(3)
            txtPrice.Text = table.Rows(i).Item(4)
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If btnEdit.Text = "Edit" Then
            ' Make me a Cancel button, then enable the Save button. Disable every other button too. Temporarily stop the button nav handler
            handlerBtn.Stop()
            btnEdit.Text = "Cancel"
            btnSave.Enabled = True
            btnAdd.Enabled = False
            btnDelete.Enabled = False
            btnFirst.Enabled = False
            btnPrevious.Enabled = False
            btnNext.Enabled = False
            btnLast.Enabled = False

            ' Make things editable
            txtName.ReadOnly = False
            txtDesc.ReadOnly = False
            txtPrice.ReadOnly = False
            numQty.ReadOnly = False
        Else
            ' Return to normal. We don't need to re-enable the nav buttons, the handler handles it automatically.
            handlerBtn.Start()
            btnEdit.Text = "Edit"
            btnSave.Enabled = False
            btnAdd.Enabled = True
            btnDelete.Enabled = True
            txtID.Text = table.Rows(i).Item(0)
            txtName.Text = table.Rows(i).Item(1)
            txtDesc.Text = table.Rows(i).Item(2)
            numQty.Value = table.Rows(i).Item(3)
            txtPrice.Text = table.Rows(i).Item(4)

            ' Make things non-editable
            txtName.ReadOnly = True
            txtDesc.ReadOnly = True
            txtPrice.ReadOnly = True
            numQty.ReadOnly = True
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        ' Confirm
        Dim ask = MessageBox.Show("Are you sure you want to delete this entry?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ask = DialogResult.No Then
            Exit Sub
        End If

        ' Delete entry
        sql = "DELETE FROM items WHERE ID = " + txtID.Text

        ' Run Query
        cmd = New OleDbCommand(sql, con)
        cmd.ExecuteNonQuery()
        MessageBox.Show("Entry deleted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' Repopulate
        Populate()
        txtID.Text = table.Rows(0).Item(0)
        txtName.Text = table.Rows(0).Item(1)
        txtDesc.Text = table.Rows(0).Item(2)
        numQty.Value = table.Rows(0).Item(3)
        txtPrice.Text = table.Rows(0).Item(4)
    End Sub
End Class
