Imports System.Data.SqlClient

Public Class Form1

#Region "Load"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.ElementTableAdapter.Fill(Me.FormulesDataSet.element)
        Me.ComposicionsTableAdapter.Fill(Me.FormulesDataSet.composicions)
        Me.FormulesTableAdapter.Fill(Me.FormulesDataSet.formules)

    End Sub

    Private Sub FormulesBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)
        Me.Validate()
        Me.FormulesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.FormulesDataSet)
    End Sub

    Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        If Not (Asc(e.KeyChar) = 8) Then
            Dim allowedChars As String = "sn"
            If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub TextBox_KeyPress_numbers(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress, TextBox19.KeyPress
        If Not (Asc(e.KeyChar) = 8) Then
            Dim allowedChars As String = "1234567890"
            If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub
#End Region

#Region "Fórmules"
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Insertar
        If RadioButton1.Checked = True Then
            TextBox4.Text = "S"
        Else
            TextBox4.Text = "N"
        End If
        If TextBox2.Text = "" Then
            MsgBox("Has d'escriure el nom de la nova fórmula")
        Else
            Me.FormulesTableAdapter.Insert(TextBox2.Text, DateTimePicker1.Value, TextBox4.Text, 0)
            Me.FormulesTableAdapter.Fill(Me.FormulesDataSet.formules)
            TextBox2.Text = ""
            RadioButton1.Checked = True
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'Cercar
        If ComboBox1.Text = "" Then
            MsgBox("Has d'introduir el nom de la fórmula")
        Else
            Me.FormulesTableAdapter.GetDataBySelectFormules(ComboBox1.Text)
            Me.FormulesTableAdapter.FillBySelectFormules(Me.FormulesDataSet.formules, ComboBox1.Text)
        End If
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from formules", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox1.DataSource = table
        ComboBox1.DisplayMember = "nom"
        ComboBox1.ValueMember = "codi_f"
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        'Eliminar
        If ComboBox2.Text = "" Then
            MsgBox("Has d'introduir el nom de la fórmula")
        Else
            Me.FormulesTableAdapter.DeleteFormules(ComboBox2.Text)
            Me.FormulesTableAdapter.Fill(Me.FormulesDataSet.formules)
        End If
    End Sub

    Private Sub ComboBox2_Click(sender As Object, e As EventArgs) Handles ComboBox2.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from formules", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox2.DataSource = table
        ComboBox2.DisplayMember = "nom"
        ComboBox2.ValueMember = "codi_f"
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        'Editar
        If ComboBox3.Text = "" Then
            MsgBox("Has d'introduir el codi de la fórmula")
        Else
            If TextBox11.Text = "" Then
                MsgBox("Has d'introduir un nou nom per la fórmula")
            Else
                If RadioButton3.Checked = True Then
                    TextBox12.Text = "S"
                Else
                    TextBox12.Text = "N"
                End If
                Me.FormulesTableAdapter.UpdateFormules(TextBox11.Text, DateTimePicker3.Text, TextBox12.Text, ComboBox3.Text)
                Me.FormulesTableAdapter.Fill(Me.FormulesDataSet.formules)
            End If
        End If
    End Sub

    Private Sub ComboBox3_Click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from formules", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox3.DataSource = table
        ComboBox3.DisplayMember = "codi_f"
        ComboBox3.ValueMember = "codi_f"
    End Sub

#End Region

#Region "Composicions"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Insertar
        If ComboBox7.Text = "" Then
            MsgBox("Has d'introduir un codi de fómula")
        Else
            If ComboBox8.Text = "" Then
                MsgBox("Has d'introduir un codi d'element")
            Else
                If TextBox19.Text = "" Then
                    MsgBox("Has d'escriure una quantitat de l'element indicat anteriorment")
                Else
                    Me.ComposicionsTableAdapter.Insert(ComboBox7.Text, ComboBox8.Text, TextBox19.Text)
                    Me.ComposicionsTableAdapter.UpdateQuantitat_totalFormules(TextBox19.Text, ComboBox7.Text)
                    Me.ComposicionsTableAdapter.Fill(Me.FormulesDataSet.composicions)
                    Me.FormulesTableAdapter.Fill(Me.FormulesDataSet.formules)
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox7_Click(sender As Object, e As EventArgs) Handles ComboBox7.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from formules", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox7.DataSource = table
        ComboBox7.DisplayMember = "codi_f"
        ComboBox7.ValueMember = "codi_f"
    End Sub

    Private Sub ComboBox8_Click(sender As Object, e As EventArgs) Handles ComboBox8.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from element", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox8.DataSource = table
        ComboBox8.DisplayMember = "codi_e"
        ComboBox8.ValueMember = "codi_e"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Cercar
        If ComboBox9.Text = "" Then
            MsgBox("Has d'introduir un codi de composicio")
        Else
            Me.ComposicionsTableAdapter.GetDataBySelectComposicions(ComboBox9.Text)
            Me.ComposicionsTableAdapter.FillBySelectComposicions(Me.FormulesDataSet.composicions, ComboBox9.Text)
        End If
    End Sub

    Private Sub ComboBox9_Click(sender As Object, e As EventArgs) Handles ComboBox9.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from composicions", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox9.DataSource = table
        ComboBox9.DisplayMember = "codi_c"
        ComboBox9.ValueMember = "codi_c"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Eliminar
        If ComboBox10.Text = "" Then
            MsgBox("Has d'indicar un codi de composició")
        Else
            If ComboBox11.Text = "" Then
                MsgBox("Has d'indicar un codi de fórmula")
            Else
                Me.ComposicionsTableAdapter.UpdateDeleteComposicio(ComboBox11.Text, ComboBox10.Text)
                Me.ComposicionsTableAdapter.DeleteComposicions(ComboBox10.Text, ComboBox11.Text)
                Me.ComposicionsTableAdapter.Fill(Me.FormulesDataSet.composicions)
                Me.FormulesTableAdapter.Fill(Me.FormulesDataSet.formules)
            End If
        End If
    End Sub

    Private Sub ComboBox10_Click(sender As Object, e As EventArgs) Handles ComboBox10.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from composicions", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox10.DataSource = table
        ComboBox10.DisplayMember = "codi_c"
        ComboBox10.ValueMember = "codi_c"
    End Sub

    Private Sub ComboBox11_Click(sender As Object, e As EventArgs) Handles ComboBox11.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from formules", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox11.DataSource = table
        ComboBox11.DisplayMember = "codi_f"
        ComboBox11.ValueMember = "codi_f"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Editar
        If ComboBox12.Text = "" Then
            MsgBox("Has d'indicar un codi de composició")
        Else
            If ComboBox13.Text = "" Then
                MsgBox("Has d'indicar un codi de fórmula")
            Else
                If ComboBox14.Text = "" Then
                    MsgBox("Has d'indicar un codi d'element")
                Else
                    If TextBox6.Text = "" Then
                        MsgBox("Has d'indicar una nova quantitat de l'element indicat anteriorment")
                    Else
                        Me.ComposicionsTableAdapter.UpdateDeleteComposicio(ComboBox13.Text, ComboBox12.Text)
                        Me.ComposicionsTableAdapter.UpdateComposicions(ComboBox13.Text, ComboBox14.Text, TextBox6.Text, ComboBox12.Text)
                        Me.ComposicionsTableAdapter.UpdateQuantitat_totalFormules(TextBox6.Text, ComboBox13.Text)
                        Me.ComposicionsTableAdapter.Fill(Me.FormulesDataSet.composicions)
                        Me.FormulesTableAdapter.Fill(Me.FormulesDataSet.formules)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox12_Click(sender As Object, e As EventArgs) Handles ComboBox12.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from composicions", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox12.DataSource = table
        ComboBox12.DisplayMember = "codi_c"
        ComboBox12.ValueMember = "codi_c"
    End Sub

    Private Sub ComboBox13_Click(sender As Object, e As EventArgs) Handles ComboBox13.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from formules", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox13.DataSource = table
        ComboBox13.DisplayMember = "codi_f"
        ComboBox13.ValueMember = "codi_f"
    End Sub

    Private Sub ComboBox14_Click(sender As Object, e As EventArgs) Handles ComboBox14.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from element", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox14.DataSource = table
        ComboBox14.DisplayMember = "codi_e"
        ComboBox14.ValueMember = "codi_e"
    End Sub

#End Region

#Region "Elements"

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        'Insertar
        If TextBox17.Text = "" Then
            MsgBox("Has d'escriure el nom del nou element")
        Else
            Me.ElementTableAdapter.Insert(TextBox17.Text, DateTimePicker4.Value.Date)
            Me.ElementTableAdapter.Fill(Me.FormulesDataSet.element)
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'Cercar
        If ComboBox4.Text = "" Then
            MsgBox("Has d'introduir un nom d'element")
        Else
            Me.ElementTableAdapter.GetDataBySelectElement(ComboBox4.Text)
            Me.ElementTableAdapter.FillBySelectElement(Me.FormulesDataSet.element, ComboBox4.Text)
        End If
    End Sub

    Private Sub ComboBox4_Click(sender As Object, e As EventArgs) Handles ComboBox4.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from element", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox4.DataSource = table
        ComboBox4.DisplayMember = "nom"
        ComboBox4.ValueMember = "codi_e"
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        'Eliminar
        If ComboBox5.Text = "" Then
            MsgBox("Has d'introduir un nom d'element")
        Else
            Me.ElementTableAdapter.DeleteElement(ComboBox5.Text)
            Me.ElementTableAdapter.Fill(Me.FormulesDataSet.element)
        End If
    End Sub

    Private Sub ComboBox5_Click(sender As Object, e As EventArgs) Handles ComboBox5.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from element", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox5.DataSource = table
        ComboBox5.DisplayMember = "nom"
        ComboBox5.ValueMember = "codi_e"
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'Editar
        If ComboBox6.Text = "" Then
            MsgBox("Has d'introduir un codi d'element")
        Else
            If TextBox9.Text = "" Then
                MsgBox("Has d'escriure un nou nom per l'element")
            Else
                Me.ElementTableAdapter.UpdateElement(TextBox9.Text, DateTimePicker2.Text, ComboBox6.Text)
                Me.ElementTableAdapter.Fill(Me.FormulesDataSet.element)
            End If
        End If
    End Sub

    Private Sub ComboBox6_Click(sender As Object, e As EventArgs) Handles ComboBox6.Click
        Dim connection As New SqlConnection("SERVER= DESKTOP-9KSV0B9; Database = formules; integrated security=true")
        Dim command As New SqlCommand("select * from element", connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        ComboBox6.DataSource = table
        ComboBox6.DisplayMember = "codi_e"
        ComboBox6.ValueMember = "codi_e"
    End Sub

#End Region

#Region "Extres"

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.FormulesTableAdapter.GetDataBySelectFormulesInactives()
        Me.FormulesTableAdapter.FillBySelectFormulesInactives(Me.FormulesDataSet.formules)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.FormulesTableAdapter.GetDataBySelectDataFormules(DateTimePicker5.Text)
        Me.FormulesTableAdapter.FillBySelectDataFormules(Me.FormulesDataSet.formules, DateTimePicker5.Text)
    End Sub

#End Region

End Class
