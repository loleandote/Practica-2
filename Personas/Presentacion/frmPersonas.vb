Public Class frmPersonas
    Private p As Persona
    'Private lstPersonas As ListBox
    'Private btnAbrirBD As Button
    'Private btnAnadir As Button
    'Private btnActualizar, btnEliminar, btnLimpiar As Button
    'Private txtID, txtNombre As TextBox



    Public Sub New()

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    Private Sub BtnAbrirBD_Click(sender As Object, e As EventArgs) Handles btnAbrirBD.Click
        Dim pAux As Persona
        Me.p = New Persona
        Try
            Me.p.LeerTodasPersonas()
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
        For Each pAux In Me.p.PerDAO.Personas
            Me.lstPersonas.Items.Add(pAux.IDPersona)
        Next
        btnAbrirBD.Enabled = False
        btnAnadir.Enabled = True
    End Sub

    Private Sub LstPersonas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstPersonas.SelectedIndexChanged
        Me.btnActualizar.Enabled = True
        Me.btnEliminar.Enabled = True
        Me.txtID.Enabled = False
        If Not Me.lstPersonas.SelectedItem Is Nothing Then
            Me.p = New Persona(Me.lstPersonas.SelectedItem.ToString)
            Try
                Me.p.LeerPersona()
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.txtID.Text = Me.p.IDPersona.ToString
            Me.txtNombre.Text = Me.p.Nombre.ToString
        End If
    End Sub

    Private Sub BtnAnadir_Click(sender As Object, e As EventArgs) Handles btnAnadir.Click
        If Me.txtID.Text <> String.Empty And Me.txtNombre.Text <> String.Empty Then
            p = New Persona
            p.IDPersona = txtID.Text
            p.Nombre = txtNombre.Text

            Try
                If p.InsertarPersona() <> 1 Then
                    MessageBox.Show("INSERT return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            Me.lstPersonas.Items.Add(p.IDPersona)
        End If
    End Sub



    Private Sub BtnActualizar_Click(sender As Object, e As EventArgs) Handles btnActualizar.Click

        If Not p Is Nothing Then
            p.Nombre = txtNombre.Text
            Try
                If p.ActualizarPersona() <> 1 Then
                    MessageBox.Show("UPDATE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            MessageBox.Show(p.Nombre & " actualizado correctamente!")
        End If

    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If Not Me.p Is Nothing Then
            If MessageBox.Show("¿Estás seguro que quieres borrar " & Me.p.IDPersona & "?", "Por favor, confirma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    If Me.p.BorrarPersona() <> 1 Then
                        MessageBox.Show("DELETE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End Try
                Me.lstPersonas.Items.Remove(p.IDPersona)
            End If
            Me.btnLimpiar.PerformClick()
        End If

    End Sub

    Private Sub BtnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Me.txtID.Text = String.Empty
        Me.txtNombre.Text = String.Empty
        Me.btnActualizar.Enabled = False
        Me.btnEliminar.Enabled = False
        Me.txtID.Enabled = True
    End Sub


End Class