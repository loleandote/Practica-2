Public Class frmPersonas
    Private p As Persona


    Private Sub BtnAbrirBD_Click(sender As Object, e As EventArgs) Handles btnAbrirBD.Click
        Dim pAux As Persona
        p = New Persona
        Try
            p.LeerTodasPersonas()
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
        For Each pAux In p.PerDAO.Personas
            lstPersonas.Items.Add(pAux.IDPersona)
        Next
        btnAbrirBD.Enabled = False
        btnAnadir.Enabled = True
    End Sub

    Private Sub LstPersonas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstPersonas.SelectedIndexChanged
        btnActualizar.Enabled = True
        btnEliminar.Enabled = True
        txtID.Enabled = False
        If Not lstPersonas.SelectedItem Is Nothing Then
            p = New Persona(lstPersonas.SelectedItem.ToString)
            Try
                p.LeerPersona()
            Catch ex As Exception
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End Try
            txtID.Text = p.IDPersona.ToString
            txtNombre.Text = p.Nombre.ToString
        End If
    End Sub

    Private Sub BtnAnadir_Click(sender As Object, e As EventArgs) Handles btnAnadir.Click
        If txtID.Text <> String.Empty And txtNombre.Text <> String.Empty Then
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
            lstPersonas.Items.Add(p.IDPersona)
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
        If Not p Is Nothing Then
            If MessageBox.Show("¿Estás seguro que quieres borrar " & p.IDPersona & "?", "Por favor, confirma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    If p.BorrarPersona() <> 1 Then
                        MessageBox.Show("DELETE return <> 1", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End Try
                lstPersonas.Items.Remove(p.IDPersona)
            End If
            btnLimpiar.PerformClick()
        End If

    End Sub

    Private Sub BtnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        txtID.Text = String.Empty
        txtNombre.Text = String.Empty
        btnActualizar.Enabled = False
        btnEliminar.Enabled = False
        txtID.Enabled = True
    End Sub


End Class