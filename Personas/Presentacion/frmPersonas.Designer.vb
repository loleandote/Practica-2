<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPersonas
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lstPersonas = New System.Windows.Forms.ListBox()
        Me.btnAbrirBD = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me.btnAnadir = New System.Windows.Forms.Button()
        Me.btnActualizar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstPersonas
        '
        Me.lstPersonas.FormattingEnabled = True
        Me.lstPersonas.ItemHeight = 16
        Me.lstPersonas.Location = New System.Drawing.Point(58, 76)
        Me.lstPersonas.Name = "lstPersonas"
        Me.lstPersonas.Size = New System.Drawing.Size(120, 84)
        Me.lstPersonas.TabIndex = 0
        '
        'btnAbrirBD
        '
        Me.btnAbrirBD.Location = New System.Drawing.Point(58, 212)
        Me.btnAbrirBD.Name = "btnAbrirBD"
        Me.btnAbrirBD.Size = New System.Drawing.Size(75, 23)
        Me.btnAbrirBD.TabIndex = 1
        Me.btnAbrirBD.Text = "Conectar!"
        Me.btnAbrirBD.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(308, 76)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(25, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "ID:"
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(311, 110)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(100, 22)
        Me.txtID.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(311, 154)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Nombre:"
        '
        'txtNombre
        '
        Me.txtNombre.Location = New System.Drawing.Point(314, 190)
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.Size = New System.Drawing.Size(100, 22)
        Me.txtNombre.TabIndex = 5
        '
        'btnAnadir
        '
        Me.btnAnadir.Location = New System.Drawing.Point(314, 233)
        Me.btnAnadir.Name = "btnAnadir"
        Me.btnAnadir.Size = New System.Drawing.Size(75, 23)
        Me.btnAnadir.TabIndex = 6
        Me.btnAnadir.Text = "Añadir"
        Me.btnAnadir.UseVisualStyleBackColor = True
        '
        'btnActualizar
        '
        Me.btnActualizar.Location = New System.Drawing.Point(419, 233)
        Me.btnActualizar.Name = "btnActualizar"
        Me.btnActualizar.Size = New System.Drawing.Size(82, 23)
        Me.btnActualizar.TabIndex = 7
        Me.btnActualizar.Text = "Actualizar"
        Me.btnActualizar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(419, 278)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(75, 23)
        Me.btnLimpiar.TabIndex = 9
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(314, 278)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(75, 23)
        Me.btnEliminar.TabIndex = 8
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'frmPersonas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(621, 392)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnActualizar)
        Me.Controls.Add(Me.btnAnadir)
        Me.Controls.Add(Me.txtNombre)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnAbrirBD)
        Me.Controls.Add(Me.lstPersonas)
        Me.Margin = New System.Windows.Forms.Padding(1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPersonas"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Personas"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstPersonas As ListBox
    Friend WithEvents btnAbrirBD As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtID As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtNombre As TextBox
    Friend WithEvents btnAnadir As Button
    Friend WithEvents btnActualizar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
End Class
