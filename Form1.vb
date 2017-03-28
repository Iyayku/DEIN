Imports System.Data
Imports System.Data.OleDb
Public Class Form1
    Private cn As New OleDbConnection
    Private daNomZona As OleDbDataAdapter
    Private daZonas As OleDbDataAdapter
    Private daEstaciones As OleDbDataAdapter
    Private ds As New DataSet
    Private nombreBBDD = Application.StartupPath & "\marzo_2017.accdb"
    Private WithEvents bs As New BindingSource
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ConfigurarAccesoADatos()
        EnlazarADatos()
        bs.MoveLast()
        bs.MoveFirst()
    End Sub
    Sub ConfigurarAccesoADatos()

        cn.ConnectionString = "provider= Microsoft.ace.oledb.12.0;" & "Data source = " & nombreBBDD
        cn.Open()
        daZonas = New OleDbDataAdapter("select * from N_ZONAS", cn)
        daZonas.Fill(ds, "misZonas")
        daEstaciones = New OleDbDataAdapter("select * from N_ESTACIONES", cn)
        daEstaciones.Fill(ds, "misEstaciones")
        
        daNomZona = New OleDbDataAdapter("select * from N_ZONAS", cn)
        daNomZona.Fill(ds, "nomZona")


        cn.Close()

        Dim claves(0) As DataColumn
        claves(0) = New DataColumn
        claves(0) = ds.Tables("misEstaciones").Columns("CODESTACION")
        ''claves(1) = New DataColumn
        ''claves(1) = ds.Tables("mis_clientes").Columns("asd")
        ds.Tables("misEstaciones").PrimaryKey = claves

      


    End Sub
    Sub EnlazarADatos()

        bs.DataSource = ds
        bs.DataMember = "misEstaciones"

        txtCodigoEst.DataBindings.Add(New Binding("text", bs, "CODESTACION", True))
        txtEstacion.DataBindings.Add(New Binding("text", bs, "NOMESTACION", True))
        txtKilometros.DataBindings.Add(New Binding("text", bs, "KILOMETROS", True))
        txtEsquiables.DataBindings.Add(New Binding("text", bs, "ESQUIABLES", True))


        txtZona.DataBindings.Add(New Binding("text", bs, "ZONESTACION", True))

        Dim cbEstacion As OleDbCommandBuilder = New OleDbCommandBuilder(daEstaciones)
        '


    End Sub

    Private Sub bs_PositionChanged(sender As Object, e As System.EventArgs) Handles bs.PositionChanged
        Dim idZona As Integer
        Try
            idZona = txtZona.Text
            ds.Tables("nomZona").Clear()
            daNomZona = New OleDbDataAdapter("select * from N_ZONAS where CODZONA =" & idZona, cn)
            daNomZona.Fill(ds, "nomZona")
            txtNombreZona.DataBindings.Add(New Binding("text", bs, "NOMZONA", True))
        Catch ex As Exception

        End Try
        Try
            txtPorcentEsquiable.Text = txtEsquiables.Text * 100 / txtKilometros.Text
        Catch ex As Exception

        End Try

        If bs.Position = 0 Then
            btnAnterior.Enabled = False
            btnSiguiente.Enabled = True
        ElseIf bs.Position = bs.Count - 1 Then
            btnSiguiente.Enabled = False
            btnAnterior.Enabled = True
        Else
            btnSiguiente.Enabled = True
            btnAnterior.Enabled = True

        End If

        
    End Sub

    Private Sub btnSiguiente_Click(sender As System.Object, e As System.EventArgs) Handles btnSiguiente.Click
        bs.MoveNext()
    End Sub

    Private Sub btnAnterior_Click(sender As System.Object, e As System.EventArgs) Handles btnAnterior.Click
        bs.MovePrevious()
    End Sub

    Private Sub btnNuevaMedicion_Click(sender As System.Object, e As System.EventArgs) Handles btnNuevaMedicion.Click
        Dim esquiables = txtEsquiables.Text
     


        If esquiables > txtKilometros.Text Then
            MessageBox.Show("No se puede superar los kilometros de la estación")
        Else
            
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("¿Desea modificar los km. esquiables???", "ski ska ska ski", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2)
                If respuesta = Windows.Forms.DialogResult.Yes Then
                MessageBox.Show("Modificadoo")

                    If Not ds.GetChanges() Is Nothing Then
                        Try
                        daEstaciones.Update(ds, "misEstaciones")
                            ds.AcceptChanges()

                        Catch ex As Exception
                            MessageBox.Show("Error al actualizar" & ex.Message)
                        ds.Tables("misEstaciones").RejectChanges()

                        End Try

                    End If
                Else
                    MessageBox.Show("nooooo")
                End If

        End If
    End Sub

    Private Sub txtEsquiables_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtEsquiables.KeyPress
        If Not e.KeyChar Like "[0-9]" And e.KeyChar <> Convert.ToChar(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnBusqueda_Click(sender As System.Object, e As System.EventArgs) Handles btnBusqueda.Click
        Dim dr As DataRow
        Try
            dr = ds.Tables("misEstaciones").Rows.Find(txtBuscar.Text)
        Catch ex As Exception

        End Try

        If Not (dr Is Nothing) Then
            imgExiste.ImageLocation = Application.StartupPath & "\ok.jpg"

            txtCodigoEst.Text = dr("CODESTACION")
            txtEstacion.Text = dr("NOMESTACION")
            txtKilometros.Text = dr("KILOMETROS")
            txtEsquiables.Text = dr("ESQUIABLES")
            txtZona.Text = dr("ZONESTACION")
            txtPorcentEsquiable.Text = txtEsquiables.Text * 100 / txtKilometros.Text


        Else
            MessageBox.Show("ERROR: No existe.")
            
            imgExiste.ImageLocation = Application.StartupPath & "\no.jpg"

        End If
    End Sub

    Private Sub btnSalir_Click(sender As System.Object, e As System.EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnInforme_Click(sender As System.Object, e As System.EventArgs) Handles btnInforme.Click
        frmInforme.Show()

    End Sub

    Private Sub btnHospital_Click(sender As System.Object, e As System.EventArgs) Handles btnHospital.Click
        Form3.Show()
    End Sub
End Class
