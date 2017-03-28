Imports System.Data
Imports System.Data.OleDb
Public Class Form3
    Private cn As New OleDbConnection
    Private daEnfermedades As OleDbDataAdapter
    Private daMedico As OleDbDataAdapter
    Private daPacientes As OleDbDataAdapter
    Private daSistema As OleDbDataAdapter
    Private daNomPacientes As OleDbDataAdapter
    Private ds As New DataSet
    Private nombreBBDD = Application.StartupPath & "\marzo_2017.accdb"
    Private WithEvents bs As New BindingSource

    Private Sub Form3_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ConfigurarAccesoADatos()
        EnlazarADatos()
        bs.MoveLast()
        bs.MoveFirst()

    End Sub
    Sub ConfigurarAccesoADatos()

        cn.ConnectionString = "provider= Microsoft.ace.oledb.12.0;" & "Data source = " & nombreBBDD
        cn.Open()
        daEnfermedades = New OleDbDataAdapter("select * from H_Enfermedad", cn)
        daEnfermedades.Fill(ds, "misEnfermedades")
        daMedico = New OleDbDataAdapter("select * from H_Medico", cn)
        daMedico.Fill(ds, "misMedicos")
        daPacientes = New OleDbDataAdapter("select * from H_Paciente", cn)
        daPacientes.Fill(ds, "misPacientes")
        daSistema = New OleDbDataAdapter("select * from H_Sistema", cn)
        daSistema.Fill(ds, "misSistemas")
        daSistema = New OleDbDataAdapter("select * from H_Sistema", cn)
        daSistema.Fill(ds, "paciente")
        cn.Close()




    End Sub
    Sub EnlazarADatos()

        bs.DataSource = ds
        bs.DataMember = "misEnfermedades"


        
        lstEnfermedades.DataSource = ds.Tables("misEnfermedades")
        lstEnfermedades.DisplayMember = "nombre_enfermedad"
        lstEnfermedades.ValueMember = "sistema_enfermedad"

        
        lstMedicos.DataSource = ds.Tables("misMedicos")
        lstMedicos.DisplayMember = "apellido_medico"
        lstMedicos.ValueMember = "codigo_medico"



    End Sub

    Private Sub lstEnfermedades_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstEnfermedades.SelectedIndexChanged
        Dim sistema As Integer
        Try
            sistema = ds.Tables("misEnfermedades").Rows(lstEnfermedades.SelectedIndex).Item("sistema_enfermedad")
            ds.Tables("misSistemas").Clear()
            daSistema = New OleDbDataAdapter("select * from H_sistema where codigo_sistema =" & sistema, cn)
            daSistema.Fill(ds, "misSistemas")
            bs.DataMember = "misSistemas"

            lblTipoSistema.DataBindings.Add(New Binding("text", bs, "nombre_sistema", True))
        Catch ex As Exception

        End Try

        




    End Sub

    Private Sub lstMedicos_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstMedicos.SelectedIndexChanged
        Dim claveMedico As Integer
        Try

            claveMedico = ds.Tables("misMedicos").Rows(lstMedicos.SelectedIndex).Item("codigo_medico")

        Catch ex As Exception

        End Try
        ds.Tables("paciente").Clear()
        daNomPacientes = New OleDbDataAdapter("select * from H_Paciente where medico_paciente = " & claveMedico, cn)
        daNomPacientes.Fill(ds, "paciente")
        lstPacientes.DataSource = ds.Tables("paciente")
        lstPacientes.DisplayMember = "apellido_paciente"
        lstPacientes.ValueMember = "codigo_paciente"
    End Sub

    Private Sub btnSalir_Click(sender As System.Object, e As System.EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub
End Class