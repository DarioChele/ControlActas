Imports System.Data.OleDb
Public Class cls_querys
    Public arreglo_usuario(3) As String
    Private _descripcion As String
    Private _contrasena As String
    Dim obj_conexion As New cls_conexion

    Property descripcion()
        Get
            Return (Me._descripcion)
        End Get
        Set(ByVal value)
            Me._descripcion = value
        End Set
    End Property
    Property contrasena()
        Get
            Return (Me._contrasena)
        End Get
        Set(ByVal value)
            Me._contrasena = value
        End Set
    End Property
    Sub New(ByVal descripcion As String, ByVal contrasena As String)
        Me.descripcion = descripcion
        Me.contrasena = contrasena
    End Sub
    'Protected Overrides Sub finalize()
    '    Me.finalize()
    'End Sub
    Sub New()

    End Sub


    Public Function buscar_usuario()
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        cmd.Connection = obj_conexion.con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "Select * from usuario where descripcion = '" & Me.descripcion & "' and contrasena = '" & Me.contrasena & "'"
        Try
            dr = cmd.ExecuteReader
            If dr.Read Then
                arreglo_usuario(0) = dr("codigo").ToString
                arreglo_usuario(1) = dr("descripcion").ToString
                arreglo_usuario(2) = dr("contrasena").ToString
                arreglo_usuario(3) = dr("tipo").ToString
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
            cmd = Nothing
        End Try
        obj_conexion.desconectar()
        Return arreglo_usuario
    End Function

    Function query(ByVal buscar As String, ByVal condicion As String, ByVal tabla As String)
        Dim devolver As String
        devolver = ""
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        cmd.Connection = obj_conexion.con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "Select " & buscar & " from " & tabla & " where " & condicion
        Try
            dr = cmd.ExecuteReader
            If dr.Read Then
                devolver = dr(buscar).ToString
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
            cmd = Nothing
        End Try
        obj_conexion.desconectar()

        Return devolver
    End Function

    Function query_especial(ByVal clausula As String, ByVal buscar As String, ByVal alias_columna As String, ByVal tabla As String)
        Dim devolver As String
        devolver = ""
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        cmd.Connection = obj_conexion.con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "Select " & clausula & "(" & buscar & ") " & alias_columna & " from " & tabla
        Try
            dr = cmd.ExecuteReader
            If dr.Read Then
                devolver = dr(alias_columna).ToString
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
            cmd = Nothing
        End Try
        obj_conexion.desconectar()
        Return devolver
    End Function


    Public Sub eliminar_dato(ByVal tabla As String, ByVal condicion As String)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        cmd.Connection = obj_conexion.con
        cmd.CommandType = CommandType.Text
        Try
            cmd.CommandText = "delete from " & tabla & " where " & condicion
            cmd.ExecuteNonQuery()
            cmd = Nothing
            'MsgBox("Registro Eliminado Con Exito", MsgBoxStyle.OkOnly, "Mensaje")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
    End Sub
    Public Function actualiza_usuario(ByVal codigo As String, ByVal descripcion As String, ByVal contrasena As String, ByVal tipo As String)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        cmd.Connection = obj_conexion.con
        cmd.CommandType = CommandType.Text
        Try
            cmd.CommandText = "update usuario set descripcion= '" & descripcion & "', contrasena= '" & contrasena & "', tipo= '" & tipo & "' where codigo = '" & codigo & "'"
            cmd.ExecuteNonQuery()
            cmd = Nothing
            MsgBox("Registro Modificado Con Exito", MsgBoxStyle.OkOnly, "Mensaje")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
        Return True
    End Function
    Public Function insertar_cabecera(ByVal numero_acta As Integer, ByVal fecha As Date, ByVal codigo_recibe As String, ByVal observacion As String, ByVal tipo As String, ByVal status As String, ByVal contrato As String, ByVal egreso As String, ByVal ingreso As String, ByVal persona_entrega As String)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Try
            cmd.Connection = obj_conexion.con
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "insert into cabecera_acta values (" & numero_acta & ", '" & fecha & "', '" & codigo_recibe & "', '" & observacion & "', '" & tipo & "', '" & status & "', '" & contrato & "', '" & egreso & "', '" & ingreso & "','" & persona_entrega & "')"
            cmd.ExecuteNonQuery()
            cmd = Nothing
            MsgBox("Registro ingresado con exito")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
        Return True
    End Function
    Public Function insertar_detalle(ByVal numero_acta As Integer, ByVal numero_linea As Integer, ByVal cantidad As Integer, ByVal u_funcionamiento As String, ByVal serie As String, ByVal descripcion As String, ByVal status As String)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Try
            cmd.Connection = obj_conexion.con
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "insert into detalle_acta values (" & numero_acta & ", '" & numero_linea & "', '" & cantidad & "', '" & u_funcionamiento & "', '" & serie & "', '" & descripcion & "', '" & status & "')"
            cmd.ExecuteNonQuery()
            cmd = Nothing
            'MsgBox("Registro ingresado con exito")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
        Return True
    End Function
    Public Function actualiza_status_acta(ByVal status As String, ByVal numero_acta As Integer)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Try
            cmd.Connection = obj_conexion.con
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "update cabecera_acta set status = '" & status & "' where numero_acta = '" & numero_acta & "'"
            cmd.ExecuteNonQuery()
            cmd = Nothing
            'MsgBox("Registro ingresado con exito")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
        Return True
    End Function
    Public Function insertar_dato(ByVal tabla As String, ByVal columna As String, ByVal dato As String, ByVal condicion As String)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Try
            cmd.Connection = obj_conexion.con
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "update " & tabla & " set " & columna & " = '" & dato & "' where " & condicion
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
        Return True
    End Function
    Public Function insertar_usuario_acta(ByVal numero_acta As Integer)
        Dim codigo_usuario As String = ""
        codigo_usuario = Me.query("codigo", "descripcion = '" & frm_principal.tls_usuario.Text & "'", "usuario")
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Try
            cmd.Connection = obj_conexion.con
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "insert into usuario_acta values ('" & codigo_usuario & "', '" & numero_acta & "')"
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
        End Try
        obj_conexion.desconectar()
        Return True
    End Function
    Public Function insertar_usuario(ByVal codigo As String, ByVal descripcion As String, ByVal contrasena As String, ByVal tipo As String)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Try
            cmd.Connection = obj_conexion.con
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "insert into usuario values ('" & codigo & "', '" & descripcion & "', '" & contrasena & "', '" & tipo & "')"
            cmd.ExecuteNonQuery()
            cmd = Nothing
            MsgBox("Registro ingresado con exito")
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
        Return True
    End Function
    Public Function numero_acta()
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim numero As Integer = 0
        cmd.Connection = obj_conexion.con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select numero = max(numero_acta) from cabecera_acta"
        Try
            dr = cmd.ExecuteReader
            If dr.Read Then
                numero = Trim(dr("numero"))
            Else
                MsgBox("Error no Realizo la Consulta", MsgBoxStyle.OkOnly, "Mensaje")
            End If
            dr.Close()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
        Return (numero)
    End Function
    Public Function contar_datos(ByVal tabla As String, ByVal condicion As String)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Dim dr As OleDbDataReader
        Dim usuarios As Integer
        usuarios = 0
        cmd.Connection = obj_conexion.con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select contar=count(*) from " & tabla & " where " & condicion & ""
        Try
            dr = cmd.ExecuteReader
            If dr.Read Then
                usuarios = dr("contar")
            Else
                MsgBox("Error no hay cuentas de Usuario", MsgBoxStyle.OkOnly, "Mensaje")
            End If
            dr.Close()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()
        Return (usuarios)
    End Function

    Public Function guarda_dato(ByVal orden As String, ByVal tabla As String, ByVal dato As String)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Try
            cmd.Connection = obj_conexion.con
            cmd.CommandType = CommandType.Text
            cmd.CommandText = orden & " " & tabla & " values (" & dato & "')"
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        obj_conexion.desconectar()

        Return True
    End Function

    Public Function nuevo_empleado(ByVal vicepresidencia, ByVal gerencia, ByVal departamento, ByVal provincia, ByVal codigo, ByVal cedula, ByVal nombres, ByVal status, ByVal ubicacion)
        obj_conexion.conectar()
        Dim cmd As New OleDbCommand
        Try
            cmd.Connection = obj_conexion.con
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "insert into empleado values ('" & vicepresidencia & "','" & gerencia & "','" & departamento & "','" & provincia & "','" & codigo & "','" & cedula & "','" & nombres & "','" & status & "','" & ubicacion & "')"
            cmd.ExecuteNonQuery()
            cmd = Nothing
            MsgBox("Registro Ingresado con Exito")
        Catch ex As Exception
        End Try
        obj_conexion.desconectar()
        Return True
    End Function
End Class
