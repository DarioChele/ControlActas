Imports System.Data.OleDb

Public Class cls_empleado
    Private _vicepresidencia As String
    Private _gerencia As String
    Private _departamento As String
    Private _provincia As String
    Private _codigo As String
    Private _cedula As String
    Private _nombres As String
    Private _status As String
    Private _ubicacion As String
    Property vicepresidencia()
        Get
            Return Me._vicepresidencia
        End Get
        Set(ByVal value)
            Me._vicepresidencia = value
        End Set
    End Property
    Property gerencia()
        Get
            Return Me._gerencia
        End Get
        Set(ByVal value)
            Me._gerencia = value
        End Set
    End Property
    Property departamento()
        Get
            Return Me._departamento
        End Get
        Set(ByVal value)
            Me._departamento = value
        End Set
    End Property
    Property provincia()
        Get
            Return Me._provincia
        End Get
        Set(ByVal value)
            Me._provincia = value
        End Set
    End Property
    Property codigo()
        Get
            Return Me._codigo
        End Get
        Set(ByVal value)
            Me._codigo = value
        End Set
    End Property
    Property cedula()
        Get
            Return Me._cedula
        End Get
        Set(ByVal value)
            Me._cedula = value
        End Set
    End Property
    Property nombres()
        Get
            Return Me._nombres
        End Get
        Set(ByVal value)
            Me._nombres = value
        End Set
    End Property
    Property status()
        Get
            Return Me._status
        End Get
        Set(ByVal value)
            Me._status = value
        End Set
    End Property
    Property ubicacion()
        Get
            Return Me._ubicacion
        End Get
        Set(ByVal value)
            Me._ubicacion = value
        End Set
    End Property

    Sub New()
    End Sub
    Sub New(ByVal vicepresidencia As String, ByVal gerencia As String, ByVal departamento As String, ByVal provincia As String _
             , ByVal codigo As String, ByVal cedula As String, ByVal nombres As String, ByVal status As String, ByVal ubicacion As String)
        Me.vicepresidencia = vicepresidencia.ToUpper
        Me.gerencia = gerencia.ToUpper
        Me.departamento = departamento.ToUpper
        Me.provincia = provincia.ToUpper
        Me.codigo = codigo.ToUpper
        Me.cedula = cedula.ToUpper
        Me.nombres = nombres.ToUpper
        Me.status = status.ToUpper
        Me.ubicacion = ubicacion.ToUpper
        Dim obj_query As New cls_querys
        If obj_query.query("codigo", "codigo = '" & Me.codigo & "'", "empleado").ToString <> "" Then
            MsgBox("Codigo ya esta en Uso" & vbCrLf & "Asigne otro Codigo de Empleado", MsgBoxStyle.Exclamation, "Error")
        Else
            obj_query.nuevo_empleado(Me.vicepresidencia, Me.gerencia, Me.departamento, Me.provincia, Me.codigo, Me.cedula, Me.nombres, Me.status, Me.ubicacion)
        End If

    End Sub
    Public Function borrar(ByVal codigo As String)
        Dim obj_query As New cls_querys
        obj_query.eliminar_dato("empleado", "codigo = '" & codigo & "'")
        Return True
    End Function
    Public Function consultar(ByVal dato As String, ByVal condicion As String, ByVal tabla As String)
        Dim obj_query As New cls_querys
        Return obj_query.query(dato, condicion, tabla)
    End Function

    Public Function insertar_dato(ByVal tabla As String, ByVal columna As String, ByVal dato As String, ByVal condicion As String)
        Dim obj_query As New cls_querys
        obj_query.insertar_dato(tabla, columna, dato.ToUpper, condicion)
        Return True
    End Function
    Public Function guardar_dato(ByVal codigo As String, ByVal descripcion As String)
        Dim obj_query As New cls_querys
        obj_query.guarda_dato("insert into", "status", "'" & codigo & "', '" & descripcion)

        Return True
    End Function

    Public Function consulta_especial(ByVal clausula As String, ByVal dato As String, ByVal alias_columna As String, ByVal tabla As String)
        Dim obj_query As New cls_querys
        Return obj_query.query_especial(clausula, dato, alias_columna, tabla)
    End Function
    Public Function actualizar_dato(ByVal cogigo As String, ByVal columna As String, ByVal dato As String)
        Dim obj_query As New cls_querys
        obj_query.insertar_dato("empleado", columna, dato, "codigo = '" & codigo & "'")
        ''update <tabla> set < columna > = < dato > where < condicion >
        Return True
    End Function

End Class
