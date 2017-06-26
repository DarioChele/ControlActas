Public Class cls_usuario
    Private _codigo As String
    Private _descripcion As String
    Private _contrasena As String
    Private _tipo As String
    Dim obj_query As New cls_querys
    Property codigo()
        Get
            Return Me._codigo
        End Get
        Set(ByVal value)
            Me._codigo = value
        End Set
    End Property

    Property descripcion()
        Get
            Return Me._descripcion
        End Get
        Set(ByVal value)
            Me._descripcion = value
        End Set
    End Property

    Property contrasena()
        Get
            Return Me._contrasena
        End Get
        Set(ByVal value)
            Me._contrasena = value
        End Set
    End Property
    Property tipo()
        Get
            Return Me._tipo
        End Get
        Set(ByVal value)
            Me._tipo = value
        End Set
    End Property
    Sub New(ByVal user As String, ByVal password As String)
        Me.descripcion = user
        Me.contrasena = password
    End Sub
    Sub New(ByVal user As String, ByVal password As String, ByVal type As String)
        Me.descripcion = user
        Me.contrasena = password
        Me.tipo = type
    End Sub
    Sub New()
    End Sub
    Public Sub agregar()
        Dim sigue As String = obj_query.query("descripcion", "descripcion = '" & Me.descripcion & "'", "usuario")
        If sigue = "" Then
            Dim bandera As String = "sigue"
            Dim i As Integer = 1
            Dim codigo As String = ""
            While bandera <> ""
                i = i + 1
                If i < 10 Then
                    codigo = "00" & i.ToString
                Else
                    codigo = "0" & i.ToString
                End If
                bandera = obj_query.query("descripcion", "codigo = '" & codigo & "'", "usuario")
            End While
            Me.codigo = codigo
            obj_query.insertar_usuario(Me.codigo, Me.descripcion, Me.contrasena, Me.tipo)
        Else
            MsgBox("Registro ya existe")
        End If
    End Sub
    Public Function obtener_dato(ByVal buscar As String, ByVal condicion As String, ByVal tabla As String)
        Dim valor As String = obj_query.query(buscar, condicion, tabla)
        Return valor
    End Function

    Public Sub actualiza_usuario()
        obj_query.actualiza_usuario(Me.codigo, Me.descripcion, Me.contrasena, Me.tipo)
    End Sub
    Public Sub eliminar(ByVal descripcion As String)
        Me.descripcion = descripcion
        obj_query.eliminar_dato("usuario", "descripcion = '" & Me.descripcion.ToString & "'")
    End Sub
End Class
