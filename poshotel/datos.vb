Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Linq

Module datos
    Public sConexion As String = String.Empty

#Region "Checando xml y conexion"
    Sub Lectura_xml_conf(Ruta As String)
        Dim reader As XmlTextReader = New XmlTextReader(Ruta)
        Dim Servidor As String, Basedatos As String, Usuario As String, Clave As String
        Dim Datos As New DataSet
        Dim c As Integer, a As Integer

        On Error GoTo Errores
        Datos.ReadXml(reader)

        For c = 0 To Datos.Tables(0).Rows().Count - 1
            Servidor = Datos.Tables(0).Rows(c).Item(0).ToString
            Basedatos = Datos.Tables(0).Rows(c).Item(1).ToString
            Usuario = Datos.Tables(0).Rows(c).Item(2).ToString
            Clave = Datos.Tables(0).Rows(c).Item(3).ToString
        Next

        sConexion = "Server=" & Servidor & ";Database=" & Basedatos & ";User Id=" & Usuario & ";
Password=" & Clave & ";"

        Exit Sub

Errores:
        MessageBox.Show("No pudo leerse el archivo de configuración y arranque...")

    End Sub

    Function ChecandoConexion(sCadena As String) As Boolean
        Dim Conec As New SqlClient.SqlConnection

        Conec.ConnectionString = sCadena

        Try
            Conec.Open()
            ' Insert code to process data.
        Catch ex As Exception
            ChecandoConexion = False
            MessageBox.Show("Fallo la conexión a la base de datos")
        Finally
            Conec.Close()
            ChecandoConexion = True
        End Try
    End Function
#End Region
#Region "Funciones y procesos de datos"

    Sub GrabarBitacora(sConcepto As String, sDescripcion As String, iSuc As Integer, iUs As Integer)
        Dim sVista As String = Oracionsql(1, "MAX(row)+1", "Bitacora", "concepto = '" & sConcepto.Trim.ToUpper & "' and idsuc =" & iSuc.ToString.Trim)
        Dim iLinea As Integer = Val(Obtenervalor(sVista))
        Dim svalores As String = iLinea.ToString.Trim & ", '" & sConcepto.Trim.ToUpper & "', '" & sDescripcion.ToUpper.Trim & "',getdate()," & iSuc.ToString.Trim & "," & iUs.ToString.Trim
        Dim sQuery As String = Oracionsql(2, "row,concepto,descripcion,fecha,idsuc,iduser", "Bitacora", "", svalores)


        Ejecutaquery(sQuery)

    End Sub
    Function Noexiste(Cadena As String, Tabla As String, Condicion As String) As Boolean
        Dim Conectado As New SqlConnection(Cadena)
        Dim tblTabla As New DataTable
        Dim sQuery As String = Oracionsql(1, "", Tabla, Condicion)

        Conectado.Open()

        Using Adaptador As New SqlDataAdapter(sQuery, Conectado)
            Adaptador.Fill(tblTabla)
        End Using

        Noexiste = IIf(tblTabla.Rows.Count < 1, True, False)

        Conectado.Close()

    End Function

    Function Oracionsql(TipoOperacion As Integer, Campos As String, Tabla As String, Condicion As String, Optional Valores As String = "", Optional Orden As String = "") As String
        Dim Oracion As String

        Select Case TipoOperacion
            Case 1
                Oracion = "Select " & IIf(Campos.Trim.Length = 0, "*", Campos) & " from " & Tabla.Trim &
                    IIf(Condicion.Trim.Length > 0, " where " & Condicion.Trim, String.Empty) & IIf(Orden.Trim.Length > 0, " order by " & Orden.Trim & " asc", String.Empty)
            Case 2
                Oracion = "insert into " & Tabla.Trim & " (" & Campos.Trim & ") values (" & Valores.Trim & ")"
            Case 3
                Oracion = "truncate table " & Tabla.Trim
            Case 4
                Oracion = "delete from " & Tabla.Trim & IIf(Condicion.Trim.Length > 0, " where " & Condicion.Trim, String.Empty)
            Case 5
                Oracion = "update " & Tabla.Trim & " set " & Valores.Trim & " where " & Condicion.Trim
        End Select
        Oracionsql = Oracion
    End Function

    Function PasardatosaTabla(sQuery As String) As DataTable
        Dim Conectado As New SqlConnection(sConexion)
        Dim tblTabla As New DataTable
        Dim aLista As New ArrayList()

        Conectado.Open()

        Using Adaptador As New SqlDataAdapter(sQuery, Conectado)
            Adaptador.Fill(tblTabla)
        End Using

        PasardatosaTabla = tblTabla.Copy

        Conectado.Close()

    End Function

    Function Obtenervalor(sQuery As String) As String
        Dim Conectado As New SqlConnection(sConexion)
        Dim cmComando As New SqlCommand(sQuery, Conectado)

        Try
            Conectado.Open()

            Dim reader As SqlDataReader = cmComando.ExecuteReader()

            If reader.Read = True Then Obtenervalor = reader(0).ToString

        Catch ex As Exception
            MessageBox.Show("Error en la conexión a la base de datos" & ex.Message)
            Obtenervalor = String.Empty
        Finally
            Conectado.Close()
        End Try


    End Function

    Sub Ejecutaquery(sQuery As String)
        Dim Conectado As New SqlConnection(sConexion)
        Dim cmComando As New SqlCommand(sQuery, Conectado)
        Dim iEstado As Integer

        Try
            Conectado.Open()
            iEstado = cmComando.ExecuteNonQuery

        Catch ex As Exception
            MessageBox.Show("Error en la conexión a la base de datos" & ex.Message)
        Finally
            Conectado.Close()
        End Try

    End Sub

    Sub LlenarListado(lbLista As ListBox, sQuery As String)

        lbLista.Items.Clear()
        For Each row As DataRow In PasardatosaTabla(sQuery).Rows
            lbLista.Items.Add(row(0).ToString.Trim.ToUpper)

        Next

    End Sub

    Sub Autocompleta_textbox(oTexto As TextBox, sQuery As String)
        Dim oTabla As DataTable = PasardatosaTabla(sQuery)

        With oTexto
            .AutoCompleteCustomSource.Clear()
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            .AutoCompleteMode = AutoCompleteMode.Suggest

            For Each Linea As DataRow In oTabla.Rows
                .AutoCompleteCustomSource.Add(Linea(0).ToString.Trim.ToUpper)
            Next

        End With
    End Sub

#End Region
End Module
