Imports System.Text.RegularExpressions

Module Procesos
    Public iBandera1 As Integer
    Public sEmpresa As String, sUsername As String
    Public iSuperusuarioaut As Integer = 1
    Public iAcceso As Integer = 0
    Public sOrigen As String


    Sub Main(sRuta As String)

        If iBandera1 = 0 Then
            Lectura_xml_conf(sRuta.Trim & "\config.xml")
            If sConexion.Trim.Length = 0 Then End
            If ChecandoConexion(sConexion.Trim) = False Then MsgBox("No hay conexión a base de datos...") : End
        End If
    End Sub

    Sub AutocompCtexto(tControl As TextBox, sQuery As String)
        Dim aControl As AutoCompleteStringCollection

        aControl = New AutoCompleteStringCollection

        For Each row As DataRow In PasardatosaTabla(sQuery).Rows
            aControl.Add(row(0).ToString.Trim)
        Next

        tControl.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        tControl.AutoCompleteSource = AutoCompleteSource.CustomSource
        tControl.AutoCompleteCustomSource = aControl

    End Sub

    Sub AutocompCombo(tControl As ComboBox, sQuery As String)
        Dim aControl As AutoCompleteStringCollection

        aControl = New AutoCompleteStringCollection

        For Each row As DataRow In PasardatosaTabla(sQuery).Rows
            aControl.Add(row(0).ToString.Trim)
            tControl.Items.Add(row(0).ToString)
        Next

        tControl.AutoCompleteMode = AutoCompleteMode.Append
        tControl.AutoCompleteSource = AutoCompleteSource.CustomSource
        tControl.AutoCompleteCustomSource = aControl

    End Sub

    Function ExisteValorcelda(dgControl As DataGridView, iColumna As Integer, iFila As Integer, sValor As String) As Boolean
        Dim bResp As Boolean = False
        Dim iAvanza As Integer = 0

        If (dgControl.RowCount - 1) = 1 Then ExisteValorcelda = False : Exit Function


        For Each row As DataGridViewRow In dgControl.Rows
            iAvanza = iAvanza + 1
            If dgControl.NewRowIndex = iAvanza Then bResp = False : Exit For
            If row.Cells(iColumna).Value.ToString.ToUpper.Trim = sValor.Trim.ToUpper Then bResp = True : Exit For
        Next

        ExisteValorcelda = bResp

    End Function

    Function ObtenerFecha() As String
        ObtenerFecha = Date.Now
    End Function

    Function Usuarioconectado()
        Usuarioconectado = "No hay usuario conectado..."
    End Function

    Function validar_nombre(sNombre As String) As Boolean
        Dim sPatronex As New System.Text.RegularExpressions.Regex("([A-Z]{1}[a-z]{1,30}[- ]{0,1}|[A-Z]{1}[- \']{1}[A-Z]{0,1}[a-z]{1,30}[- ]{0,1}|[a-z]{1,2}[ -\']{1}[A-Z]{1}[a-z]{1,30}){2,5}")
        Dim bResultado As System.Text.RegularExpressions.Match = sPatronex.Match(sNombre)

        validar_nombre = bResultado.Success
    End Function

    Function Validando_usuario(sUser As String, sPassw As String) As Boolean

        If Noexiste(sConexion, "Usuarios", "Usuario = '" & sUser.Trim.ToUpper & "' and Password = '" & sPassw.Trim.ToUpper & "'") = True Then
            Validando_usuario = False
            MessageBox.Show("No existe el registro de tal usuario, volver a intentarlo...")
            Exit Function
        End If

        If sUser.Trim.ToUpper = "ADMIN" Then
            If Noexiste(sConexion, "bitacora", "concepto = 'ACCESO' and Descripcion like '%ADMIN%'") = False Then
                Validando_usuario = False
                MessageBox.Show("Ya existe registro de acceso de ADMIN, no puede proceder a utilizarlo en el futuro..")
                Exit Function
            End If
        End If
        Validando_usuario = True

    End Function

    Function solo_fecha(sFecha As String) As Date
        Dim sFe As String = String.Empty
        Dim sDe As String = String.Empty
        Dim iC As Integer = 0

        sDe = InStrRev(sFecha.Trim, " ")
        sFe = Left(sFecha.Trim, sDe).Trim

        solo_fecha = sFe
    End Function

    Function solo_hora(sFecha As String) As String
        Dim sFe As String = Left(sFecha.Trim, InStrRev(sFecha.Trim, " "))
        Dim sHo As String = Mid(sFecha.Trim, InStrRev(sFecha.Trim, " "))

        sFe = Replace(sFe.Trim, " ", "/")
        Dim dFe As Date = Date.Parse(sFe)

        sHo = Date.Parse(Left(sHo.Trim, 5))

        solo_hora = sHo
    End Function

    Function Contar_caracteres(sOracion As String, sEncontrar As String) As Integer
        Dim iC As Integer = 0
        Dim iE As Integer = 0

        For iC = 1 To sOracion.Trim.Length
            If Mid(sOracion.Trim, iC, 1) = sEncontrar Then iE = iE + 1
        Next

        Contar_caracteres = iE

    End Function

End Module
