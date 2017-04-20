Imports Telerik.WinControls.UI
Imports Telerik.WinControls.Layouts
Imports System.Linq
Imports Telerik.WinControls
Imports Telerik.WinControls.Primitives
Imports Telerik.WinControls.Themes
Imports System.ComponentModel

Public Class RadForm1

    Private iConteopriv As Integer
    Private iSucursal As Integer
    Private iUsuario As Integer
    Private Objeto As New RadLiveTileElement
    Private tbServHotel As New DataTable
    Private tbDiversos As New DataTable
    Private tbHabitaciones As New DataTable
    Private tbClientes As New DataTable

    Public Sub Bloqueopantprincipal(Optional iOp As Integer = 0)
        Dim sQuery As String = String.Empty
        Dim iConta As Integer = 0

        Select Case iOp
            Case 4
                RadButton2.Enabled = True
                TabPage1.Parent = Nothing
                TabPage2.Parent = Nothing
                TabPage6.Parent = Nothing
                TabPage3.Parent = Nothing
                TabPage7.Parent = TabControl1

            Case 2
                RadButton2.Enabled = True
                TabPage1.Parent = Nothing
                TabPage2.Parent = Nothing
                TabPage6.Parent = Nothing
                TabPage7.Parent = Nothing
                TabPage3.Parent = TabControl1

                TabPage4.Parent = Nothing
                TabPage5.Parent = TabControl2
            Case 0
                RadButton2.Enabled = False
                RadDropDownButton1.Enabled = False
                TabPage2.Parent = Nothing
                TabPage3.Parent = Nothing
                TabPage6.Parent = Nothing
                TabPage7.Parent = Nothing
            Case 1
                RadButton2.Enabled = False
                RadDropDownButton1.Enabled = True
                TabPage1.Parent = Nothing
                TabPage2.Parent = TabControl1
                TabPage3.Parent = Nothing
                TabPage6.Parent = Nothing
                TabPage7.Parent = Nothing
            Case 3
                FlowLayoutPanel1.Controls.Clear()
                FlowLayoutPanel2.Controls.Clear()

                RadButton2.Enabled = True
                BotonesPrecios()
                BotonesHabitacion()
                RadDateTimePicker1.Value = Date.Today
                TabPage1.Parent = Nothing
                TabPage2.Parent = Nothing
                TabPage6.Parent = Nothing
                TabPage7.Parent = Nothing
                TabPage3.Parent = TabControl1
                TabPage4.Focus()

                tbClientes = PasardatosaTabla(Oracionsql(1, "*", "clientes", ""))

                Autocompleta_textbox(TextBox11, Oracionsql(1, "cliente", "clientes", "",, "cliente"))

                tarjetas.Parent = Nothing
        End Select
    End Sub

    Sub BotonesHabitacion()
        Dim sQuery As String = Oracionsql(1, "hab,personas,maximoper", "habitaciones", "status = 1 and idsuc = " & iSucursal.ToString.Trim)
        Dim oTablahab As DataTable = PasardatosaTabla(sQuery)
        Dim sHab As String = String.Empty
        Dim iPer As Integer = 0
        Dim iMax As Integer = 0

        For Each row As DataRow In oTablahab.Rows
            sHab = row.Item("hab")
            iPer = row.Item("personas")
            iMax = row.Item("maximoper")
            Dim oBotoh As New RadToggleButton
            With oBotoh
                .Text = sHab.ToUpper.Trim
                .Tag = iPer.ToString.Trim & "_" & iMax.ToString.Trim
                .Width = 60 * sHab.Trim.Length
                If Objeto.Name.ToString.Substring(1) = sHab.Trim Then .ToggleState = Enumerations.ToggleState.On
                .ThemeName = "TelerikMetroTouch"
            End With
            AddHandler oBotoh.Click, AddressOf obotoh_Click
            FlowLayoutPanel2.Controls.Add(oBotoh)
        Next
    End Sub
    Private Sub obotoh_Click(sender As Object, e As EventArgs)

    End Sub
    Sub BotonesPrecios()
        Dim sQuery As String = Oracionsql(1, "idsrv,servicio, isnull(on_active,0) as on_active, precio_efectivo,precio_tarjeta,extra_ocupante_efe,extra_ocupante_tar,finaliza, non_date", "vservicios", "status = 1 and idsuc = " & iSucursal.ToString.Trim & " and (isnull(dire_pos_a,0) <= Duracion_hrs and isnull(dire_pos_a,0) >=0)",, "idsrv")
        Dim oTablarecep As DataTable = PasardatosaTabla(sQuery)
        Dim sServ As String = String.Empty
        Dim iServ As Integer = 0
        Dim dFre As Date = Now
        Dim sResulta As String = String.Empty

        For Each row As DataRow In oTablarecep.Rows
            sServ = row.Item("servicio")
            iServ = row.Item("idsrv")
            Dim oBoton As New RadToggleButton
            With oBoton
                .Text = sServ.ToUpper.Trim
                .Name = iServ.ToString.Trim
                .Width = 8 * sServ.Trim.Length
                .ThemeName = "TelerikMetroTouch"
                If row.Item("on_active") = 1 Then
                    .ToggleState = Enumerations.ToggleState.On

                    With RadioButton1
                        .Checked = True
                        .Tag = row.Item("precio_efectivo")
                        TextBox2.Text = Format(.Tag, "currency")
                    End With
                    With RadioButton2
                        .Checked = False
                        .Tag = row.Item("precio_tarjeta")
                    End With
                    With TextBox1
                        .Text = Format(row.Item("extra_ocupante_efe"), "currency")
                        .Tag = row.Item("extra_ocupante_efe") & "-" & row.Item("extra_ocupante_tar")
                    End With


                    RadDateTimePicker1.Value = solo_fecha(row.Item("finaliza"))
                    RadDateTimePicker1.Tag = RadDateTimePicker1.Value.ToString

                    sResulta = solo_hora(row.Item("finaliza"))

                    RadTimePicker1.Value = sResulta
                    RadTimePicker1.Tag = sResulta
                End If
                .TextWrap = True
            End With

            If IsDBNull(row.Item("non_date")) = False Then
                If Val(row.Item("non_date")) = 1 Then RadDateTimePicker1.Enabled = False
            End If

            AddHandler oBoton.Click, AddressOf oboton_Click
            FlowLayoutPanel1.Controls.Add(oBoton)
        Next

    End Sub
    Private Sub oboton_Click(sender As Object, e As EventArgs)
        Dim sQuery As String = String.Empty
        Dim oBot As RadToggleButton
        Dim sResulta As String = String.Empty
        Dim dFre As Date = Now

        oBot = sender

        sQuery = Oracionsql(1, "non_date", "vservicios", "status =1 and idsuc =" & iSucursal.ToString.Trim & " and idsrv = " & oBot.Name.Trim)
        Dim iMod As Integer = Val(Obtenervalor(sQuery))

        If iMod = 1 Then
            RadDateTimePicker1.Enabled = False
        Else
            RadDateTimePicker1.Enabled = True
        End If

        If oBot.ToggleState = Enumerations.ToggleState.On Then
            TextBox2.Text = String.Empty : TextBox1.Text = String.Empty
            RadioButton1.Checked = False : RadioButton2.Checked = False
            Exit Sub
        End If
        For Each oBotun As RadToggleButton In FlowLayoutPanel1.Controls
            If oBot.Name.Trim <> oBotun.Name.Trim Then oBotun.ToggleState = Enumerations.ToggleState.Off
        Next

        sQuery = Oracionsql(1, " * ", "vservicios", "status = 1 And idsuc =" & iSucursal.ToString.Trim & " And idsrv = " & oBot.Name.Trim)
        Dim Tablarecibe As DataTable = PasardatosaTabla(sQuery)

        With RadioButton1
            .Checked = True
            .Tag = Tablarecibe.Rows(0).Item("precio_efectivo")
            TextBox2.Text = Format(.Tag, "currency")
        End With
        With RadioButton2
            .Checked = False
            .Tag = Tablarecibe.Rows(0).Item("precio_tarjeta")
        End With
        With TextBox1
            If IsDBNull(Tablarecibe.Rows(0).Item("extra_ocupante_efe")) = False Then
                .Text = Format(Tablarecibe.Rows(0).Item("extra_ocupante_efe"), "currency")
                .Tag = Tablarecibe.Rows(0).Item("extra_ocupante_efe") & "-" & Tablarecibe.Rows(0).Item("extra_ocupante_tar")
            Else
                .Text = String.Empty
                .Tag = String.Empty
            End If
        End With

        RadDateTimePicker1.Tag = String.Empty
        RadDateTimePicker1.Value = solo_fecha(Tablarecibe.Rows(0).Item("finaliza"))
        RadDateTimePicker1.Tag = RadDateTimePicker1.Value

        sResulta = solo_hora(Tablarecibe.Rows(0).Item("finaliza"))
        dFre = sResulta

        RadTimePicker1.Value = dFre
        RadTimePicker1.Tag = dFre.ToString

    End Sub
    Sub Llenando_hab(iOpcion As Integer)
        Dim sQuery As String = Oracionsql(1, "hab", "habitaciones", "activa = 1 And status =1 And idsuc=" & iSucursal.ToString.Trim)
        Dim Tablareceptora As DataTable = PasardatosaTabla(sQuery)
        Dim iHab As Integer = 0

        For Each row As DataRow In Tablareceptora.Rows
            iHab = row.Item("hab")

            Dim radlElemento As New RadLiveTileElement
            With radlElemento
                .Text = iHab.ToString.Trim
                .TextWrap = True
                .ToolTipText = "Habitación " & iHab.ToString.Trim
                .Name = "H" & iHab.ToString.Trim
                .Tag = "Libres"

                .BackColor = Color.Blue
            End With
            AddHandler radlElemento.Click, AddressOf RadLElemento_Click
            With TileGroupElement1
                .Items.Add(radlElemento)
            End With
        Next
    End Sub
    Private Sub RadLElemento_Click(sender As Object, e As EventArgs)
        Dim Cosa As New RadLiveTileElement

        Objeto = sender
        Cosa = sender

        Select Case Cosa.Tag.ToString.Trim.ToUpper
            Case "LIBRES"
                Bloqueopantprincipal(3)
            Case "OCUPADOS"
                TileGroupElement1.Items.Add(Objeto)
                Cosa.Tag = "LIBRES"
                Cosa.BackColor = Color.Blue
        End Select

    End Sub
    Private Sub RadForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Procesos.Main(My.Application.Info.DirectoryPath)
        Bloqueopantprincipal()
        RadLabelElement1.Text = Date.Today
        Dim sQuery As String = Oracionsql(1, "Sucursal", "Sucursales", "activo = 1")
        RadLabelElement2.Text = Obtenervalor(sQuery)
        sQuery = Oracionsql(1, "idsuc", "Sucursales", "activo =1")
        iSucursal = Val(Obtenervalor(sQuery))
        sQuery = Oracionsql(1, "count(*)", "habitaciones", "")
        RadLabelElement3.Text = "Habitaciones " & Obtenervalor(sQuery).Trim
        RadLabelElement4.Text = String.Empty
        Llenando_hab(0)
    End Sub

    Private Sub RadButton3_Click(sender As Object, e As EventArgs) Handles RadButton3.Click
        End
    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click
        If RadTextBox1.Text.Trim.Length = 0 Then MessageBox.Show("Alguno de los campos de captura están vacíos proceda a editarlos...") : Exit Sub
        If RadTextBox2.Text.Trim.Length = 0 Then MessageBox.Show("Alguno de los campos de captura están vacíos proceda a editarlos...") : Exit Sub

        If Validando_usuario(RadTextBox1.Text, RadTextBox2.Text) = True Then
            MsgBox("Bienvenido, accediendo al punto de venta...", vbInformation, "POS HOTEL")
            Bloqueopantprincipal(1)
            RadDropDownButton1.Text = RadTextBox1.Text.Trim.ToUpper
            Dim sQuery As String = Oracionsql(1, "iduser", "Usuarios", "usuario = '" & RadTextBox1.Text.ToUpper.Trim & "'")
            iUsuario = Val(Obtenervalor(sQuery))
            GrabarBitacora("ACCESO", "Acceso usuario " & RadTextBox1.Text.ToUpper.Trim, iSucursal, iUsuario)
        Else
            iConteopriv = iConteopriv + 1
            If iConteopriv = 3 Then MessageBox.Show("Llego a las 3 oportunidades de acceso...")
            GrabarBitacora("ACCESO", "Acceso Fallido en " & iConteopriv.ToString.Trim & " oportunidades usuario " & RadTextBox1.Text.ToUpper.Trim, iSucursal, 0)
        End If
    End Sub

    Sub LlenarRaddrListado(lbLista As RadDropDownList, sQuery As String)

        lbLista.Items.Clear()
        For Each row As DataRow In PasardatosaTabla(sQuery).Rows
            lbLista.Items.Add(row(0).ToString.Trim.ToUpper)
        Next

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Calculo_efectivo(1000)
    End Sub

    Sub Calculo_efectivo(dValor As Double)

        If TextBox6.Text.Trim.Length = 0 Then Exit Sub

        TextBox7.Text = Format(Val(Format(TextBox7.Text, "General Number")) + dValor, "currency")

        If TextBox6.Text.Trim.Length > 0 Then
            TextBox8.Text = Format(Val(Format(TextBox7.Text, "General Number")) - Val(Format(TextBox6.Text, "General Number")), "Currency")
        End If

    End Sub

    Function Coste_elegido(iOpcion As Integer, Optional lPosN As Boolean = False, Optional lTotal As Boolean = False) As Double

        If lPosN = False Then Coste_elegido = Barrido_DT(iOpcion)
        If lPosN = True Then Coste_elegido = 0
        If lTotal = True Then
            Coste_elegido = Barrido_DT(iOpcion) + 0
        End If


    End Function

    Private Sub RadRadioButton1_ToggleStateChanged(sender As Object, args As StateChangedEventArgs)
    End Sub

    Private Sub RadRadioButton2_ToggleStateChanged(sender As Object, args As StateChangedEventArgs)
    End Sub

    Private Sub RadButton17_Click(sender As Object, e As EventArgs)

    End Sub

    Function Barrido_DT(iOp As Integer) As Double

    End Function

    Private Sub RadButton2_Click(sender As Object, e As EventArgs) Handles RadButton2.Click
        LimpiarControles()
        Bloqueopantprincipal(1)
    End Sub

    Sub LimpiarControles(Optional iOp As Integer = 0)

        Select Case iOp
            Case 0
                TextBox1.Text = String.Empty : TextBox2.Text = String.Empty
                DataGridView1.Rows.Clear()
                TextBox3.Text = String.Empty : TextBox4.Text = String.Empty : TextBox5.Text = String.Empty
                TextBox6.Text = String.Empty : TextBox7.Text = String.Empty : TextBox8.Text = String.Empty
                TextBox18.Text = String.Empty
                TextBox9.Text = String.Empty : TextBox10.Text = String.Empty
                TextBox11.Text = String.Empty : TextBox12.Text = String.Empty : TextBox13.Text = String.Empty
                TextBox14.Text = String.Empty : TextBox15.Text = String.Empty : TextBox16.Text = String.Empty
                TextBox17.Text = String.Empty
                RadDateTimePicker1.Tag = Nothing
                RadTimePicker1.Tag = Nothing
        End Select
    End Sub

    Private Sub RadButton16_Click(sender As Object, e As EventArgs)
        LimpiarControles()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Dim sR As String = String.Empty

        If RadioButton2.Checked = True Then

            efectivo.Parent = Nothing
            desctos.Parent = Nothing
            cliente.Parent = Nothing
            tarjetas.Parent = RadPageView1
            desctos.Parent = RadPageView1
            cliente.Parent = RadPageView1
            tarjetas.Focus()

            TextBox2.Text = Format(RadioButton2.Tag, "currency")
            TextBox1.Text = Format(Mid(TextBox1.Tag.Trim, InStr(TextBox1.Tag.Trim, "-", CompareMethod.Text) + 1, TextBox1.Tag.Trim.Length), "Currency")


            If TextBox1.Tag IsNot Nothing Then
                Calculaimportes(RadDateTimePicker1, RadTimePicker1, TextBox1, TextBox2, 2)
            End If

        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Dim sR As String = String.Empty

        If RadioButton1.Checked = True Then

            efectivo.Parent = Nothing
            tarjetas.Parent = Nothing
            desctos.Parent = Nothing
            cliente.Parent = Nothing
            efectivo.Parent = RadPageView1
            desctos.Parent = RadPageView1
            cliente.Parent = RadPageView1
            efectivo.Focus()

            If RadioButton1.Tag IsNot Nothing Then TextBox2.Text = Format(RadioButton1.Tag, "Currency") : TextBox1.Text = Format(Mid(TextBox1.Tag.Trim, 1, InStr(TextBox1.Tag.Trim, "-", CompareMethod.Text) - 1), "Currency")


            If TextBox1.Tag IsNot Nothing Then
                Calculaimportes(RadDateTimePicker1, RadTimePicker1, TextBox1, TextBox2, 1)
            End If
        End If
    End Sub

    Sub Habdeshabbuthues(sNombre As String)
        Dim aRadto() As RadToggleButton = {RadToggleButton1, RadToggleButton2, RadToggleButton3, RadToggleButton4, RadToggleButton5, RadToggleButton6, RadToggleButton7, RadToggleButton8, RadToggleButton9, RadToggleButton10}
        Dim iConteo As Integer = 0
        Dim sHab As String = String.Empty

        For Each oBut As RadToggleButton In FlowLayoutPanel2.Controls
            If oBut.ToggleState = Enumerations.ToggleState.On Then sHab = oBut.Text : Exit For
        Next

        Dim sQuery As String = Oracionsql(1, "maximoper", "habitaciones", "idsuc=" & iSucursal & " and hab =" & sHab.Trim & " and activa = 1 and status = 1")
        Dim iMax As Integer = Val(Obtenervalor(sQuery))


        If Val(sNombre.Trim) > iMax Then
            MsgBox("No puede rebasar al límite de huespedes, es sólo de " & iMax.ToString.Trim & "...", vbExclamation, "Cuidado...")
            For iConteo = 0 To 9 Step 1
                If aRadto(iConteo).Text.Trim = sNombre.Trim Then
                    If aRadto(iConteo).ToggleState = Enumerations.ToggleState.On Then
                        aRadto(iConteo).ToggleState = Enumerations.ToggleState.Off
                    End If
                End If
            Next
            Exit Sub
        End If

        For iConteo = 0 To 9 Step 1
            If aRadto(iConteo).Text.Trim <> sNombre.Trim Then
                If aRadto(iConteo).ToggleState = Enumerations.ToggleState.On Then
                    aRadto(iConteo).ToggleState = Enumerations.ToggleState.Off
                End If
            End If
        Next


    End Sub

    Private Sub RadToggleButton1_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton1.ToggleStateChanged
        Dim oBut As RadToggleButton = sender

        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)
    End Sub

    Private Sub RadToggleButton2_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton2.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)
    End Sub

    Private Sub RadToggleButton3_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton3.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)

    End Sub

    Private Sub RadToggleButton6_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton6.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)

    End Sub

    Private Sub RadToggleButton5_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton5.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)

    End Sub

    Private Sub RadToggleButton4_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton4.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)

    End Sub

    Private Sub RadToggleButton9_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton9.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)

    End Sub

    Private Sub RadToggleButton8_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton8.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)

    End Sub

    Private Sub RadToggleButton7_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton7.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)

    End Sub

    Private Sub RadToggleButton10_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton10.ToggleStateChanged
        Dim oBut As RadToggleButton = sender
        If args.ToggleState = Enumerations.ToggleState.On Then Habdeshabbuthues(oBut.Text)

    End Sub

    Private Sub RadButton4_Click(sender As Object, e As EventArgs) Handles RadButton4.Click
        Dim aRadto() As RadToggleButton = {RadToggleButton1, RadToggleButton2, RadToggleButton3, RadToggleButton4, RadToggleButton5, RadToggleButton6, RadToggleButton7, RadToggleButton8, RadToggleButton9, RadToggleButton10}
        Dim sHuespedes As String = String.Empty
        Dim sHab As String = String.Empty
        Dim sTipo As String = String.Empty
        Dim iConteo As Integer = 0

        If RadLabelElement4.Text.Trim.Length > 0 Then MsgBox("No puede proceder hasta que determine correctamente fecha y hora de salida...", vbExclamation, "Cuidado...") : Exit Sub

        For iConteo = 0 To 9 Step 1
            If aRadto(iConteo).ToggleState = Enumerations.ToggleState.On Then sHuespedes = aRadto(iConteo).Text : Exit For
        Next

        If sHuespedes.Trim.Length = 0 Then MsgBox("Debe determinar el número de huespedes de la habitación...", vbExclamation + vbCritical, "Cuidado...") : Exit Sub

        For Each oBut As RadToggleButton In FlowLayoutPanel2.Controls
            If oBut.ToggleState = Enumerations.ToggleState.On Then
                FlowLayoutPanel2.Controls.Remove(oBut)
                sHab = oBut.Text
                Exit For
            End If

        Next

        For Each oBut As RadToggleButton In FlowLayoutPanel1.Controls
            If oBut.ToggleState = Enumerations.ToggleState.On Then sTipo = oBut.Text : Exit For
        Next

        For iConteo = 0 To 9 Step 1
            If aRadto(iConteo).ToggleState = Enumerations.ToggleState.On Then
                aRadto(iConteo).ToggleState = Enumerations.ToggleState.Off
            End If
        Next

        Dim squery As String = Oracionsql(1, "personas,maximoper", "habitaciones", "activa=1 and status = 1 and idsuc =" & iSucursal.ToString.Trim & " and hab =" & sHab.Trim)

        Dim sTabhab As New DataTable

        sTabhab = PasardatosaTabla(squery)

        Dim iPersonas As Integer = sTabhab.Rows(0).Item("personas")
        Dim iMaximo As Integer = sTabhab.Rows(0).Item("maximoper")
        Dim iExc As Integer = 0
        Dim sMontexc As String = TextBox1.Text

        If Val(sHuespedes) > iPersonas Then
            If Val(sHuespedes) <= iMaximo Then
                iExc = Val(sHuespedes) - iPersonas
            End If
        End If

        Dim sTipopg As String = IIf(RadioButton1.Checked = True, "EFE", "TAR")
        Dim sImporth As String = TextBox2.Text

        Dim sFecha As String = Format(RadDateTimePicker1.Value, "dd/MM/yy")
        Dim sHora As String = Format(RadTimePicker1.Value, "HH:mm")

        Dim dTotal As Double = 0

        dTotal = Val(Format(sImporth, "General Number")) + IIf(iExc > 0, Val(Format(sMontexc, "General Number")), 0)

        TextBox3.Text = Format(Val(Format(TextBox3.Text, "General Number")) + dTotal, "Currency")

        TextBox6.Text = Format((Val(Format(TextBox3.Text, "General Number")) + Val(Format(TextBox4.Text, "General Number"))) - Val(Format(TextBox9.Text, "General Number")), "Currency")


        DataGridView1.Rows.Add(New String() {sHab, sTipopg, sHuespedes, iExc.ToString.Trim, sFecha, sHora, Format(dTotal, "Currency"), sTipo})

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Calculo_efectivo(1000)
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Calculo_efectivo(500)
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Calculo_efectivo(200)
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Calculo_efectivo(100)
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        Calculo_efectivo(50)
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Calculo_efectivo(20)
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Calculo_efectivo(10)
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        Calculo_efectivo(5)
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Calculo_efectivo(2)
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        Calculo_efectivo(1)
    End Sub

    Private Sub RadToggleButton11_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton11.ToggleStateChanged
        If args.ToggleState = Enumerations.ToggleState.On Then
            RadToggleButton12.ToggleState = Enumerations.ToggleState.Off
            RadToggleButton13.ToggleState = Enumerations.ToggleState.Off
        End If
    End Sub

    Private Sub RadToggleButton12_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton12.ToggleStateChanged
        If args.ToggleState = Enumerations.ToggleState.On Then
            RadToggleButton11.ToggleState = Enumerations.ToggleState.Off
            RadToggleButton13.ToggleState = Enumerations.ToggleState.Off
        End If
    End Sub

    Private Sub RadToggleButton13_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RadToggleButton13.ToggleStateChanged
        If args.ToggleState = Enumerations.ToggleState.On Then
            RadToggleButton11.ToggleState = Enumerations.ToggleState.Off
            RadToggleButton12.ToggleState = Enumerations.ToggleState.Off
        End If
    End Sub

    Private Sub DataGridView1_UserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs) Handles DataGridView1.UserDeletingRow
        Dim sImporte As String = DataGridView1.Rows(e.Row.Index).Cells(6).Value

        If MsgBox("Desea eliminar el registro?", vbYesNo + vbCritical, "Confirme...") = vbNo Then e.Cancel = True : Exit Sub

        TextBox3.Text = Format(Val(Format(TextBox3.Text, "General Number")) - Val(Format(sImporte, "General Number")), "Currency")
        TextBox6.Text = Format((Val(Format(TextBox3.Text, "General Number")) + Val(Format(TextBox4.Text, "General Number"))) - Val(Format(TextBox9.Text, "General Number")), "Currency")

        FlowLayoutPanel2.Controls.Clear()

        BotonesHabitacion()

    End Sub
    Private Sub RadTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles RadTimePicker1.ValueChanged
        Dim otime As RadTimePicker = sender
        Dim iDsrv As Integer = 0
        Dim lDifer As Long = 0
        Dim dDif As Double = 0
        Dim iTipoPag As Integer = 0

        If IsNothing(otime.Tag) = False Then

            If RadDateTimePicker1.Enabled = True Then Exit Sub

            If RadioButton1.Checked = True Then iTipoPag = 1
            If RadioButton2.Checked = True Then iTipoPag = 2

            Calculaimportes(RadDateTimePicker1, otime, TextBox1, TextBox2, iTipoPag)

        End If

    End Sub

    Private Sub RadDateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles RadDateTimePicker1.ValueChanged
        Dim oDate As RadDateTimePicker = sender
        Dim iTipopag As Integer = 0

        If IsNothing(oDate.Tag) = True Then Exit Sub
        If oDate.Tag.ToString.Length = 0 Then Exit Sub

        If RadioButton1.Checked = True Then iTipoPag = 1
        If RadioButton2.Checked = True Then iTipoPag = 2

        Calculaimportes(oDate, RadTimePicker1, TextBox1, TextBox2, iTipoPag)

    End Sub

    Sub CalcularExc(lDifer As Long, sTipo As String, sMontoex As String)
        Dim sMonto As String = String.Empty

        If sMontoex.Trim.Length = 0 Then Exit Sub

        Select Case sTipo.Trim.ToUpper
            Case "E"
                sMonto = Mid(sMontoex.Trim, 1, InStr(sMontoex.Trim, "-", CompareMethod.Text) - 1)
            Case "T"
                sMonto = Mid(sMontoex.Trim, InStr(sMontoex.Trim, "-", CompareMethod.Text) + 1, sMontoex.Trim.Length)
        End Select
        TextBox1.Text = Format(Val(sMonto) * lDifer, "Currency")
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        Dim oT As TextBox = sender

        If Contar_caracteres(TextBox11.Text, " ") >= 2 Then
            tbClientes.Select("cliente = '" & TextBox11.Text.Trim & "'")

            TextBox12.Text = tbClientes.Rows(0).Field(Of String)("direccion")
            TextBox13.Text = tbClientes.Rows(0).Field(Of String)("poblacion")
            TextBox14.Text = tbClientes.Rows(0).Field(Of String)("estado")
            TextBox15.Text = tbClientes.Rows(0).Field(Of String)("cpostal")
            TextBox16.Text = tbClientes.Rows(0).Field(Of String)("email")
            TextBox17.Text = tbClientes.Rows(0).Field(Of String)("telefono")

        End If
    End Sub

    Private Sub RadButton6_Click(sender As Object, e As EventArgs) Handles RadButton6.Click
        Dim iConteo As Integer = 0
        Dim aRadto() As RadToggleButton = {RadToggleButton1, RadToggleButton2, RadToggleButton3, RadToggleButton4, RadToggleButton5, RadToggleButton6, RadToggleButton7, RadToggleButton8, RadToggleButton9, RadToggleButton10}

        For iConteo = 0 To 9 Step 1
            If aRadto(iConteo).ToggleState = Enumerations.ToggleState.On Then
                aRadto(iConteo).ToggleState = Enumerations.ToggleState.Off
            End If
        Next


        LimpiarControles()
        FlowLayoutPanel1.Controls.Clear()
        FlowLayoutPanel2.Controls.Clear()

        BotonesPrecios()
        BotonesHabitacion()

    End Sub

    Sub Calculaimportes(ofecha As RadDateTimePicker, oHora As RadTimePicker, oT1 As TextBox, oT2 As TextBox, iTP As Integer)
        Dim lDifD As Long = DateDiff("d", ofecha.Tag, ofecha.Value)
        Dim lDifH As Long = DateDiff("h", oHora.Tag, oHora.Value) / 2

        If ofecha.Enabled = True Then

            If lDifD < 0 Then
                RadLabelElement4.Text = "No debe elegir una fecha anterior al actual...."
                MsgBox("Deberá dar una fecha de la actual hacia adelante para dar el servicio...", vbCritical, "Cuidado...") : Exit Sub
            ElseIf lDifD > 0 Then
                Select Case iTP
                    Case 1
                        TextBox2.Text = Format(Val(RadioButton1.Tag) * (lDifD + IIf(lDifD >= 1, 1, 0)), "Currency")
                        CalcularExc(lDifD, "E", TextBox1.Tag)
                    Case 2
                        TextBox2.Text = Format(Val(RadioButton2.Tag) * (lDifD + IIf(lDifD >= 1, 1, 0)), "Currency")
                        CalcularExc(lDifD, "T", TextBox1.Tag)
                End Select
            End If
        Else
            If lDifH > 0 Then

                If InStr(lDifH.ToString.Trim, ".", CompareMethod.Text) > 0 Then
                    RadLabelElement4.Text = "No es un tiempo válido para rentar, aumente su tiempo..."
                Else
                    Select Case iTP
                        Case 1
                            TextBox2.Text = Format(Val(RadioButton1.Tag) * (lDifH + 1), "Currency")
                            CalcularExc(lDifH + 1, "E", TextBox1.Tag)
                        Case 2
                            TextBox2.Text = Format(Val(RadioButton2.Tag) * (lDifH + 1), "Currency")
                            CalcularExc(lDifH + 1, "T", TextBox1.Tag)
                    End Select
                    RadLabelElement4.Text = String.Empty
                End If
            ElseIf lDifH < 0 Then
                oHora.Value = DateAdd("d", 1, oHora.Value)
                ofecha.Value = oHora.Value

            End If
        End If

    End Sub

    Private Sub RadLiveTileElement3_Click(sender As Object, e As EventArgs) Handles RadLiveTileElement3.Click
        Bloqueopantprincipal(2)
    End Sub

    Private Sub RadLiveTileElement5_Click(sender As Object, e As EventArgs) Handles RadLiveTileElement5.Click
        Bloqueopantprincipal(4)
    End Sub
End Class
