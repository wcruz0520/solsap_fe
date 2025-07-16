Imports System.Xml
Imports System.IO
Imports System.Windows.Forms

Public Class frmGeneradorATS

    Private oForm As SAPbouiCOM.Form
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application



    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioGeneradorATS()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmGeneradorATS") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmGeneradorATS.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmGeneradorATS").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmGeneradorATS")
            oForm.Freeze(True)

            Dim imgSRI As SAPbouiCOM.PictureBox
            imgSRI = oForm.Items.Item("imgSRI").Specific
            imgSRI.Picture = Application.StartupPath & "\SRI.jpg"

            llena_comboAnio()
            llena_comboMes()


            oForm.Visible = True
            oForm.Freeze(False)
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent

        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                   And pVal.FormTypeEx = "frmGeneradorATS" Then
            If pVal.BeforeAction = False And pVal.ItemUID = "btnIngre" Then

                oForm = rsboApp.Forms.Item("frmGeneradorATS")

                Dim txtClave As SAPbouiCOM.EditText
                txtClave = oForm.Items.Item("txtClave").Specific
                If txtClave.Value.ToString().Equals("S0ls@p2o1f") Then
                    ofrmConfMenu.CargaFormularioMenuDeConfiguraciones()
                    oForm.Close()
                Else
                    rsboApp.StatusBar.SetText(NombreAddon + " - Clave Incorrecta!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If

            ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnGenerar" Then

                oForm = rsboApp.Forms.Item("frmGeneradorATS")

                Dim cboAnio As SAPbouiCOM.ComboBox
                cboAnio = oForm.Items.Item("cboAnio").Specific
                Dim cboMes As SAPbouiCOM.ComboBox
                cboMes = oForm.Items.Item("cboMes").Specific

                Dim anio As Integer = CInt(cboAnio.Value.ToString)
                Dim mes As Integer = CInt(cboMes.Value.ToString)

                Dim Directorioxmlgen = CreaXML_ATS(anio, mes)


                If Directorioxmlgen <> "" Then

                    Dim resp = rsboApp.MessageBox("El Documento XML se Genero Exitosamente!!", 1, "Abrir Directorio", "Abrir XML", "Cerrar")

                    If resp = 1 Then

                        'System.Diagnostics.Process.Start(System.IO.Path.GetDirectoryName(Directorioxmlgen))
                        Dim link = Directorioxmlgen
                        Dim MiProceso As New System.Diagnostics.Process
                        MiProceso.Start("explorer.exe", System.IO.Path.GetDirectoryName(Directorioxmlgen))

                    ElseIf resp = 2 Then

                        'System.Diagnostics.Process.Start(Directorioxmlgen)
                        Dim Proc As New Process()
                        Proc.StartInfo.FileName = Directorioxmlgen
                        Proc.Start()
                        Proc.Dispose()

                    End If

                End If

            ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnSalir" Then
                oForm = rsboApp.Forms.Item("frmGeneradorATS")
                oForm.Close()
            End If
        End If

    End Sub


    Private Function RecorreFormulario(ByVal oApp As SAPbouiCOM.Application, ByVal Formulario As String) As Boolean
        Try
            For Each oForm In oApp.Forms
                Select Case oForm.UniqueID
                    Case Formulario
                        oForm.Visible = True
                        oForm.Select()
                        Return True
                End Select
            Next

            For Each oForm In oApp.Forms
                If oForm.UniqueID = Formulario Then
                    oForm.Visible = True
                    oForm.Select()
                    ' oForm.Close()
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Private Sub llena_comboAnio()

        Dim queryAnio As String = "select * from ""@SS_ANIO"""
        'Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing

        oForm = rsboApp.Forms.Item("frmGeneradorATS")
        Dim cboAnio As SAPbouiCOM.ComboBox
        cboAnio = oForm.Items.Item("cboAnio").Specific

        Dim oRecordSet As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet("select * from ""@SS_ANIO"" order by ""Code""")
        ValoresValidos = cboAnio.ValidValues
        While cboAnio.ValidValues.Count > 0
            cboAnio.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End While
        If oRecordSet.RecordCount > 1 Then
            While (oRecordSet.EoF = False)
                'rsboApp.SetStatusBarMessage("Valor" + oRecordSet.Fields.Item("Code").Value.ToString + " descripcion: " + oRecordSet.Fields.Item("Name").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                ValoresValidos.Add(Convert.ToString(oRecordSet.Fields.Item("Code").Value), Convert.ToString(oRecordSet.Fields.Item("Name").Value))
                oRecordSet.MoveNext()
            End While
        End If



    End Sub
    Private Sub llena_comboMes()

        Dim queryAnio As String = "select * from ""@SS_MES"""
        'Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing

        oForm = rsboApp.Forms.Item("frmGeneradorATS")
        Dim cboMes As SAPbouiCOM.ComboBox
        cboMes = oForm.Items.Item("cboMes").Specific

        Dim oRecordSet As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet("select * from ""@SS_MES"" order by cast(""Code"" as int) asc ")
        ValoresValidos = cboMes.ValidValues
        While cboMes.ValidValues.Count > 0
            cboMes.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End While
        If oRecordSet.RecordCount > 1 Then
            While (oRecordSet.EoF = False)
                'rsboApp.SetStatusBarMessage("Valor" + oRecordSet.Fields.Item("Code").Value.ToString + " descripcion: " + oRecordSet.Fields.Item("Name").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                ValoresValidos.Add(Convert.ToString(oRecordSet.Fields.Item("Code").Value), Convert.ToString(oRecordSet.Fields.Item("Name").Value))
                oRecordSet.MoveNext()
            End While
        End If
    End Sub

    Private Function CreaXML_ATS(anio As Integer, mes As Integer) As String


        Try

            Dim nom As String = "ATS-" + mes.ToString().Trim() + "-" + anio.ToString().Trim() + ".xml"


            Dim selectFileDialog As New SelectFileDialog("C:\", "", "|*", DialogType.FOLDER)
            selectFileDialog.Open()

            If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFolder) Then

                GenerarXML_ATs(selectFileDialog.SelectedFolder & "\" & nom, anio, mes)

                rsboApp.StatusBar.SetText(NombreAddon + " - Archivo Generado correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Return selectFileDialog.SelectedFolder & "\" & nom

            End If



        Catch ex As Exception

            Utilitario.Util_Log.Escribir_Log("Error al generar ATS " & ex.Message, "GeneradorATS")

        End Try


        Return String.Empty


    End Function


    Private Sub GenerarXML_ATs(ByVal Ruta As String, ByVal anio As Integer, ByVal mes As Integer)

        Dim archivo As TextWriter = New StreamWriter(Ruta)
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        'System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator = "."

        Try
            Dim DocEntryDocComp As String = ""
            Dim DocEntryDocComp2 As String = ""
            Dim Sql As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Sql = "call " & rCompany.CompanyDB & ".SBOSP_CABCIA(" + anio.ToString + "," + mes.ToString + ",'' ) "
            Else
                Sql = "EXEC SBOSP_CABCIA " + anio.ToString + "," + mes.ToString + ",'' "
            End If
            Utilitario.Util_Log.Escribir_Log("query SBOSP_CABCIA: " & Sql.ToString(), "frmGeneradorATS")
            Dim rs As SAPbobsCOM.Recordset
            rs = oFuncionesB1.getRecordSet(Sql)

            archivo.WriteLine("<?xml version='1.0' encoding='UTF-8'?>")
            archivo.WriteLine("<iva>")
            archivo.WriteLine(" <TipoIDInformante>" & rs.Fields.Item("TipoIDInformante").Value.ToString() & "</TipoIDInformante>")
            archivo.WriteLine(" <IdInformante>" & rs.Fields.Item("IdInformante").Value.ToString() & "</IdInformante>")
            archivo.WriteLine(" <razonSocial>" & rs.Fields.Item("razonSocial").Value.ToString() & "</razonSocial>")
            archivo.WriteLine(" <Anio>" & rs.Fields.Item("Anio").Value.ToString() & "</Anio>")
            archivo.WriteLine(" <Mes>" & rs.Fields.Item("Mes").Value.ToString() & "</Mes>")
            archivo.WriteLine(" <numEstabRuc>" & rs.Fields.Item("numEstabRuc").Value.ToString() & "</numEstabRuc>")
            archivo.WriteLine(" <totalVentas>" & Convert.ToDecimal(rs.Fields.Item("totalVentas").Value.ToString()).ToString("###0.00") & "</totalVentas>")
            'archivo.WriteLine(" <totalVentas>" & String.Format("{0:#,0.####}", rs.Fields.Item("totalVentas").Value.ToString()) & "</totalVentas>")
            archivo.WriteLine(" <codigoOperativo>" & rs.Fields.Item("codigoOperativo").Value.ToString() & "</codigoOperativo>")
            archivo.WriteLine(" <compras>")

            Dim Sql2 As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Sql2 = "call " & rCompany.CompanyDB & ".SBOSP_DET_COMPRAS(" + anio.ToString + "," + mes.ToString + ",'' ) "
            Else
                Sql2 = "EXEC SBOSP_DET_COMPRAS " + anio.ToString + "," + mes.ToString + ",'' "
            End If
            Utilitario.Util_Log.Escribir_Log("query SBOSP_DET_COMPRAS: " & Sql2.ToString(), "frmGeneradorATS")
            Dim rs2 As SAPbobsCOM.Recordset
            rs2 = oFuncionesB1.getRecordSet(Sql2)

            rsboApp.StatusBar.SetText(NombreAddon + " - Consultando Compras.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            If rs2.RecordCount > 0 Then
                While (rs2.EoF = False)
                    DocEntryDocComp = rs2.Fields.Item("DocEntry").Value.ToString()
                    If DocEntryDocComp <> DocEntryDocComp2 Then

                        archivo.WriteLine("     <detalleCompras>")
                        archivo.WriteLine("         <codSustento>" + rs2.Fields.Item("codSustento").Value.ToString() + "</codSustento>")
                        archivo.WriteLine("         <tpIdProv>" + rs2.Fields.Item("tpIdProv").Value.ToString() + "</tpIdProv>")
                        archivo.WriteLine("         <idProv>" + rs2.Fields.Item("idProv").Value.ToString() + "</idProv>")
                        archivo.WriteLine("         <tipoComprobante>" + rs2.Fields.Item("tipoComprobante").Value.ToString() + "</tipoComprobante>")
                        If rs2.Fields.Item("tpIdProv").Value.ToString() = "03" Then

                            archivo.WriteLine("         <tipoProv>" + rs2.Fields.Item("tipoProv").Value.ToString() + "</tipoProv>")
                            archivo.WriteLine("         <denoProv>" + rs2.Fields.Item("denopr").Value.ToString() + "</denoProv>")

                        End If
                        archivo.WriteLine("         <parteRel>NO</parteRel>")
                        archivo.WriteLine("         <fechaRegistro>" + Convert.ToDateTime(rs2.Fields.Item("fechaRegistro").Value.ToString()).ToString("dd/MM/yyyy") + "</fechaRegistro>")
                        archivo.WriteLine("         <establecimiento>" + rs2.Fields.Item("establecimiento").Value.ToString() + "</establecimiento>")
                        archivo.WriteLine("         <puntoEmision>" + rs2.Fields.Item("puntoEmision").Value.ToString() + "</puntoEmision>")
                        archivo.WriteLine("         <secuencial>" + rs2.Fields.Item("secuencial").Value.ToString() + "</secuencial>")
                        archivo.WriteLine("         <fechaEmision>" + Convert.ToDateTime(rs2.Fields.Item("fechaEmision").Value.ToString()).ToString("dd/MM/yyyy") + "</fechaEmision>")
                        archivo.WriteLine("         <autorizacion>" + rs2.Fields.Item("autorizacion").Value.ToString() + "</autorizacion>")
                        archivo.WriteLine("         <baseNoGraIva>" + Convert.ToDecimal(rs2.Fields.Item("baseNoGraIva").Value.ToString()).ToString("###0.00") + "</baseNoGraIva>")
                        archivo.WriteLine("         <baseImponible>" + Convert.ToDecimal(rs2.Fields.Item("baseImponible").Value.ToString()).ToString("###0.00") + "</baseImponible>")
                        archivo.WriteLine("         <baseImpGrav>" + Convert.ToDecimal(rs2.Fields.Item("baseImpGrav").Value.ToString()).ToString("###0.00") + "</baseImpGrav>")
                        archivo.WriteLine("         <baseImpExe>0.00</baseImpExe>")
                        archivo.WriteLine("         <montoIce>0.00</montoIce>")
                        archivo.WriteLine("         <montoIva>" + Convert.ToDecimal(rs2.Fields.Item("montoIva").Value.ToString()).ToString("###0.00") + "</montoIva>")
                        archivo.WriteLine("         <valRetBien10>" + Convert.ToDecimal(rs2.Fields.Item("valRetBien10").Value.ToString()).ToString("###0.00") + "</valRetBien10>")
                        archivo.WriteLine("         <valRetServ20>" + Convert.ToDecimal(rs2.Fields.Item("valRetServ20").Value.ToString()).ToString("###0.00") + "</valRetServ20>")
                        archivo.WriteLine("         <valorRetBienes>" + Convert.ToDecimal(rs2.Fields.Item("valorRetBienes").Value.ToString()).ToString("###0.00") + "</valorRetBienes>")
                        archivo.WriteLine("         <valRetServ50>" + Convert.ToDecimal(rs2.Fields.Item("valRetServ50").Value.ToString()).ToString("###0.00") + "</valRetServ50>")
                        archivo.WriteLine("         <valorRetServicios>" + Convert.ToDecimal(rs2.Fields.Item("valorRetServicios").Value.ToString()).ToString("###0.00") + "</valorRetServicios>")
                        archivo.WriteLine("         <valRetServ100>" + Convert.ToDecimal(rs2.Fields.Item("valRetServ100").Value.ToString()).ToString("###0.00") + "</valRetServ100>")
                        archivo.WriteLine("         <totbasesImpReemb>" + Convert.ToDecimal(rs2.Fields.Item("totbasesImpReemb").Value.ToString()).ToString("###0.00") + "</totbasesImpReemb>")




                        'archivo.WriteLine("<valorRetBienes>" + rs2.Fields.Item("valorRetBienes").Value.ToString() + "</valorRetBienes>")
                        'archivo.WriteLine("<valorRetServicios>" + rs2.Fields.Item("valorRetServicios").Value.ToString() + "</valorRetServicios>")
                        'archivo.WriteLine("<valRetServ100>" + rs2.Fields.Item("valRetServ100").Value.ToString() + "</valRetServ100>")

                        archivo.WriteLine("         <pagoExterior>")
                        archivo.WriteLine("             <pagoLocExt>" + rs2.Fields.Item("pagoLocExt").Value.ToString() + "</pagoLocExt>")
                        archivo.WriteLine("             <tipoRegi>" + rs2.Fields.Item("tipoRegi").Value.ToString() + "</tipoRegi>")
                        If rs2.Fields.Item("tipoRegi").Value.ToString() = "01" Then
                            archivo.WriteLine("             <paisEfecPagoGen>" + rs2.Fields.Item("paisEfecPagoGen").Value.ToString() + "</paisEfecPagoGen>")
                        End If
                        If rs2.Fields.Item("tipoRegi").Value.ToString() = "02" Then
                            archivo.WriteLine("             <paisEfecPagoParFis>" + rs2.Fields.Item("paisEfecPagoParFis").Value.ToString() + "</paisEfecPagoParFis>")
                        End If
                        archivo.WriteLine("             <paisEfecPago>" + rs2.Fields.Item("paisEfecPago").Value.ToString() + "</paisEfecPago>")
                        archivo.WriteLine("             <aplicConvDobTrib>" + rs2.Fields.Item("aplicConvDobTrib").Value.ToString() + "</aplicConvDobTrib>")
                        archivo.WriteLine("             <pagExtSujRetNorLeg>" + rs2.Fields.Item("pagExtSujRetNorLeg").Value.ToString() + "</pagExtSujRetNorLeg>")
                        archivo.WriteLine("         </pagoExterior>")
                    End If


                    If (Double.Parse(rs2.Fields.Item("DocTotal").Value.ToString()) >= 1000) And CInt(anio.ToString + Right(("0" + mes.ToString), 2)) <= 202312 Then
                        archivo.WriteLine("         <formasDePago>")
                        archivo.WriteLine("             <formaPago>" + rs2.Fields.Item("formaPago").Value.ToString() + "</formaPago>")
                        archivo.WriteLine("         </formasDePago>")

                    ElseIf (Double.Parse(rs2.Fields.Item("DocTotal").Value.ToString()) >= 500) And CInt(anio.ToString + Right(("0" + mes.ToString), 2)) >= 202312 Then
                        archivo.WriteLine("         <formasDePago>")
                        archivo.WriteLine("             <formaPago>" + rs2.Fields.Item("formaPago").Value.ToString() + "</formaPago>")
                        archivo.WriteLine("         </formasDePago>")


                    End If


                    'inicia Air

                    Dim CodCompra As Integer = Integer.Parse(rs2.Fields.Item("DocEntry").Value.ToString())
                    Dim tipoComp As String = rs2.Fields.Item("tipoComprobante").Value.ToString()

                    If tipoComp = "01" Or tipoComp = "03" Or tipoComp = "41" Or tipoComp = "02" Or tipoComp = "12" Or tipoComp = "19" Or tipoComp = "20" Then
                        If (CodCompra > 0) Then 'une Then las retenciones de la compra

                            If tipoComp <> "41" Then

                                'End If

                                Dim Sql3 As String = ""
                                If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    Sql3 = "call " & rCompany.CompanyDB & ".SBOSP_AIR(" + CodCompra.ToString + ",'' ) "
                                Else
                                    Sql3 = "EXEC SBOSP_AIR " + CodCompra.ToString + ",'' "
                                End If
                                Utilitario.Util_Log.Escribir_Log("query SBOSP_AIR: " & Sql3.ToString(), "frmGeneradorATS")
                                Dim rs3 As SAPbobsCOM.Recordset
                                rs3 = oFuncionesB1.getRecordSet(Sql3)

                                archivo.WriteLine("         <air>")

                                If rs3.RecordCount > 0 Then

                                    While (rs3.EoF = False)

                                        archivo.WriteLine("             <detalleAir>")
                                        archivo.WriteLine("                 <codRetAir>" + rs3.Fields.Item("codRetAir").Value.ToString() + "</codRetAir>")
                                        archivo.WriteLine("                 <baseImpAir>" + Convert.ToDecimal(rs3.Fields.Item("baseImpAir").Value.ToString()).ToString("###0.00") + "</baseImpAir>")
                                        archivo.WriteLine("                 <porcentajeAir>" + rs3.Fields.Item("porcentajeAir").Value.ToString() + "</porcentajeAir>")
                                        archivo.WriteLine("                 <valRetAir>" + Convert.ToDecimal(rs3.Fields.Item("valRetAir").Value.ToString()).ToString("###0.00") + "</valRetAir>")
                                        archivo.WriteLine("             </detalleAir>")

                                        rs3.MoveNext()

                                    End While
                                End If

                                archivo.WriteLine("         </air>")
                                oFuncionesB1.Release(rs3)
                            End If
                        End If


                        If (rs2.Fields.Item("secRetencion1").Value.ToString <> "") Then

                            archivo.WriteLine("         <estabRetencion1>" + rs2.Fields.Item("estabRetencion1").Value.ToString() + "</estabRetencion1>")
                            archivo.WriteLine("         <ptoEmiRetencion1>" + rs2.Fields.Item("ptoEmiRetencion1").Value.ToString() + "</ptoEmiRetencion1>")
                            archivo.WriteLine("         <secRetencion1>" + rs2.Fields.Item("secRetencion1").Value.ToString() + "</secRetencion1>")
                            archivo.WriteLine("         <autRetencion1>" + rs2.Fields.Item("autRetencion1").Value.ToString() + "</autRetencion1>")
                            archivo.WriteLine("         <fechaEmiRet1>" + Convert.ToDateTime(rs2.Fields.Item("fechaEmiRet1").Value.ToString()).ToString("dd/MM/yyyy") + "</fechaEmiRet1>")
                        End If

                    End If
                    If tipoComp = "04" Then

                        archivo.WriteLine("         <docModificado>" + rs2.Fields.Item("docModificado").Value.ToString() + "</docModificado>")
                        archivo.WriteLine("         <estabModificado>" + rs2.Fields.Item("estabModificado").Value.ToString() + "</estabModificado>")
                        archivo.WriteLine("         <ptoEmiModificado>" + rs2.Fields.Item("ptoEmiModificado").Value.ToString() + "</ptoEmiModificado>")
                        archivo.WriteLine("         <secModificado>" + rs2.Fields.Item("secModificado").Value.ToString() + "</secModificado>")
                        archivo.WriteLine("         <autModificado>" + rs2.Fields.Item("autModificado").Value.ToString() + "</autModificado>")

                    End If

                    If tipoComp = "41" Then


                        Dim SqlReem As String = ""
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            SqlReem = "call " & rCompany.CompanyDB & ".SBOSP_DET_COMPRAS_REEM(" + DocEntryDocComp + ") "
                        Else
                            SqlReem = "EXEC SBOSP_DET_COMPRAS_REEM " + DocEntryDocComp
                        End If
                        Utilitario.Util_Log.Escribir_Log("query SBOSP_DET_COMPRAS_REEM: " & SqlReem.ToString(), "frmGeneradorATS")
                        Dim rsReem As SAPbobsCOM.Recordset
                        rsReem = oFuncionesB1.getRecordSet(SqlReem)

                        If rsReem.RecordCount > 0 And (DocEntryDocComp <> DocEntryDocComp2) Then
                            archivo.WriteLine("         <reembolsos>")
                            While (rsReem.EoF = False)
                                If Not (rsReem.Fields.Item("tipoComprobanteReemb").Value.ToString) = "" Then

                                    archivo.WriteLine("             <reembolso>")
                                    archivo.WriteLine("                 <tipoComprobanteReemb>" + rsReem.Fields.Item("tipoComprobanteReemb").Value.ToString() + "</tipoComprobanteReemb>")
                                    archivo.WriteLine("                 <tpIdProvReemb>" + rsReem.Fields.Item("tpIdProvReemb").Value.ToString() + "</tpIdProvReemb>")
                                    archivo.WriteLine("                 <idProvReemb>" + rsReem.Fields.Item("idProvReemb").Value.ToString() + "</idProvReemb>")
                                    archivo.WriteLine("                 <establecimientoReemb>" + rsReem.Fields.Item("establecimientoReemb").Value.ToString() + "</establecimientoReemb>")
                                    archivo.WriteLine("                 <puntoEmisionReemb>" + rsReem.Fields.Item("puntoEmisionReemb").Value.ToString() + "</puntoEmisionReemb>")
                                    archivo.WriteLine("                 <secuencialReemb>" + rsReem.Fields.Item("secuencialReemb").Value.ToString() + "</secuencialReemb>")
                                    archivo.WriteLine("                 <fechaEmisionReemb>" + Convert.ToDateTime(rsReem.Fields.Item("fechaEmisionReemb").Value.ToString()).ToString("dd/MM/yyyy") + "</fechaEmisionReemb>")
                                    archivo.WriteLine("                 <autorizacionReemb>" + rsReem.Fields.Item("autorizacionReemb").Value.ToString() + "</autorizacionReemb>")
                                    archivo.WriteLine("                 <baseImponibleReemb>" + Convert.ToDecimal(rsReem.Fields.Item("baseImponibleReemb").Value.ToString()).ToString("###0.00") + "</baseImponibleReemb>")
                                    archivo.WriteLine("                 <baseImpGravReemb>" + Convert.ToDecimal(rsReem.Fields.Item("baseImpGravReemb").Value.ToString()).ToString("###0.00") + "</baseImpGravReemb>")
                                    archivo.WriteLine("                 <baseNoGraIvaReemb>" + Convert.ToDecimal(rsReem.Fields.Item("baseNoGraIvaReemb").Value.ToString()).ToString("###0.00") + "</baseNoGraIvaReemb>")
                                    archivo.WriteLine("                 <baseImpExeReemb>" + Convert.ToDecimal(rsReem.Fields.Item("baseImpExeReemb").Value.ToString()).ToString("###0.00") + "</baseImpExeReemb>")
                                    archivo.WriteLine("                 <montoIceRemb>" + Convert.ToDecimal(rsReem.Fields.Item("montoIceRemb").Value.ToString()).ToString("###0.00") + "</montoIceRemb>")
                                    archivo.WriteLine("                 <montoIvaRemb>" + Convert.ToDecimal(rsReem.Fields.Item("montoIvaRemb").Value.ToString()).ToString("###0.00") + "</montoIvaRemb>")
                                    archivo.WriteLine("             </reembolso>")


                                End If
                                rsReem.MoveNext()
                            End While
                            archivo.WriteLine("         </reembolsos>")
                        End If
                        oFuncionesB1.Release(rsReem)

                    End If
                    archivo.WriteLine("     </detalleCompras>")
                    DocEntryDocComp2 = rs2.Fields.Item("DocEntry").Value.ToString()

                    rs2.MoveNext() 'detalle de compras
                    'DocEntryDocComp2 = rs2.Fields.Item("DocEntry").Value.ToString()
                End While
            End If
            archivo.WriteLine(" </compras>")


            Dim Sql4 As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Sql4 = "call " & rCompany.CompanyDB & ".SBOSP_VENTAS(" + anio.ToString + "," + mes.ToString + ",'' ) "
            Else
                Sql4 = "EXEC SBOSP_VENTAS " + anio.ToString + "," + mes.ToString + ",'' "
            End If
            Utilitario.Util_Log.Escribir_Log("query SBOSP_VENTAS: " & Sql4.ToString(), "frmGeneradorATS")
            Dim rs4 As SAPbobsCOM.Recordset
            rs4 = oFuncionesB1.getRecordSet(Sql4)

            rsboApp.StatusBar.SetText(NombreAddon + " - Consultando ventas.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            archivo.WriteLine(" <ventas>")

            If rs4.RecordCount > 0 Then
                While (rs4.EoF = False)
                    archivo.WriteLine("     <detalleVentas>")
                    archivo.WriteLine("         <tpIdCliente>" + rs4.Fields.Item("tpIdCliente").Value.ToString() + "</tpIdCliente>")
                    archivo.WriteLine("         <idCliente>" + rs4.Fields.Item("idCliente").Value.ToString() + "</idCliente>")

                    '<parteRelVtas>NO</parteRelVtas>
                    If rs4.Fields.Item("tpIdCliente").Value.ToString() <> "07" Then
                        archivo.WriteLine("         <parteRelVtas>NO</parteRelVtas>")
                    End If

                    If rs4.Fields.Item("tpIdCliente").Value.ToString() = "06" Then

                        archivo.WriteLine("         <tipoCliente>" + rs4.Fields.Item("tipoCliente").Value.ToString() + "</tipoCliente>")
                        archivo.WriteLine("         <denoCli>" + rs4.Fields.Item("denoCli").Value.ToString() + "</denoCli>")

                    End If


                    archivo.WriteLine("         <tipoComprobante>" + rs4.Fields.Item("tipoComprobante").Value.ToString() + "</tipoComprobante>")
                    archivo.WriteLine("         <tipoEmision>" + rs4.Fields.Item("tipoEmision").Value.ToString() + "</tipoEmision>")
                    archivo.WriteLine("         <numeroComprobantes>" + rs4.Fields.Item("numeroComprobantes").Value.ToString() + "</numeroComprobantes>")
                    archivo.WriteLine("         <baseNoGraIva>" + Convert.ToDecimal(rs4.Fields.Item("baseNoGraIva").Value.ToString()).ToString("###0.00") + "</baseNoGraIva>")
                    archivo.WriteLine("         <baseImponible>" + Convert.ToDecimal(rs4.Fields.Item("baseImponible").Value.ToString()).ToString("###0.00") + "</baseImponible>")
                    archivo.WriteLine("         <baseImpGrav>" + Convert.ToDecimal(rs4.Fields.Item("baseImpGrav").Value.ToString()).ToString("###0.00") + "</baseImpGrav>")
                    'archivo.WriteLine("         <baseImpExe>" + Convert.ToDecimal(rs4.Fields.Item("baseImpExe").Value.ToString()).ToString("###0.00") + "</baseImpExe>")
                    archivo.WriteLine("         <montoIva>" + Convert.ToDecimal(rs4.Fields.Item("montoIva").Value.ToString()).ToString("###0.00") + "</montoIva>")

                    '<montoIce> 0 </montoIce>
                    archivo.WriteLine("         <montoIce>0.00</montoIce>")
                    archivo.WriteLine("         <valorRetIva>" + Convert.ToDecimal(rs4.Fields.Item("valorRetIva").Value.ToString()).ToString("###0.00") + "</valorRetIva>")
                    archivo.WriteLine("         <valorRetRenta>" + Convert.ToDecimal(rs4.Fields.Item("valorRetRenta").Value.ToString()).ToString("###0.00") + "</valorRetRenta>")


                    If (rs4.Fields.Item("tipoComprobante").Value.ToString() = "18") Or (rs4.Fields.Item("tipoComprobante").Value.ToString() = "05") Or (rs4.Fields.Item("tipoComprobante").Value.ToString() = "41") Then

                        archivo.WriteLine("         <formasDePago>")
                        archivo.WriteLine("             <formaPago>20</formaPago>")
                        archivo.WriteLine("         </formasDePago>")

                    End If

                    archivo.WriteLine("     </detalleVentas>")
                    rs4.MoveNext()
                End While
            End If

            archivo.WriteLine(" </ventas>")

            Dim Sql6 As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Sql6 = "call " & rCompany.CompanyDB & ".SBOSP_VENEST(" + anio.ToString + "," + mes.ToString + ",'' ) "
            Else
                Sql6 = "EXEC SBOSP_VENEST " + anio.ToString + "," + mes.ToString + ",'' "
            End If
            Utilitario.Util_Log.Escribir_Log("query SBOSP_VENEST: " & Sql6.ToString(), "frmGeneradorATS")
            Dim rs6 As SAPbobsCOM.Recordset
            rs6 = oFuncionesB1.getRecordSet(Sql6)

            rsboApp.StatusBar.SetText(NombreAddon + " - Consultando ventas establecimiento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            archivo.WriteLine(" <ventasEstablecimiento>")

            If rs6.RecordCount > 0 Then
                While (rs6.EoF = False)
                    archivo.WriteLine("     <ventaEst>")
                    archivo.WriteLine("         <codEstab>" + rs6.Fields.Item("codEstab").Value.ToString() + "</codEstab>")
                    archivo.WriteLine("         <ventasEstab>" + Convert.ToDecimal(rs6.Fields.Item("ventasEstab").Value.ToString()).ToString("###0.00") + "</ventasEstab>")
                    archivo.WriteLine("         <ivaComp>" + Convert.ToDecimal(rs6.Fields.Item("ivaComp").Value.ToString()).ToString("###0.00") + "</ivaComp>")
                    archivo.WriteLine("     </ventaEst>")
                    rs6.MoveNext()
                End While
            End If

            archivo.WriteLine(" </ventasEstablecimiento>")



            Dim Sql5 As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Sql5 = "call " & rCompany.CompanyDB & ".SBOSP_DET_EXPORTACIONES(" + anio.ToString + "," + mes.ToString + ",'' ) "
            Else
                Sql5 = "EXEC SBOSP_DET_EXPORTACIONES " + anio.ToString + "," + mes.ToString + ",'' "
            End If
            Utilitario.Util_Log.Escribir_Log("query SBOSP_DET_EXPORTACIONES: " & Sql5.ToString(), "frmGeneradorATS")
            Dim rs5 As SAPbobsCOM.Recordset
            rs5 = oFuncionesB1.getRecordSet(Sql5)

            rsboApp.StatusBar.SetText(NombreAddon + " - Consultando exportaciones.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            If rs5.RecordCount > 0 Then
                archivo.WriteLine(" <exportaciones>")
                While (rs5.EoF = False)
                    archivo.WriteLine("     <detalleExportaciones>")
                    archivo.WriteLine("         <tpIdClienteEx>" + rs5.Fields.Item("tpIdClienteEx").Value.ToString() + "</tpIdClienteEx>")
                    archivo.WriteLine("         <idClienteEx>" + rs5.Fields.Item("idClienteEx").Value.ToString() + "</idClienteEx>")

                    '<parteRelVtas>NO</parteRelVtas>
                    archivo.WriteLine("         <parteRelExp>" + rs5.Fields.Item("parteRelExp").Value.ToString() + "</parteRelExp>")
                    archivo.WriteLine("         <tipoCli>" + rs5.Fields.Item("tipoCli").Value.ToString() + "</tipoCli>")
                    archivo.WriteLine("         <denoExpCli>" + rs5.Fields.Item("denoExpCli").Value.ToString() + "</denoExpCli>")
                    archivo.WriteLine("         <tipoRegi>" + rs5.Fields.Item("tipoRegi").Value.ToString() + "</tipoRegi>")
                    archivo.WriteLine("         <paisEfecPagoGen>" + rs5.Fields.Item("paisEfecPagoGen").Value.ToString() + "</paisEfecPagoGen>")

                    If Not rs5.Fields.Item("paisEfecPagoParFis").Value.ToString() = "" Then
                        archivo.WriteLine("         <paisEfecPagoParFis>" + rs5.Fields.Item("paisEfecPagoParFis").Value.ToString() + "</paisEfecPagoParFis>")
                    End If


                    archivo.WriteLine("         <paisEfecExp>" + rs5.Fields.Item("paisEfecExp").Value.ToString() + "</paisEfecExp>")
                    archivo.WriteLine("         <exportacionDe>" + rs5.Fields.Item("exportacionDe").Value.ToString() + "</exportacionDe>")
                    archivo.WriteLine("         <tipoComprobante>" + rs5.Fields.Item("tipoComprobante").Value.ToString() + "</tipoComprobante>")

                    If rs5.Fields.Item("exportacionDe").Value.ToString() = "01" Then

                        archivo.WriteLine("         <distAduanero>" + rs5.Fields.Item("distAduanero").Value.ToString() + "</distAduanero>")
                        archivo.WriteLine("         <anio>" + rs5.Fields.Item("anio").Value.ToString() + "</anio>")
                        archivo.WriteLine("         <regimen>" + rs5.Fields.Item("regimen").Value.ToString() + "</regimen>")
                        archivo.WriteLine("         <correlativo>" + rs5.Fields.Item("correlativo").Value.ToString() + "</correlativo>")
                        archivo.WriteLine("         <docTransp>" + rs5.Fields.Item("docTransp").Value.ToString() + "</docTransp>")

                    End If
                    'archivo.WriteLine("             <tipIngExt>" + rs5.Fields.Item("tipIngExt").Value.ToString() + "</tipIngExt>")
                    'archivo.WriteLine("             <ingExtGravOtroPais>" + rs5.Fields.Item("ingExtGravOtroPais").Value.ToString() + "</ingExtGravOtroPais>")
                    archivo.WriteLine("         <fechaEmbarque>" + rs5.Fields.Item("fechaEmbarque").Value.ToString() + "</fechaEmbarque>")
                    archivo.WriteLine("         <valorFOB>" + Convert.ToDecimal(rs5.Fields.Item("valorFOB").Value.ToString()).ToString("###0.00") + "</valorFOB>")
                    archivo.WriteLine("         <valorFOBComprobante>" + Convert.ToDecimal(rs5.Fields.Item("valorFOBComprobante").Value.ToString()).ToString("###0.00") + "</valorFOBComprobante>")
                    archivo.WriteLine("         <establecimiento>" + rs5.Fields.Item("establecimiento").Value.ToString() + "</establecimiento>")
                    archivo.WriteLine("         <puntoEmision>" + rs5.Fields.Item("puntoEmision").Value.ToString() + "</puntoEmision>")
                    archivo.WriteLine("         <secuencial>" + rs5.Fields.Item("secuencial").Value.ToString() + "</secuencial>")
                    archivo.WriteLine("         <autorizacion>" + rs5.Fields.Item("autorizacion").Value.ToString() + "</autorizacion>")
                    archivo.WriteLine("         <fechaEmision>" + rs5.Fields.Item("fechaEmision").Value.ToString() + "</fechaEmision>")

                    archivo.WriteLine("     </detalleExportaciones>")
                    rs5.MoveNext()
                End While
                archivo.WriteLine(" </exportaciones>")
            End If




            Dim Sql7 As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Sql7 = "call " & rCompany.CompanyDB & ".SBOSP_DET_DOCANULADOS(" + anio.ToString + "," + mes.ToString + ",'' ) "
            Else
                Sql7 = "EXEC SBOSP_DET_DOCANULADOS " + anio.ToString + "," + mes.ToString + ",'' "
            End If
            Utilitario.Util_Log.Escribir_Log("query SBOSP_DET_DOCANULADOS: " & Sql7.ToString(), "frmGeneradorATS")
            Dim rs7 As SAPbobsCOM.Recordset
            rs7 = oFuncionesB1.getRecordSet(Sql7)

            rsboApp.StatusBar.SetText(NombreAddon + " - Consultando Documentos Anulados.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            If rs7.RecordCount > 0 Then
                archivo.WriteLine(" <anulados>")
                While (rs7.EoF = False)
                    archivo.WriteLine("     <detalleAnulados>")
                    archivo.WriteLine("         <tipoComprobante>" + rs7.Fields.Item("tipoComprobante").Value.ToString() + "</tipoComprobante>")
                    archivo.WriteLine("         <establecimiento>" + rs7.Fields.Item("establecimiento").Value.ToString() + "</establecimiento>")
                    archivo.WriteLine("         <puntoEmision>" + rs7.Fields.Item("puntoEmision").Value.ToString() + "</puntoEmision>")
                    archivo.WriteLine("         <secuencialInicio>" + rs7.Fields.Item("secuencialInicio").Value.ToString() + "</secuencialInicio>")
                    archivo.WriteLine("         <secuencialFin>" + rs7.Fields.Item("secuencialFin").Value.ToString() + "</secuencialFin>")
                    archivo.WriteLine("         <autorizacion>" + rs7.Fields.Item("autorizacion").Value.ToString() + "</autorizacion>")
                    archivo.WriteLine("     </detalleAnulados>")

                    rs7.MoveNext()
                End While
                archivo.WriteLine(" </anulados>")
            End If


            archivo.WriteLine("</iva>")


            archivo.Close()


            oFuncionesB1.Release(rs)
            oFuncionesB1.Release(rs2)

            oFuncionesB1.Release(rs4)
            oFuncionesB1.Release(rs5)
            oFuncionesB1.Release(rs6)
            oFuncionesB1.Release(rs7)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("error al generar ats: " & ex.Message.ToString(), "frmGeneradorATS")
            rsboApp.StatusBar.SetText(NombreAddon + " - Error al generar ATS..!" & ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'archivo.Close()
        End Try

    End Sub


End Class
