Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports Spire.Pdf
Imports System.Drawing.Printing
Imports System.IO
Public Class frmImprimir
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Dim fila As Integer = 0
    Dim DE As Integer = 0
    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioImprimir(ByVal DocEntry As Integer)
        Try
            Dim xmlDoc As New Xml.XmlDocument
            Dim strPath As String

            If RecorreFormulario(rsboApp, "frmImprimir") Then Exit Sub

            strPath = System.Windows.Forms.Application.StartupPath & "\frmImprimir.srf"
            xmlDoc.Load(strPath)

            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmImprimir").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmImprimir")
            DE = 0

            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTXCH").Specific
            Dim oDBDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_PM_DET1") ' "UDO_D1" es un ejemplo del nombre de la tabla hija

            Dim oConditions As SAPbouiCOM.Conditions = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)
            Dim oCondition As SAPbouiCOM.Condition = oConditions.Add()
            oCondition.Alias = "DocEntry"  ' Campo en la tabla que quieres filtrar
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL  ' Operación de la condición (igual, mayor, menor, etc.)
            oCondition.CondVal = DocEntry.ToString
            oDBDataSource.Query(oConditions)
            oMatrix.LoadFromDataSource()

            For i As Integer = 1 To oMatrix.RowCount
                Dim oCheckBox As SAPbouiCOM.CheckBox = CType(oMatrix.Columns.Item("U_Impreso").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                If Not oCheckBox.Checked Then oMatrix.CommonSetting.SetCellEditable(i, 1, True)
            Next

            DE = DocEntry

            oForm.Visible = True
        Catch ex As Exception

        End Try
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
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "frmImprimir" Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK

                        If Not pVal.BeforeAction Then
                            Select Case pVal.ItemUID
                                Case "MTXCH"
                                    fila = pVal.Row
                                Case "btnImp"
                                    If fila <> 0 Then
                                        Imprimir("Comprobante")
                                    Else
                                        rsboApp.SetStatusBarMessage("Seleccione un numero de cheque! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    End If

                                Case "btnImpCh"
                                    If fila <> 0 Then
                                        Imprimir("Cheque")
                                    Else
                                        rsboApp.SetStatusBarMessage("Seleccione un numero de cheque! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    End If
                            End Select
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function Imprimir(ByVal documento As String)
        Try
            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTXCH").Specific
            Dim DocEntryCh As String = mMatrix.Columns.Item("U_PagTran").Cells.Item(fila).Specific.Value.ToString
            Dim NumLinea As Integer = CInt(mMatrix.Columns.Item("LineId").Cells.Item(fila).Specific.Value.ToString)

            CargaFormato(CInt(DocEntryCh), documento)

            If ActualizaLinea(NumLinea) Then
                mMatrix.CommonSetting.SetCellEditable(fila, 1, False)
                fila = 0
            End If

            'mMatrix.LoadFromDataSource()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub CargaFormato(DocEntry As Integer, Documento As String)
        Try
            Dim menu As Object

            If Documento = "Comprobante" Then

                menu = oFuncionesB1.ObtenerUIDMenu("RptComEgr", Functions.VariablesGlobales._RutaArchivoRPTPM) '"13056")

                If menu <> "" Then

                    For Each f As SAPbouiCOM.Form In rsboApp.Forms
                        If f.TypeEx = "410000100" Then f.Close()
                    Next

                    rsboApp.ActivateMenuItem(menu)

                    Dim forpara As SAPbouiCOM.Form = rsboApp.Forms.GetForm("410000100", 0)
                    forpara.Select()

                    TryCast(forpara.Items.Item("1000003").Specific, SAPbouiCOM.EditText).Value = DocEntry.ToString
                    TryCast(forpara.Items.Item("1").Specific, SAPbouiCOM.Button).Item.Click()
                    forpara.Visible = False

                End If

            ElseIf Documento = "Cheque" Then

                menu = oFuncionesB1.ObtenerUIDMenu("ChqPM", Functions.VariablesGlobales._RutaArchivoRPTPM) '"13056")

                If menu <> "" Then

                    For Each f As SAPbouiCOM.Form In rsboApp.Forms
                        If f.TypeEx = "410000100" Then f.Close()
                    Next

                    rsboApp.ActivateMenuItem(menu)

                    Dim forpara As SAPbouiCOM.Form = rsboApp.Forms.GetForm("410000100", 0)
                    forpara.Select()

                    TryCast(forpara.Items.Item("1000003").Specific, SAPbouiCOM.EditText).Value = DocEntry.ToString
                    TryCast(forpara.Items.Item("1").Specific, SAPbouiCOM.Button).Item.Click()
                    forpara.Visible = False

                End If

            End If
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error Cargando formatos " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function ActualizaLinea(ByVal NumLinea As Integer) As Boolean
        Try
            Dim companyService As SAPbobsCOM.CompanyService = rCompany.GetCompanyService()
            Dim generalService As SAPbobsCOM.GeneralService = companyService.GetGeneralService("SSMTPAGOS")
            Dim generalParams As SAPbobsCOM.GeneralDataParams = generalService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)

            generalParams.SetProperty("DocEntry", DE.ToString)

            Dim generalData As SAPbobsCOM.GeneralData = generalService.GetByParams(generalParams)
            Dim lineCollection As SAPbobsCOM.GeneralDataCollection = generalData.Child("SS_PM_DET1")

            For i As Integer = 0 To lineCollection.Count - 1
                Dim line As SAPbobsCOM.GeneralData = lineCollection.Item(i)

                If CInt(line.GetProperty("LineId")) = NumLinea Then
                    line.SetProperty("U_Impreso", "Y")
                    Exit For
                End If
            Next
            generalService.Update(generalData)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
