Imports SAPbouiCOM

Public Class frmSRI

    Private oForm As SAPbouiCOM.Form
    Private rMatrix As SAPbouiCOM.Matrix
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Public num As Integer = 0


    Dim TipoSerie As SAPbouiCOM.ComboBox = Nothing
    Dim IdSerie As SAPbouiCOM.EditText = Nothing
    Dim NombreSerie As SAPbouiCOM.EditText = Nothing
    Dim UltimoNumeroAntiguoProveedor As SAPbouiCOM.EditText = Nothing

    Dim Prefijo As SAPbouiCOM.EditText = Nothing
    Dim RangoInicial As SAPbouiCOM.EditText = Nothing
    Dim RangoFinal As SAPbouiCOM.EditText = Nothing
    Dim DocNumInicial As SAPbouiCOM.EditText = Nothing
    Dim Resolucion As SAPbouiCOM.EditText = Nothing
    Dim ClaveTecnica As SAPbouiCOM.EditText = Nothing
    Dim FechaInicial As SAPbouiCOM.EditText = Nothing
    Dim FechaFinal As SAPbouiCOM.EditText = Nothing
    Dim Contingencia As SAPbouiCOM.EditText = Nothing

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioSRI()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmSRI") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmSRI.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmSRI").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmSRI")

            'Dim ipLogo As SAPbouiCOM.PictureBox
            'ipLogo = oForm.Items.Item("ipLogo").Specific
            'ipLogo.Picture = System.Windows.Forms.Application.StartupPath & "\imagen_UPD.jpg"

            'Dim flFactura As SAPbouiCOM.Folder
            'flFactura = oForm.Items.Item("Item_1").Specific
            'flFactura.Select()

            oForm.Visible = True
            oForm.Select()


        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla frmSRI: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try

            If pVal.FormUID = "frmSRI" AndAlso pVal.BeforeAction = True AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(pVal.FormUID)
                Dim mEdit As SAPbouiCOM.EditText = Nothing
                Dim mItem As SAPbouiCOM.IItem = Nothing
                Try
                    mItem = mForm.Items.Add("code", BoFormItemTypes.it_EDIT)
                    mItem.Left = 460
                    mItem.Top = 10
                    mItem.Height = 14
                    mItem.Width = 80
                    mItem.Enabled = False
                    mItem.DisplayDesc = False
                    mEdit = mItem.Specific
                    mEdit.DataBind.SetBound(True, "@SS_SER", "code")
                    oFuncionesB1.Release(mItem)
                    mForm.DataBrowser.BrowseBy = "code" 'Next


                Catch ex As Exception
                Finally
                    oFuncionesB1.Release(mForm)
                    oFuncionesB1.Release(mEdit)
                    oFuncionesB1.Release(mItem)
                End Try

            ElseIf pVal.FormUID = "frmSRI" AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK AndAlso pVal.BeforeAction = True _
                AndAlso pVal.ItemUID = "MTX_SER" AndAlso pVal.ColUID = "COL1" Then
                'If Not pVal.Before_Action Then
                '    If pVal.ItemUID = "MTX_SER" And pVal.ColUID = "COL1" Then
                'rMatrix = oForm.Items.Item("MTX_SER").Specific

                Try
                    Dim sFolio As String = ""
                    Dim DocNumInicial As Integer = 0

                    'Dim oForm1 As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                    ' Dim oItemmatrix As Item = oForm.Items.Item("MTX_SER").Specific
                    Dim oMatrix As Matrix = oForm.Items.Item("MTX_SER").Specific

                    'SecInicio
                    Dim DocNumIni As Integer = 0
                    Try
                        DocNumIni = Convert.ToInt32((CType(oMatrix.Columns.Item("SecInicio").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText)).Value.ToString())
                    Catch ex As Exception
                    End Try


                    'UltimoSec
                    Dim NumeroInicialPrefijo As Integer = 0
                    Try
                        NumeroInicialPrefijo = Convert.ToInt32((CType(oMatrix.Columns.Item("UltimoSec").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText)).Value.ToString())
                    Catch ex As Exception
                    End Try



                    'TipoDoc
                    Dim TipoDoc As String = ""
                    Try
                        TipoDoc = (CType(oMatrix.Columns.Item("TipoD").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox)).Value.ToString()
                    Catch ex As Exception
                    End Try

                    'NombreSerie
                    Dim Prefij As String = ""
                    Try
                        Prefij = (CType(oMatrix.Columns.Item("SerN").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText)).Value.ToString()
                    Catch ex As Exception
                    End Try

                    'idSerie
                    Dim idSerie As Integer = 0
                    Try
                        idSerie = Convert.ToInt32((CType(oMatrix.Columns.Item("SerId").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText)).Value.ToString())
                    Catch ex As Exception
                    End Try

                    'obtener el Next
                    'Dim NexNumber As Integer = 0
                    ''NexNumber = Convert.ToInt32(oFuncionesAddon.getRSvalue("SELECT ""NextNumber"" FROM ""NNM1"" WHERE ""Series"" = " + idSerie.ToString() + "", "NextNumber", "0"))

                    ''--((A.DOCNUM - DocNumInicial)+ NumeroInicialPrefijo) + UltimoNumeroAntiguoProveedor
                    'Dim SgtPre As SAPbouiCOM.StaticText = oForm.Items.Item("SgtPre").Specific
                    'SgtPre.Item.ForeColor = RGB(0, 0, 0) 'RGB(6, 69, 173) 'RGB(0, 101, 184)
                    'If TipoDoc = "FC" Then
                    '    SgtPre.Caption = "El siguiente consecutivo es: " + Prefij + (((NexNumber - DocNumIni) + NumeroInicialPrefijo) + UltimoNumeroAntiguoProveedor).ToString()
                    'Else
                    '    SgtPre.Caption = "El siguiente consecutivo es: " + Prefij + NexNumber.ToString()
                    'End If




                    'oMatrix.SelectionMode = BoMatrixSelect.ms_Auto
                    'oMatrix.SelectRow(pVal.Row, True, False)
                Catch ex As Exception
                End Try



                '    End If
                'End If


            ElseIf pVal.FormUID = "frmSRI" AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK Then
                If Not pVal.Before_Action Then

                    If pVal.ItemUID = "MTX_SER" And pVal.ColUID = "SerId" Then '
                        rMatrix = oForm.Items.Item("MTX_SER").Specific

                        TipoSerie = rMatrix.Columns.Item("TipoD").Cells.Item(pVal.Row).Specific
                        IdSerie = rMatrix.Columns.Item("SerId").Cells.Item(pVal.Row).Specific
                        NombreSerie = rMatrix.Columns.Item("SerN").Cells.Item(pVal.Row).Specific

                        ofrmSRIConsulta.CargaFormularioConsulta(pVal.Row, TipoSerie, IdSerie, NombreSerie)

                    ElseIf pVal.ItemUID = "MTX_SER2" And pVal.ColUID = "SerId" Then ' PREFIJO


                        rMatrix = oForm.Items.Item("MTX_SER2").Specific

                        TipoSerie = rMatrix.Columns.Item("TipoD").Cells.Item(pVal.Row).Specific
                        IdSerie = rMatrix.Columns.Item("SerId").Cells.Item(pVal.Row).Specific
                        NombreSerie = rMatrix.Columns.Item("SerN").Cells.Item(pVal.Row).Specific

                        ofrmSRIConsulta.CargaFormularioConsulta(pVal.Row, TipoSerie, IdSerie, NombreSerie)


                        'rMatrix = oForm.Items.Item("MTX_SER").Specific

                        'IdSerie = rMatrix.Columns.Item("SerId").Cells.Item(pVal.Row).Specific
                        'TipoSerie = rMatrix.Columns.Item("TipoD").Cells.Item(pVal.Row).Specific

                        'Prefijo = rMatrix.Columns.Item("SerN").Cells.Item(pVal.Row).Specific
                        'RangoInicial = rMatrix.Columns.Item("SecIni").Cells.Item(pVal.Row).Specific
                        'RangoFinal = rMatrix.Columns.Item("SecFin").Cells.Item(pVal.Row).Specific
                        'DocNumInicial = rMatrix.Columns.Item("Sec").Cells.Item(pVal.Row).Specific
                        'Resolucion = rMatrix.Columns.Item("Resol").Cells.Item(pVal.Row).Specific
                        'ClaveTecnica = rMatrix.Columns.Item("ClaveT").Cells.Item(pVal.Row).Specific
                        'FechaInicial = rMatrix.Columns.Item("FiniD").Cells.Item(pVal.Row).Specific
                        'FechaFinal = rMatrix.Columns.Item("FfinD").Cells.Item(pVal.Row).Specific
                        'Contingencia = rMatrix.Columns.Item("Con").Cells.Item(pVal.Row).Specific

                        'ofrmSRIConsultaPre.CargaFormularioConsultaResoluciones(IdSerie, TipoSerie, Prefijo, RangoInicial, RangoFinal, DocNumInicial,
                        '                                                        Resolucion, ClaveTecnica, FechaInicial, FechaFinal, Contingencia)


                        '    ofrmConsulta.CargaFormularioConsulta("Lista de Acuerdos Globales", pVal.Row, "1250000025", "", "")
                        'ElseIf pVal.ItemUID = "gIngresos" And pVal.ColUID = "U_EXX_PT" Then ' Producto terminado
                        '    oForm = rsboApp.Forms.Item("frmHojaCosto")
                        '    Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("gIngresos").Specific
                        '    Dim oDT As SAPbouiCOM.DataTable = oGrid.DataTable
                        '    Dim Acuerdo As String = ""
                        '    Acuerdo = oDT.GetValue(2, pVal.Row)
                        '    Dim oSemana As String = oForm.Items.Item("txtSem").Specific.value.ToString()
                        '    If oSemana.Equals("") Then
                        '        rsboApp.StatusBar.SetText("Debe ingresar una semana calendario!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        BubbleEvent = False
                        '    Else
                        '        ofrmConsulta.CargaFormularioConsulta("Lista de Articulos", pVal.Row, "4", IIf(Acuerdo = "0", "", Acuerdo), oSemana)
                        '    End If

                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent

        Try
            Dim typeExx, idFormm As String
            typeExx = oFuncionesB1.FormularioActivo(idFormm)

            If typeExx = "frmSRI" Then
                If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                    Try
                        Dim mEdit As SAPbouiCOM.EditText = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("code").Specific
                        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("code").Enabled = False
                        mEdit.Value = CDbl(oFuncionesB1.getCorrelativo("code", "@SS_SER"))
                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
                        mMatrix.AddRow()
                        ''mMatrix.Columns.Item("COL1").Cells.Item(mMatrix.RowCount).Specific.String = mMatrix.RowCount
                        ''rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("COL_GRUART").Enabled = True
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If

                If pVal.MenuUID = "Agregar" And pVal.BeforeAction = False Then
                    rsboApp.Forms.ActiveForm.Freeze(True)
                    Try
                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
                        mMatrix.AddRow()
                        For i As Integer = 1 To mMatrix.RowCount
                            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
                        Next

                        Dim comb3 As SAPbouiCOM.ComboBox = mMatrix.Columns.Item("TipoD").Cells.Item(mMatrix.RowCount).Specific
                        comb3.Select("FC", BoSearchKey.psk_ByValue)


                        mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
                        mMatrix.Columns.Item("SerN").Cells.Item(mMatrix.RowCount).Specific.String = ""

                        If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                            rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                        End If
                    Catch ex As Exception
                    Finally
                        rsboApp.Forms.ActiveForm.Freeze(False)
                    End Try

                End If

                If pVal.MenuUID = "Agregar_FIS" And pVal.BeforeAction = False Then
                    rsboApp.Forms.ActiveForm.Freeze(True)
                    Try
                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER2").Specific
                        mMatrix.AddRow()
                        For i As Integer = 1 To mMatrix.RowCount
                            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
                        Next


                        mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
                        mMatrix.Columns.Item("SerN").Cells.Item(mMatrix.RowCount).Specific.String = ""

                        If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                            rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                        End If
                    Catch ex As Exception
                    Finally
                        rsboApp.Forms.ActiveForm.Freeze(False)
                    End Try

                End If

                If pVal.MenuUID = "Eliminar" And pVal.BeforeAction = False Then
                    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
                    mMatrix.DeleteRow(num)
                    For i As Integer = 1 To mMatrix.RowCount
                        mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
                    Next

                    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                    End If
                End If
                If pVal.MenuUID = "Eliminar_FIS" And pVal.BeforeAction = False Then
                    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER2").Specific
                    mMatrix.DeleteRow(num)
                    For i As Integer = 1 To mMatrix.RowCount
                        mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
                    Next

                    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                    End If
                End If
            End If

            'If pVal.MenuUID = "1282" And rsboApp.Forms.ActiveForm.UniqueID = "frmSRI" And pVal.BeforeAction = False Then
            '    Try
            '        Dim mEdit As SAPbouiCOM.EditText = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("code").Specific
            '        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("code").Enabled = False
            '        mEdit.Value = CDbl(oFuncionesB1.getCorrelativo("code", "@SS_SER"))
            '        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
            '        mMatrix.AddRow()
            '        ''mMatrix.Columns.Item("COL1").Cells.Item(mMatrix.RowCount).Specific.String = mMatrix.RowCount
            '        ''rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("COL_GRUART").Enabled = True
            '    Catch ex As Exception
            '        MsgBox(ex.Message)
            '    End Try
            'End If

            'If rsboApp.Forms.ActiveForm.UniqueID = "frmSRI" And pVal.MenuUID = "Agregar" And pVal.BeforeAction = False Then
            '    rsboApp.Forms.ActiveForm.Freeze(True)
            '    Try
            '        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
            '        mMatrix.AddRow()
            '        For i As Integer = 1 To mMatrix.RowCount
            '            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            '        Next

            '        Dim comb3 As SAPbouiCOM.ComboBox = mMatrix.Columns.Item("TipoD").Cells.Item(mMatrix.RowCount).Specific
            '        comb3.Select("FC", BoSearchKey.psk_ByValue)


            '        mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
            '        mMatrix.Columns.Item("SerN").Cells.Item(mMatrix.RowCount).Specific.String = ""

            '        If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            '            rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            '        End If
            '    Catch ex As Exception
            '    Finally
            '        rsboApp.Forms.ActiveForm.Freeze(False)
            '    End Try

            'End If
            'If rsboApp.Forms.ActiveForm.UniqueID = "frmSRI" And pVal.MenuUID = "Agregar_FIS" And pVal.BeforeAction = False Then
            '    rsboApp.Forms.ActiveForm.Freeze(True)
            '    Try
            '        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER2").Specific
            '        mMatrix.AddRow()
            '        For i As Integer = 1 To mMatrix.RowCount
            '            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            '        Next


            '        mMatrix.Columns.Item("SerId").Cells.Item(mMatrix.RowCount).Specific.String = ""
            '        mMatrix.Columns.Item("SerN").Cells.Item(mMatrix.RowCount).Specific.String = ""

            '        If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            '            rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            '        End If
            '    Catch ex As Exception
            '    Finally
            '        rsboApp.Forms.ActiveForm.Freeze(False)
            '    End Try

            'End If
            ''If rsboApp.Forms.ActiveForm.UniqueID = "frmSRI" And pVal.MenuUID = "Agregar_TRI" And pVal.BeforeAction = False Then
            ''    rsboApp.Forms.ActiveForm.Freeze(True)
            ''    Try
            ''        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_TRI").Specific
            ''        mMatrix.AddRow()
            ''        For i As Integer = 1 To mMatrix.RowCount
            ''            mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            ''        Next

            ''        mMatrix.Columns.Item("CODTRIB1").Cells.Item(mMatrix.RowCount).Specific.String = ""
            ''        mMatrix.Columns.Item("CODTRIB2").Cells.Item(mMatrix.RowCount).Specific.String = ""

            ''        If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            ''            rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            ''        End If
            ''    Catch ex As Exception
            ''    Finally
            ''        rsboApp.Forms.ActiveForm.Freeze(False)
            ''    End Try

            ''End If

            'If rsboApp.Forms.ActiveForm.UniqueID = "frmSRI" And pVal.MenuUID = "Eliminar" And pVal.BeforeAction = False Then
            '    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER").Specific
            '    mMatrix.DeleteRow(num)
            '    For i As Integer = 1 To mMatrix.RowCount
            '        mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            '    Next

            '    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            '        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            '    End If
            'End If
            'If rsboApp.Forms.ActiveForm.UniqueID = "frmSRI" And pVal.MenuUID = "Eliminar_FIS" And pVal.BeforeAction = False Then
            '    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_SER2").Specific
            '    mMatrix.DeleteRow(num)
            '    For i As Integer = 1 To mMatrix.RowCount
            '        mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            '    Next

            '    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            '        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            '    End If
            'End If
            ''If rsboApp.Forms.ActiveForm.UniqueID = "frmSRI" And pVal.MenuUID = "Eliminar_TRI" And pVal.BeforeAction = False Then
            ''    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_TRI").Specific
            ''    mMatrix.DeleteRow(num)
            ''    For i As Integer = 1 To mMatrix.RowCount
            ''        mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            ''    Next

            ''    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            ''        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            ''    End If
            ''End If

            ''' NO PERMITIR ELIMINAR
            ''If rsboApp.Forms.ActiveForm.UniqueID = "frmSRI" And pVal.MenuUID = "1283" And pVal.BeforeAction = False Then
            ''    Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_TRI").Specific
            ''    mMatrix.DeleteRow(num)
            ''    For i As Integer = 1 To mMatrix.RowCount
            ''        mMatrix.Columns.Item("COL1").Cells.Item(i).Specific.String = i
            ''    Next

            ''    If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
            ''        rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
            ''    End If
            ''End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log($"Error rsboApp_MenuEvent: {ex.Message}", "frmSRI")

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
                    ' oForm.Close()
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub rsboApp_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent
        Try


            If eventInfo.FormUID = "frmSRI" Then

                If eventInfo.ItemUID = "MTX_SER" Then

                    If eventInfo.ColUID = "COL1" Then
                        Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
                        Dim oMenus As SAPbouiCOM.Menus = Nothing

                        If eventInfo.BeforeAction = True Then

                            Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(eventInfo.FormUID)

                            If mForm.Mode = BoFormMode.fm_ADD_MODE Or mForm.Mode = BoFormMode.fm_OK_MODE Then

                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
                                Try
                                    num = eventInfo.Row

                                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                    oMenuItem = rsboApp.Menus.Item("1280")
                                    If oMenuItem.SubMenus.Exists("Agregar") Then
                                        rsboApp.Menus.RemoveEx("Agregar")

                                    End If
                                    If oMenuItem.SubMenus.Exists("Eliminar") Then
                                        rsboApp.Menus.RemoveEx("Eliminar")
                                    End If
                                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                    oCreationPackage.UniqueID = "Agregar"
                                    oCreationPackage.String = "Agregar fila"
                                    oCreationPackage.Enabled = True
                                    oCreationPackage.Position = 20
                                    oMenuItem = rsboApp.Menus.Item("1280")
                                    oMenus = oMenuItem.SubMenus
                                    oMenus.AddEx(oCreationPackage)

                                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                    oCreationPackage.UniqueID = "Eliminar"
                                    oCreationPackage.String = "Eliminar fila"
                                    oCreationPackage.Enabled = True
                                    oCreationPackage.Position = 21
                                    oMenuItem = rsboApp.Menus.Item("1280")
                                    oMenus = oMenuItem.SubMenus
                                    oMenus.AddEx(oCreationPackage)

                                    If oMenuItem.SubMenus.Exists("1283") Then
                                        rsboApp.Menus.RemoveEx("1283")
                                    End If

                                Catch ex As Exception
                                    'MessageBox.Show(ex.Message)
                                End Try
                            End If
                        Else
                            Try
                                oMenuItem = rsboApp.Menus.Item("1280")
                                If oMenuItem.SubMenus.Exists("Agregar") Then
                                    rsboApp.Menus.RemoveEx("Agregar")

                                End If
                                If oMenuItem.SubMenus.Exists("Eliminar") Then
                                    rsboApp.Menus.RemoveEx("Eliminar")
                                End If
                            Catch ex As Exception
                                rsboApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End Try
                        End If

                    End If
                ElseIf eventInfo.ItemUID = "MTX_SER2" Then
                    If eventInfo.ColUID = "COL1" Then
                        Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
                        Dim oMenus As SAPbouiCOM.Menus = Nothing

                        If eventInfo.BeforeAction = True Then

                            Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(eventInfo.FormUID)

                            If mForm.Mode = BoFormMode.fm_ADD_MODE Or mForm.Mode = BoFormMode.fm_OK_MODE Then

                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
                                Try
                                    num = eventInfo.Row

                                    oCreationPackage = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                    oMenuItem = rsboApp.Menus.Item("1280")
                                    If oMenuItem.SubMenus.Exists("Agregar_FIS") Then
                                        rsboApp.Menus.RemoveEx("Agregar_FIS")

                                    End If
                                    If oMenuItem.SubMenus.Exists("Eliminar_FIS") Then
                                        rsboApp.Menus.RemoveEx("Eliminar_FIS")
                                    End If
                                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                    oCreationPackage.UniqueID = "Agregar_FIS"
                                    oCreationPackage.String = "Agregar fila"
                                    oCreationPackage.Enabled = True
                                    oCreationPackage.Position = 20
                                    oMenuItem = rsboApp.Menus.Item("1280")
                                    oMenus = oMenuItem.SubMenus
                                    oMenus.AddEx(oCreationPackage)

                                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                    oCreationPackage.UniqueID = "Eliminar_FIS"
                                    oCreationPackage.String = "Eliminar fila"
                                    oCreationPackage.Enabled = True
                                    oCreationPackage.Position = 21
                                    oMenuItem = rsboApp.Menus.Item("1280")
                                    oMenus = oMenuItem.SubMenus
                                    oMenus.AddEx(oCreationPackage)

                                Catch ex As Exception
                                    'MessageBox.Show(ex.Message)
                                End Try
                            End If
                        Else
                            Try
                                oMenuItem = rsboApp.Menus.Item("1280")
                                If oMenuItem.SubMenus.Exists("Agregar_FIS") Then
                                    rsboApp.Menus.RemoveEx("Agregar_FIS")

                                End If
                                If oMenuItem.SubMenus.Exists("Eliminar_FIS") Then
                                    rsboApp.Menus.RemoveEx("Eliminar_FIS")
                                End If
                            Catch ex As Exception
                                rsboApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End Try
                        End If

                    End If

                End If

            End If
        Catch ex As Exception

        End Try
    End Sub

    'Private Sub rSboApp_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.FormDataEvent
    '    If BusinessObjectInfo.FormTypeEx = "frmSRI" Then
    '        Select Case BusinessObjectInfo.EventType
    '            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
    '                If BusinessObjectInfo.BeforeAction Then
    '                    Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(BusinessObjectInfo.FormUID)
    '                    Dim mEdit As SAPbouiCOM.EditText = mForm.Items.Item("code").Specific
    '                    If Integer.Parse(mEdit.Value) > 1 Then
    '                        'rsboApp.SetStatusBarMessage(NombreAddon + " - Ya existe una configuración, por favor consutarla y actualizarla de sel el caso.", BoMessageTime.bmt_Medium, True)
    '                        rsboApp.MessageBox(NombreAddon + " - Ya existe una configuración, por favor consutarla y actualizarla de sel el caso.")
    '                        BubbleEvent = False
    '                        Exit Sub
    '                    End If

    '                End If
    '        End Select
    '    End If

    'End Sub


End Class
