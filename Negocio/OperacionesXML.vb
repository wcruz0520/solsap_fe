Imports System.Xml


Namespace ssXML

    Public Class OperacionesXML


        Public Shared Function LeerXMLFactura2(ruta As String) As Entidades.wsEDoc_ConsultaRecepcion.ENTFactura

            Utilitario.Util_Log.Escribir_Log("Consultando XML Facturas a procesar..", "OperacionesXML")
            Try
                'Dim razonSocial As String = ""
                'Dim ruc As String = ""
                'Dim estab As String = ""
                'Dim ptoEmi As String = ""
                'Dim secuencial As String = ""

                Dim m_xmld As XmlDocument
                Dim m_nodelist As XmlNodeList

                Dim Factura As New Entidades.wsEDoc_ConsultaRecepcion.ENTFactura



                m_xmld = New XmlDocument()

                m_xmld.Load(ruta)

                Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")
                If (Not (nodoAut) Is Nothing) Then

                    For Each facAut As XmlNode In nodoAut.ChildNodes
                        Select Case facAut.Name
                            Case "numeroAutorizacion"
                                ' Factura.FacturaCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                                Factura.AutorizacionSRI = facAut.InnerText.ToString
                            Case "fechaAutorizacion"
                                ' Factura.FacturaCabecera._FechaAutorizacion = CDate(facAut.InnerText)
                                Factura.FechaAutorizacion = CDate(facAut.InnerText)
                        End Select
                    Next

                    Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

                    Dim comprobante = nodo.InnerText

                    Dim nodofactura As New XmlDocument()
                    nodofactura.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable


                    For Each fac As XmlNode In nodofactura.ChildNodes
                        Dim ad = fac.Name
                        Select Case ad
                            Case "factura"
                                For Each nodoInfTri As XmlNode In fac.ChildNodes
                                    Dim p = nodoInfTri.Name
                                    Select Case p
                                        Case "infoTributaria"
                                            For Each n As XmlNode In nodoInfTri.ChildNodes
                                                Select Case n.Name
                                                    Case "razonSocial"
                                                        Dim razonSocial = n.InnerText.ToString
                                                        ' Factura.FacturaCabecera._RazonSocial = razonSocial
                                                        Factura.RazonSocial = razonSocial
                                                    Case "ruc"
                                                        Dim ruc = n.InnerText.ToString
                                                        'Factura.FacturaCabecera._ruc = ruc
                                                        Factura.Ruc = ruc
                                                    Case "estab"
                                                        Dim estab = n.InnerText.ToString
                                                        ' Factura.FacturaCabecera._estab = estab
                                                        Factura.Establecimiento = estab
                                                    Case "ptoEmi"
                                                        Dim ptoEmi = n.InnerText.ToString
                                                        'Factura.FacturaCabecera._ptoEmi = ptoEmi
                                                        Factura.PuntoEmision = ptoEmi
                                                    Case "secuencial"
                                                        Dim secuencial = n.InnerText.ToString
                                                        '  Factura.FacturaCabecera._secuencial = secuencial
                                                        Factura.Secuencial = secuencial

                                                    Case "claveAcceso"
                                                        '   Factura.FacturaCabecera._claveAcceso = n.InnerText.ToString
                                                        Factura.ClaveAcceso = n.InnerText.ToString
                                                End Select

                                            Next
                                        Case "infoFactura"
                                            For Each infoFac As XmlNode In nodoInfTri.ChildNodes
                                                Select Case infoFac.Name
                                                    Case "fechaEmision"
                                                        Dim fechaemision = infoFac.InnerText
                                                        'Factura.FacturaCabecera._fechaEmision = fechaemision
                                                        Factura.FechaEmision = fechaemision
                                                    Case "contribuyenteEspecial"
                                                        Dim contribuyenteEspecial = infoFac.InnerText
                                                        '  Factura.FacturaCabecera._contribuyenteEspecial = contribuyenteEspecial
                                                        Factura.ContribuyenteEspecial = contribuyenteEspecial
                                                    Case "dirEstablecimiento"
                                                        Dim dirEstablecimiento = infoFac.InnerText
                                                        ' Factura.FacturaCabecera._dirEstablecimiento = dirEstablecimiento
                                                        Factura.DireccionEstablecimiento = dirEstablecimiento
                                                    Case "razonSocialComprador"
                                                        Dim razonSocialComprador = infoFac.InnerText
                                                        'Factura.FacturaCabecera._razonSocialComprador = razonSocialComprador
                                                        Factura.RazonSocialComprador = razonSocialComprador
                                                    Case "identificacionComprador"
                                                        Dim identificacionComprador = infoFac.InnerText
                                                        'Factura.FacturaCabecera._identificacionComprador = identificacionComprador
                                                        Factura.IdentificacionComprador = identificacionComprador
                                                    Case "direccionComprador"
                                                        Dim direccionComprador = infoFac.InnerText
                                                        'Factura.FacturaCabecera._direccionComprador = direccionComprador

                                                    Case "totalSinImpuestos"
                                                        Dim totalSinImpuestos = infoFac.InnerText
                                                        'Factura.FacturaCabecera._totalSinImpuestos = totalSinImpuestos
                                                        Factura.TotalSinImpuesto = totalSinImpuestos
                                                    Case "totalDescuento"
                                                        Dim totalDescuento = infoFac.InnerText
                                                        'Factura.FacturaCabecera._totalDescuento = totalDescuento 'CDec(totalDescuento)
                                                        Factura.TotalDescuento = totalDescuento
                                                    Case "totalConImpuestos"

                                                        Dim XDocumentoImpuestoCabecera = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFacturaImpuesto)

                                                        For Each totalConImpuestos As XmlNode In infoFac.ChildNodes
                                                            Select Case totalConImpuestos.Name
                                                                Case "totalImpuesto"
                                                                    'Dim facCabImp As New FacturaCabeceraImpuestos
                                                                    Dim impuestoCab As New Entidades.wsEDoc_ConsultaRecepcion.ENTFacturaImpuesto
                                                                    For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                                        Select Case totalImpuesto.Name
                                                                            Case "codigo"
                                                                                Dim codigo = totalImpuesto.InnerText
                                                                                'facCabImp._codigo = CInt(codigo)
                                                                                impuestoCab.Codigo = CInt(codigo)
                                                                            Case "codigoPorcentaje"
                                                                                Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                                'facCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                impuestoCab.CodigoPorcentaje = CInt(codigoPorcentaje)
                                                                            Case "baseImponible"
                                                                                Dim baseImponible = totalImpuesto.InnerText
                                                                                'facCabImp._baseImponible = baseImponible 'CDec(baseImponible)
                                                                                impuestoCab.BaseImponible = baseImponible
                                                                            Case "tarifa"
                                                                                Dim tarifa = totalImpuesto.InnerText
                                                                                ' facCabImp._tarifa = tarifa 'CDec(tarifa)
                                                                                impuestoCab.Tarifa = tarifa
                                                                            Case "valor"
                                                                                Dim valor = totalImpuesto.InnerText
                                                                                'facCabImp._valor = valor 'CDec(valor)
                                                                                impuestoCab.Valor = valor
                                                                        End Select

                                                                    Next
                                                                    'Factura.FacturaCabecera._impuestos.Add(facCabImp)
                                                                    XDocumentoImpuestoCabecera.Add(impuestoCab)

                                                            End Select
                                                        Next

                                                        If XDocumentoImpuestoCabecera.Count > 0 Then

                                                            Factura.ENTFacturaImpuesto = XDocumentoImpuestoCabecera.ToArray

                                                        End If

                                                    Case "importeTotal"
                                                        Dim importeTotal = infoFac.InnerText
                                                        'Factura.FacturaCabecera._importeTotal = importeTotal
                                                        Factura.ImporteTotal = importeTotal
                                                    Case "moneda"
                                                        Dim moneda = infoFac.InnerText

                                                    Case "pagos"
                                                        Dim xPagodocumento As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTPagos)
                                                        For Each pagos As XmlNode In infoFac.ChildNodes
                                                            Select Case pagos.Name
                                                                Case "pago"

                                                                    Dim pagodocumento = New Entidades.wsEDoc_ConsultaRecepcion.ENTPagos

                                                                    For Each pago As XmlNode In pagos.ChildNodes
                                                                        Select Case pago.Name
                                                                            Case "formaPago"
                                                                                Dim formaPago = pago.InnerText
                                                                                ' Factura.FacturaCabecera._formaPago = formaPago
                                                                                pagodocumento.FormaPago = formaPago
                                                                            Case "total"
                                                                                Dim total = pago.InnerText
                                                                                'Factura.FacturaCabecera._totalFormaPago = total
                                                                                pagodocumento.Total = total
                                                                            Case "plazo"
                                                                                Dim plazo = pago.InnerText
                                                                                'Factura.FacturaCabecera._plazo = CInt(plazo)
                                                                                pagodocumento.Plazo = plazo
                                                                            Case "unidadTiempo"
                                                                                Dim unidadTiempo = pago.InnerText
                                                                                'Factura.FacturaCabecera._unidadTiempo = unidadTiempo
                                                                                pagodocumento.UnidadTiempo = unidadTiempo
                                                                        End Select
                                                                    Next

                                                                    xPagodocumento.Add(pagodocumento)

                                                            End Select
                                                        Next

                                                        'si existen los agrego
                                                        If xPagodocumento.Count > 0 Then
                                                            Factura.ENTPagos = xPagodocumento.ToArray
                                                        End If

                                                End Select
                                            Next
                                        Case "detalles"

                                            Dim XDetalles = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura)

                                            For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                                Select Case detalles.Name
                                                    Case "detalle"
                                                        Dim FacDet As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura

                                                        For Each detalle As XmlNode In detalles.ChildNodes
                                                            Select Case detalle.Name
                                                                Case "codigoPrincipal"
                                                                    Dim codigoPrincipal = detalle.InnerText
                                                                    ' FacDet._codigoPrincipal = codigoPrincipal
                                                                    FacDet.CodigoPrincipal = codigoPrincipal
                                                                Case "codigoAuxiliar"
                                                                    Dim codigoAuxiliar = detalle.InnerText
                                                                    'FacDet._codigoAuxiliar = codigoAuxiliar
                                                                    FacDet.CodigoAuxiliar = codigoAuxiliar
                                                                Case "descripcion"
                                                                    Dim descripcion = detalle.InnerText
                                                                    'FacDet._descripcion = descripcion
                                                                    FacDet.Descripcion = descripcion
                                                                Case "cantidad"
                                                                    Dim cantidad = detalle.InnerText
                                                                    'FacDet._cantidad = cantidad 'String.Format(CultureInfo.InvariantCulture, "{0:N0}", cantidad) 'CDec(cantidad)
                                                                    FacDet.Cantidad = cantidad
                                                                Case "precioUnitario"
                                                                    Dim precioUnitario = detalle.InnerText
                                                                    ' FacDet._precioUnitario = precioUnitario 'CDec(precioUnitario)
                                                                    FacDet.PrecioUnitario = precioUnitario
                                                                Case "descuento"
                                                                    Dim descuento = detalle.InnerText
                                                                    'FacDet._descuento = descuento 'CDec(descuento)
                                                                    FacDet.Descuento = descuento
                                                                Case "precioTotalSinImpuesto"
                                                                    Dim precioTotalSinImpuesto = detalle.InnerText
                                                                    ' FacDet._precioTotalSinImpuesto = precioTotalSinImpuesto
                                                                    FacDet.PrecioTotalSinImpuesto = precioTotalSinImpuesto
                                                                Case "impuestos"
                                                                    Dim ListaImpuestoDetalle = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFacturaImpuesto)
                                                                    For Each impuestos As XmlNode In detalle.ChildNodes
                                                                        Select Case impuestos.Name
                                                                            Case "impuesto"
                                                                                Dim FacDetImp As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFacturaImpuesto
                                                                                For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                                    Select Case impuesto.Name
                                                                                        Case "codigo"
                                                                                            Dim codigo = impuesto.InnerText
                                                                                            ' FacDetImp._codigo = CInt(codigo)
                                                                                            FacDetImp.Codigo = codigo
                                                                                        Case "codigoPorcentaje"
                                                                                            Dim codigoPorcentaje = impuesto.InnerText
                                                                                            'FacDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                            FacDetImp.CodigoPorcentaje = codigoPorcentaje
                                                                                        Case "tarifa"
                                                                                            Dim tarifa = impuesto.InnerText
                                                                                            'FacDetImp._tarifa = tarifa 'CDec(tarifa)
                                                                                            FacDetImp.Tarifa = tarifa
                                                                                        Case "baseImponible"
                                                                                            Dim baseImponible = impuesto.InnerText
                                                                                            'FacDetImp._baseImponible = baseImponible 'CDec(baseImponible)
                                                                                            FacDetImp.BaseImponible = baseImponible
                                                                                        Case "valor"
                                                                                            Dim valor = impuesto.InnerText
                                                                                            'FacDetImp._valor = valor 'CDec(valor)
                                                                                            FacDetImp.Valor = valor
                                                                                    End Select
                                                                                Next
                                                                                ListaImpuestoDetalle.Add(FacDetImp)
                                                                        End Select
                                                                    Next

                                                                    If ListaImpuestoDetalle.Count > 0 Then

                                                                        FacDet.ENTDetalleFacturaImpuesto = ListaImpuestoDetalle.ToArray

                                                                    End If

                                                            End Select
                                                        Next

                                                        XDetalles.Add(FacDet)

                                                End Select
                                            Next

                                            If XDetalles.Count > 0 Then


                                                Factura.ENTDetalleFactura = XDetalles.ToArray

                                            End If


                                    End Select
                                Next
                        End Select
                    Next

                Else

                    Dim nodoFac = m_xmld.SelectSingleNode("factura")

                    For Each nodoInfTri As XmlNode In nodoFac.ChildNodes
                        Dim p = nodoInfTri.Name
                        Select Case p
                            Case "infoTributaria"
                                For Each n As XmlNode In nodoInfTri.ChildNodes
                                    Select Case n.Name
                                        Case "razonSocial"
                                            Dim razonSocial = n.InnerText.ToString
                                            ' Factura.FacturaCabecera._RazonSocial = razonSocial
                                            Factura.RazonSocial = razonSocial
                                        Case "ruc"
                                            Dim ruc = n.InnerText.ToString
                                            'Factura.FacturaCabecera._ruc = ruc
                                            Factura.Ruc = ruc
                                        Case "estab"
                                            Dim estab = n.InnerText.ToString
                                            ' Factura.FacturaCabecera._estab = estab
                                            Factura.Establecimiento = estab
                                        Case "ptoEmi"
                                            Dim ptoEmi = n.InnerText.ToString
                                            'Factura.FacturaCabecera._ptoEmi = ptoEmi
                                            Factura.PuntoEmision = ptoEmi
                                        Case "secuencial"
                                            Dim secuencial = n.InnerText.ToString
                                            '  Factura.FacturaCabecera._secuencial = secuencial
                                            Factura.Secuencial = secuencial

                                        Case "claveAcceso"
                                            '   Factura.FacturaCabecera._claveAcceso = n.InnerText.ToString
                                            Factura.ClaveAcceso = n.InnerText.ToString
                                    End Select

                                Next
                            Case "infoFactura"
                                For Each infoFac As XmlNode In nodoInfTri.ChildNodes
                                    Select Case infoFac.Name
                                        Case "fechaEmision"
                                            Dim fechaemision = infoFac.InnerText
                                            'Factura.FacturaCabecera._fechaEmision = fechaemision
                                            Factura.FechaEmision = fechaemision
                                        Case "contribuyenteEspecial"
                                            Dim contribuyenteEspecial = infoFac.InnerText
                                            '  Factura.FacturaCabecera._contribuyenteEspecial = contribuyenteEspecial
                                            Factura.ContribuyenteEspecial = contribuyenteEspecial
                                        Case "dirEstablecimiento"
                                            Dim dirEstablecimiento = infoFac.InnerText
                                            ' Factura.FacturaCabecera._dirEstablecimiento = dirEstablecimiento
                                            Factura.DireccionEstablecimiento = dirEstablecimiento
                                        Case "razonSocialComprador"
                                            Dim razonSocialComprador = infoFac.InnerText
                                            'Factura.FacturaCabecera._razonSocialComprador = razonSocialComprador
                                            Factura.RazonSocialComprador = razonSocialComprador
                                        Case "identificacionComprador"
                                            Dim identificacionComprador = infoFac.InnerText
                                            'Factura.FacturaCabecera._identificacionComprador = identificacionComprador
                                            Factura.IdentificacionComprador = identificacionComprador
                                        Case "direccionComprador"
                                            Dim direccionComprador = infoFac.InnerText
                                                        'Factura.FacturaCabecera._direccionComprador = direccionComprador

                                        Case "totalSinImpuestos"
                                            Dim totalSinImpuestos = infoFac.InnerText
                                            'Factura.FacturaCabecera._totalSinImpuestos = totalSinImpuestos
                                            Factura.TotalSinImpuesto = totalSinImpuestos
                                        Case "totalDescuento"
                                            Dim totalDescuento = infoFac.InnerText
                                            'Factura.FacturaCabecera._totalDescuento = totalDescuento 'CDec(totalDescuento)
                                            Factura.TotalDescuento = totalDescuento
                                        Case "totalConImpuestos"

                                            Dim XDocumentoImpuestoCabecera = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTFacturaImpuesto)

                                            For Each totalConImpuestos As XmlNode In infoFac.ChildNodes
                                                Select Case totalConImpuestos.Name
                                                    Case "totalImpuesto"
                                                        'Dim facCabImp As New FacturaCabeceraImpuestos
                                                        Dim impuestoCab As New Entidades.wsEDoc_ConsultaRecepcion.ENTFacturaImpuesto
                                                        For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                            Select Case totalImpuesto.Name
                                                                Case "codigo"
                                                                    Dim codigo = totalImpuesto.InnerText
                                                                    'facCabImp._codigo = CInt(codigo)
                                                                    impuestoCab.Codigo = CInt(codigo)
                                                                Case "codigoPorcentaje"
                                                                    Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                    'facCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                    impuestoCab.CodigoPorcentaje = CInt(codigoPorcentaje)
                                                                Case "baseImponible"
                                                                    Dim baseImponible = totalImpuesto.InnerText
                                                                    'facCabImp._baseImponible = baseImponible 'CDec(baseImponible)
                                                                    impuestoCab.BaseImponible = baseImponible
                                                                Case "tarifa"
                                                                    Dim tarifa = totalImpuesto.InnerText
                                                                    ' facCabImp._tarifa = tarifa 'CDec(tarifa)
                                                                    impuestoCab.Tarifa = tarifa
                                                                Case "valor"
                                                                    Dim valor = totalImpuesto.InnerText
                                                                    'facCabImp._valor = valor 'CDec(valor)
                                                                    impuestoCab.Valor = valor
                                                            End Select

                                                        Next
                                                        'Factura.FacturaCabecera._impuestos.Add(facCabImp)
                                                        XDocumentoImpuestoCabecera.Add(impuestoCab)

                                                End Select
                                            Next

                                            If XDocumentoImpuestoCabecera.Count > 0 Then

                                                Factura.ENTFacturaImpuesto = XDocumentoImpuestoCabecera.ToArray

                                            End If

                                        Case "importeTotal"
                                            Dim importeTotal = infoFac.InnerText
                                            'Factura.FacturaCabecera._importeTotal = importeTotal
                                            Factura.ImporteTotal = importeTotal
                                        Case "moneda"
                                            Dim moneda = infoFac.InnerText

                                        Case "pagos"
                                            Dim xPagodocumento As New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTPagos)
                                            For Each pagos As XmlNode In infoFac.ChildNodes
                                                Select Case pagos.Name
                                                    Case "pago"

                                                        Dim pagodocumento = New Entidades.wsEDoc_ConsultaRecepcion.ENTPagos

                                                        For Each pago As XmlNode In pagos.ChildNodes
                                                            Select Case pago.Name
                                                                Case "formaPago"
                                                                    Dim formaPago = pago.InnerText
                                                                    ' Factura.FacturaCabecera._formaPago = formaPago
                                                                    pagodocumento.FormaPago = formaPago
                                                                Case "total"
                                                                    Dim total = pago.InnerText
                                                                    'Factura.FacturaCabecera._totalFormaPago = total
                                                                    pagodocumento.Total = total
                                                                Case "plazo"
                                                                    Dim plazo = pago.InnerText
                                                                    'Factura.FacturaCabecera._plazo = CInt(plazo)
                                                                    pagodocumento.Plazo = plazo
                                                                Case "unidadTiempo"
                                                                    Dim unidadTiempo = pago.InnerText
                                                                    'Factura.FacturaCabecera._unidadTiempo = unidadTiempo
                                                                    pagodocumento.UnidadTiempo = unidadTiempo
                                                            End Select
                                                        Next

                                                        xPagodocumento.Add(pagodocumento)

                                                End Select
                                            Next

                                            'si existen los agrego
                                            If xPagodocumento.Count > 0 Then
                                                Factura.ENTPagos = xPagodocumento.ToArray
                                            End If

                                    End Select
                                Next
                            Case "detalles"

                                Dim XDetalles = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura)

                                For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                    Select Case detalles.Name
                                        Case "detalle"
                                            Dim FacDet As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFactura

                                            For Each detalle As XmlNode In detalles.ChildNodes
                                                Select Case detalle.Name
                                                    Case "codigoPrincipal"
                                                        Dim codigoPrincipal = detalle.InnerText
                                                        ' FacDet._codigoPrincipal = codigoPrincipal
                                                        FacDet.CodigoPrincipal = codigoPrincipal
                                                    Case "codigoAuxiliar"
                                                        Dim codigoAuxiliar = detalle.InnerText
                                                        'FacDet._codigoAuxiliar = codigoAuxiliar
                                                        FacDet.CodigoAuxiliar = codigoAuxiliar
                                                    Case "descripcion"
                                                        Dim descripcion = detalle.InnerText
                                                        'FacDet._descripcion = descripcion
                                                        FacDet.Descripcion = descripcion
                                                    Case "cantidad"
                                                        Dim cantidad = detalle.InnerText
                                                        'FacDet._cantidad = cantidad 'String.Format(CultureInfo.InvariantCulture, "{0:N0}", cantidad) 'CDec(cantidad)
                                                        FacDet.Cantidad = cantidad
                                                    Case "precioUnitario"
                                                        Dim precioUnitario = detalle.InnerText
                                                        ' FacDet._precioUnitario = precioUnitario 'CDec(precioUnitario)
                                                        FacDet.PrecioUnitario = precioUnitario
                                                    Case "descuento"
                                                        Dim descuento = detalle.InnerText
                                                        'FacDet._descuento = descuento 'CDec(descuento)
                                                        FacDet.Descuento = descuento
                                                    Case "precioTotalSinImpuesto"
                                                        Dim precioTotalSinImpuesto = detalle.InnerText
                                                        ' FacDet._precioTotalSinImpuesto = precioTotalSinImpuesto
                                                        FacDet.PrecioTotalSinImpuesto = precioTotalSinImpuesto
                                                    Case "impuestos"
                                                        Dim ListaImpuestoDetalle = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFacturaImpuesto)
                                                        For Each impuestos As XmlNode In detalle.ChildNodes
                                                            Select Case impuestos.Name
                                                                Case "impuesto"
                                                                    Dim FacDetImp As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleFacturaImpuesto
                                                                    For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                        Select Case impuesto.Name
                                                                            Case "codigo"
                                                                                Dim codigo = impuesto.InnerText
                                                                                ' FacDetImp._codigo = CInt(codigo)
                                                                                FacDetImp.Codigo = codigo
                                                                            Case "codigoPorcentaje"
                                                                                Dim codigoPorcentaje = impuesto.InnerText
                                                                                'FacDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                FacDetImp.CodigoPorcentaje = codigoPorcentaje
                                                                            Case "tarifa"
                                                                                Dim tarifa = impuesto.InnerText
                                                                                'FacDetImp._tarifa = tarifa 'CDec(tarifa)
                                                                                FacDetImp.Tarifa = tarifa
                                                                            Case "baseImponible"
                                                                                Dim baseImponible = impuesto.InnerText
                                                                                'FacDetImp._baseImponible = baseImponible 'CDec(baseImponible)
                                                                                FacDetImp.BaseImponible = baseImponible
                                                                            Case "valor"
                                                                                Dim valor = impuesto.InnerText
                                                                                'FacDetImp._valor = valor 'CDec(valor)
                                                                                FacDetImp.Valor = valor
                                                                        End Select
                                                                    Next
                                                                    ListaImpuestoDetalle.Add(FacDetImp)
                                                            End Select
                                                        Next

                                                        If ListaImpuestoDetalle.Count > 0 Then

                                                            FacDet.ENTDetalleFacturaImpuesto = ListaImpuestoDetalle.ToArray

                                                        End If

                                                End Select
                                            Next

                                            XDetalles.Add(FacDet)

                                    End Select
                                Next

                                If XDetalles.Count > 0 Then


                                    Factura.ENTDetalleFactura = XDetalles.ToArray

                                End If


                        End Select
                    Next

                End If


                'GuardaLog("FC clave de acceso: " + Factura.FacturaCabecera._claveAcceso + " leido Correctamente!!" + " - nombre del archivo: " + Path.GetFileName(ruta))

                Utilitario.Util_Log.Escribir_Log("FC clave de acceso:  " + Factura.ClaveAcceso + " leido Correctamente!!" + " - nombre del archivo: , ", " OperacionesXML")

                Return Factura
            Catch ex As Exception
                ' GuardaLog("Error al leer XML, EX  ==> " + ex.Message + " con nombre: " + Path.GetFileName(ruta))
                Utilitario.Util_Log.Escribir_Log("Error al leer XML, EX  ==> " + ex.Message + " con nombre: ", " OperacionesXML")

                Return Nothing
            End Try

        End Function

        Public Shared Function LeerXMLNotaCredito2(ruta As String) As Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito

            'GuardaLog("Consultando XML Notas de Credito a procesar..")

            Utilitario.Util_Log.Escribir_Log("Consultando XML Notas de Credito a procesar..", " OperacionesXML")


            Try

                Dim _RUTA As String = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\ncprueba.xml"

                Dim mensaje As String = ""
                Dim m_xmld As XmlDocument
                m_xmld = New XmlDocument()

                m_xmld.Load(ruta)


                Dim Ncredito As New Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCredito


                Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")

                If (Not (nodoAut) Is Nothing) Then

                    For Each NcAut As XmlNode In nodoAut.ChildNodes
                        Select Case NcAut.Name
                            Case "numeroAutorizacion"
                                Ncredito.AutorizacionSRI = NcAut.InnerText.ToString
                            Case "fechaAutorizacion"
                                Ncredito.FechaAutorizacion = CDate(NcAut.InnerText)
                        End Select
                    Next

                    Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

                    Dim comprobante = nodo.InnerText

                    Dim nodoNc As New XmlDocument()
                    nodoNc.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable

                    Dim razonSocial As String = ""
                    Dim ruc As String = ""
                    Dim estab As String = ""
                    Dim ptoEmi As String = ""
                    Dim secuencial As String = ""


                    For Each notac As XmlNode In nodoNc.ChildNodes
                        Dim ad = notac.Name
                        Select Case ad
                            Case "notaCredito"
                                For Each nodoInfTri As XmlNode In notac.ChildNodes
                                    Dim p = nodoInfTri.Name
                                    Select Case p
                                        Case "infoTributaria"
                                            For Each n As XmlNode In nodoInfTri.ChildNodes
                                                Select Case n.Name
                                                    Case "razonSocial"
                                                        razonSocial = n.InnerText.ToString
                                                        Ncredito.RazonSocial = razonSocial
                                                    Case "ruc"
                                                        ruc = n.InnerText.ToString
                                                        Ncredito.Ruc = ruc
                                                    Case "estab"
                                                        estab = n.InnerText.ToString
                                                        Ncredito.Establecimiento = estab
                                                    Case "ptoEmi"
                                                        ptoEmi = n.InnerText.ToString
                                                        Ncredito.PuntoEmision = ptoEmi
                                                    Case "secuencial"
                                                        secuencial = n.InnerText.ToString
                                                        Ncredito.Secuencial = secuencial
                                                    Case "claveAcceso"
                                                        Dim claveAcceso = n.InnerText.ToString
                                                        Ncredito.ClaveAcceso = claveAcceso
                                                End Select

                                            Next
                                        Case "infoNotaCredito"
                                            For Each infoNc As XmlNode In nodoInfTri.ChildNodes
                                                Select Case infoNc.Name
                                                    Case "fechaEmision"
                                                        Dim fechaemision = infoNc.InnerText
                                                        Ncredito.FechaEmision = fechaemision
                                                    Case "dirEstablecimiento"
                                                        Dim dirEstablecimiento = infoNc.InnerText
                                                        Ncredito.DireccionEstablecimiento = dirEstablecimiento
                                                    Case "razonSocialComprador"
                                                        Dim razonSocialComprador = infoNc.InnerText
                                                        Ncredito.RazonSocialComprador = razonSocialComprador
                                                    Case "identificacionComprador"
                                                        Dim identificacionComprador = infoNc.InnerText
                                                        Ncredito.IdentificacionComprador = identificacionComprador
                                                    Case "direccionComprador"
                                                        Dim direccionComprador = infoNc.InnerText
                                                       ' Ncredito.DireccionEstablecimiento = direccionComprador
                                                    Case "codDocModificado"
                                                        Dim codDocModificado = infoNc.InnerText
                                                        Ncredito.CodDocModificado = codDocModificado
                                                    Case "numDocModificado"
                                                        Dim numDocModificado = infoNc.InnerText
                                                        Ncredito.NumDocModificado = numDocModificado
                                                    Case "fechaEmisionDocSustento"
                                                        Dim fechaEmisionDocSustento = infoNc.InnerText
                                                        Ncredito.FechaEmisionDocModificado = CDate(fechaEmisionDocSustento)
                                                    Case "totalSinImpuestos"
                                                        Dim totalSinImpuestos = infoNc.InnerText
                                                        Ncredito.TotalSinImpuesto = totalSinImpuestos
                                                    Case "motivo"
                                                        Dim motivo = infoNc.InnerText
                                                        Ncredito.MotivoModificacion = motivo
                                                    Case "valorModificacion"
                                                        Dim valorModificacion = infoNc.InnerText
                                                        Ncredito.ValorModificacion = valorModificacion
                                                    Case "totalConImpuestos"
                                                        Dim XDocumentoImpuestoCabecera = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCreditoImpuesto)

                                                        For Each totalConImpuestos As XmlNode In infoNc.ChildNodes
                                                            Select Case totalConImpuestos.Name
                                                                Case "totalImpuesto"
                                                                    Dim impuestoCab As New Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCreditoImpuesto

                                                                    For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                                        Select Case totalImpuesto.Name
                                                                            Case "codigo"
                                                                                Dim codigo = totalImpuesto.InnerText
                                                                                impuestoCab.Codigo = CInt(codigo)
                                                                            Case "codigoPorcentaje"
                                                                                Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                                'NcCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                impuestoCab.CodigoPorcentaje = CInt(codigoPorcentaje)
                                                                            Case "baseImponible"
                                                                                Dim baseImponible = totalImpuesto.InnerText
                                                                                ' NcCabImp._baseImponible = baseImponible
                                                                                impuestoCab.BaseImponible = baseImponible
                                                                            Case "tarifa"
                                                                                Dim tarifa = totalImpuesto.InnerText
                                                                                ' NcCabImp._tarifa = tarifa
                                                                                impuestoCab.Tarifa = tarifa
                                                                            Case "valor"
                                                                                Dim valor = totalImpuesto.InnerText
                                                                                ' NcCabImp._valor = valor
                                                                                impuestoCab.Valor = valor
                                                                        End Select

                                                                    Next
                                                                    ' Nc.NotaCreditoCabecera._impuestos.Add(NcCabImp)
                                                                    XDocumentoImpuestoCabecera.Add(impuestoCab)
                                                            End Select
                                                        Next


                                                        If XDocumentoImpuestoCabecera.Count > 0 Then

                                                            Ncredito.ENTNotaCreditoImpuesto = XDocumentoImpuestoCabecera.ToArray

                                                        End If

                                                End Select
                                            Next
                                        Case "detalles"

                                            Dim XDetalles = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito)


                                            For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                                Select Case detalles.Name
                                                    Case "detalle"
                                                        'Dim NcDet As New NotaCreditoDetalle
                                                        'Dim NcDetImp As New NotaCreditoDetalleImpuesto
                                                        'NcDet._impuestos = New List(Of NotaCreditoDetalleImpuesto)
                                                        Dim FacDet As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito
                                                        For Each detalle As XmlNode In detalles.ChildNodes
                                                            Select Case detalle.Name
                                                                Case "codigoInterno"
                                                                    Dim codigoPrincipal = detalle.InnerText
                                                                    '  NcDet._codigoInterno = codigoPrincipal
                                                                    FacDet.CodigoPrincipal = codigoPrincipal
                                                                Case "codigoAdicional"
                                                                    Dim codigoAuxiliar = detalle.InnerText
                                                                    ' NcDet._codigoAdicional = codigoAuxiliar
                                                                    FacDet.CodigoAuxiliar = codigoAuxiliar
                                                                Case "descripcion"
                                                                    Dim descripcion = detalle.InnerText
                                                                    'NcDet._descripcion = descripcion
                                                                    FacDet.Descripcion = descripcion
                                                                Case "cantidad"
                                                                    Dim cantidad = detalle.InnerText
                                                                    '._cantidad = cantidad
                                                                    FacDet.Cantidad = cantidad
                                                                Case "precioUnitario"
                                                                    Dim precioUnitario = detalle.InnerText
                                                                    ' NcDet._precioUnitario = precioUnitario
                                                                    FacDet.PrecioUnitario = precioUnitario
                                                                Case "descuento"
                                                                    Dim descuento = detalle.InnerText
                                                                    'NcDet._descuento = descuento
                                                                    FacDet.Descuento = descuento
                                                                Case "precioTotalSinImpuesto"
                                                                    Dim precioTotalSinImpuesto = detalle.InnerText
                                                                    ' NcDet._precioTotalSinImpuesto = precioTotalSinImpuesto
                                                                    FacDet.PrecioTotalSinImpuesto = precioTotalSinImpuesto
                                                                Case "impuestos"
                                                                    Dim ListaImpuestoDetalle = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCreditoImpuesto)
                                                                    For Each impuestos As XmlNode In detalle.ChildNodes
                                                                        Select Case impuestos.Name
                                                                            Case "impuesto"
                                                                                Dim FacDetImp As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCreditoImpuesto

                                                                                For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                                    Select Case impuesto.Name
                                                                                        Case "codigo"
                                                                                            Dim codigo = impuesto.InnerText
                                                                                            'NcDetImp._codigo = CInt(codigo)
                                                                                            FacDetImp.Codigo = codigo
                                                                                        Case "codigoPorcentaje"
                                                                                            Dim codigoPorcentaje = impuesto.InnerText
                                                                                            ' NcDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                            FacDetImp.CodigoPorcentaje = codigoPorcentaje
                                                                                        Case "tarifa"
                                                                                            Dim tarifa = impuesto.InnerText
                                                                                            ' NcDetImp._tarifa = tarifa
                                                                                            FacDetImp.Tarifa = tarifa
                                                                                        Case "baseImponible"
                                                                                            Dim baseImponible = impuesto.InnerText
                                                                                            ' NcDetImp._baseImponible = baseImponible
                                                                                            FacDetImp.BaseImponible = baseImponible
                                                                                        Case "valor"
                                                                                            Dim valor = impuesto.InnerText
                                                                                            ' NcDetImp._valor = valor
                                                                                            FacDetImp.Valor = valor
                                                                                    End Select
                                                                                Next
                                                                                ListaImpuestoDetalle.Add(FacDetImp)
                                                                        End Select
                                                                    Next

                                                                    If ListaImpuestoDetalle.Count > 0 Then

                                                                        FacDet.ENTDetalleNotaCreditoImpuesto = ListaImpuestoDetalle.ToArray

                                                                    End If

                                                            End Select
                                                        Next

                                                        XDetalles.Add(FacDet)

                                                End Select
                                            Next

                                            If XDetalles.Count > 0 Then


                                                Ncredito.ENTDetalleNotaCredito = XDetalles.ToArray

                                            End If


                                    End Select
                                Next
                        End Select
                    Next

                Else


                    Dim notaCredito = m_xmld.SelectSingleNode("notaCredito")

                    For Each nodoInfTri As XmlNode In notaCredito.ChildNodes
                        Dim p = nodoInfTri.Name
                        Select Case p
                            Case "infoTributaria"
                                For Each n As XmlNode In nodoInfTri.ChildNodes
                                    Select Case n.Name
                                        Case "razonSocial"
                                            Dim razonSocial = n.InnerText.ToString
                                            Ncredito.RazonSocial = razonSocial
                                        Case "ruc"
                                            Dim ruc = n.InnerText.ToString
                                            Ncredito.Ruc = ruc
                                        Case "estab"
                                            Dim estab = n.InnerText.ToString
                                            Ncredito.Establecimiento = estab
                                        Case "ptoEmi"
                                            Dim ptoEmi = n.InnerText.ToString
                                            Ncredito.PuntoEmision = ptoEmi
                                        Case "secuencial"
                                            Dim secuencial = n.InnerText.ToString
                                            Ncredito.Secuencial = secuencial
                                        Case "claveAcceso"
                                            Dim claveAcceso = n.InnerText.ToString
                                            Ncredito.ClaveAcceso = claveAcceso
                                    End Select

                                Next
                            Case "infoNotaCredito"
                                For Each infoNc As XmlNode In nodoInfTri.ChildNodes
                                    Select Case infoNc.Name
                                        Case "fechaEmision"
                                            Dim fechaemision = infoNc.InnerText
                                            Ncredito.FechaEmision = fechaemision
                                        Case "dirEstablecimiento"
                                            Dim dirEstablecimiento = infoNc.InnerText
                                            Ncredito.DireccionEstablecimiento = dirEstablecimiento
                                        Case "razonSocialComprador"
                                            Dim razonSocialComprador = infoNc.InnerText
                                            Ncredito.RazonSocialComprador = razonSocialComprador
                                        Case "identificacionComprador"
                                            Dim identificacionComprador = infoNc.InnerText
                                            Ncredito.IdentificacionComprador = identificacionComprador
                                        Case "direccionComprador"
                                            Dim direccionComprador = infoNc.InnerText
                                                       ' Ncredito.DireccionEstablecimiento = direccionComprador
                                        Case "codDocModificado"
                                            Dim codDocModificado = infoNc.InnerText
                                            Ncredito.CodDocModificado = codDocModificado
                                        Case "numDocModificado"
                                            Dim numDocModificado = infoNc.InnerText
                                            Ncredito.NumDocModificado = numDocModificado
                                        Case "fechaEmisionDocSustento"
                                            Dim fechaEmisionDocSustento = infoNc.InnerText
                                            Ncredito.FechaEmisionDocModificado = CDate(fechaEmisionDocSustento)
                                        Case "totalSinImpuestos"
                                            Dim totalSinImpuestos = infoNc.InnerText
                                            Ncredito.TotalSinImpuesto = totalSinImpuestos
                                        Case "motivo"
                                            Dim motivo = infoNc.InnerText
                                            Ncredito.MotivoModificacion = motivo
                                        Case "valorModificacion"
                                            Dim valorModificacion = infoNc.InnerText
                                            Ncredito.ValorModificacion = valorModificacion
                                        Case "totalConImpuestos"
                                            Dim XDocumentoImpuestoCabecera = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCreditoImpuesto)

                                            For Each totalConImpuestos As XmlNode In infoNc.ChildNodes
                                                Select Case totalConImpuestos.Name
                                                    Case "totalImpuesto"
                                                        Dim impuestoCab As New Entidades.wsEDoc_ConsultaRecepcion.ENTNotaCreditoImpuesto

                                                        For Each totalImpuesto As XmlNode In totalConImpuestos.ChildNodes
                                                            Select Case totalImpuesto.Name
                                                                Case "codigo"
                                                                    Dim codigo = totalImpuesto.InnerText
                                                                    impuestoCab.Codigo = CInt(codigo)
                                                                Case "codigoPorcentaje"
                                                                    Dim codigoPorcentaje = totalImpuesto.InnerText
                                                                    'NcCabImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                    impuestoCab.CodigoPorcentaje = CInt(codigoPorcentaje)
                                                                Case "baseImponible"
                                                                    Dim baseImponible = totalImpuesto.InnerText
                                                                    ' NcCabImp._baseImponible = baseImponible
                                                                    impuestoCab.BaseImponible = baseImponible
                                                                Case "tarifa"
                                                                    Dim tarifa = totalImpuesto.InnerText
                                                                    ' NcCabImp._tarifa = tarifa
                                                                    impuestoCab.Tarifa = tarifa
                                                                Case "valor"
                                                                    Dim valor = totalImpuesto.InnerText
                                                                    ' NcCabImp._valor = valor
                                                                    impuestoCab.Valor = valor
                                                            End Select

                                                        Next
                                                        ' Nc.NotaCreditoCabecera._impuestos.Add(NcCabImp)
                                                        XDocumentoImpuestoCabecera.Add(impuestoCab)
                                                End Select
                                            Next


                                            If XDocumentoImpuestoCabecera.Count > 0 Then

                                                Ncredito.ENTNotaCreditoImpuesto = XDocumentoImpuestoCabecera.ToArray

                                            End If

                                    End Select
                                Next
                            Case "detalles"

                                Dim XDetalles = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito)


                                For Each detalles As XmlNode In nodoInfTri.ChildNodes
                                    Select Case detalles.Name
                                        Case "detalle"
                                            'Dim NcDet As New NotaCreditoDetalle
                                            'Dim NcDetImp As New NotaCreditoDetalleImpuesto
                                            'NcDet._impuestos = New List(Of NotaCreditoDetalleImpuesto)
                                            Dim FacDet As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCredito
                                            For Each detalle As XmlNode In detalles.ChildNodes
                                                Select Case detalle.Name
                                                    Case "codigoInterno"
                                                        Dim codigoPrincipal = detalle.InnerText
                                                        '  NcDet._codigoInterno = codigoPrincipal
                                                        FacDet.CodigoPrincipal = codigoPrincipal
                                                    Case "codigoAdicional"
                                                        Dim codigoAuxiliar = detalle.InnerText
                                                        ' NcDet._codigoAdicional = codigoAuxiliar
                                                        FacDet.CodigoAuxiliar = codigoAuxiliar
                                                    Case "descripcion"
                                                        Dim descripcion = detalle.InnerText
                                                        'NcDet._descripcion = descripcion
                                                        FacDet.Descripcion = descripcion
                                                    Case "cantidad"
                                                        Dim cantidad = detalle.InnerText
                                                        '._cantidad = cantidad
                                                        FacDet.Cantidad = cantidad
                                                    Case "precioUnitario"
                                                        Dim precioUnitario = detalle.InnerText
                                                        ' NcDet._precioUnitario = precioUnitario
                                                        FacDet.PrecioUnitario = precioUnitario
                                                    Case "descuento"
                                                        Dim descuento = detalle.InnerText
                                                        'NcDet._descuento = descuento
                                                        FacDet.Descuento = descuento
                                                    Case "precioTotalSinImpuesto"
                                                        Dim precioTotalSinImpuesto = detalle.InnerText
                                                        ' NcDet._precioTotalSinImpuesto = precioTotalSinImpuesto
                                                        FacDet.PrecioTotalSinImpuesto = precioTotalSinImpuesto
                                                    Case "impuestos"
                                                        Dim ListaImpuestoDetalle = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCreditoImpuesto)
                                                        For Each impuestos As XmlNode In detalle.ChildNodes
                                                            Select Case impuestos.Name
                                                                Case "impuesto"
                                                                    Dim FacDetImp As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleNotaCreditoImpuesto

                                                                    For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                        Select Case impuesto.Name
                                                                            Case "codigo"
                                                                                Dim codigo = impuesto.InnerText
                                                                                'NcDetImp._codigo = CInt(codigo)
                                                                                FacDetImp.Codigo = codigo
                                                                            Case "codigoPorcentaje"
                                                                                Dim codigoPorcentaje = impuesto.InnerText
                                                                                ' NcDetImp._codigoPorcentaje = CInt(codigoPorcentaje)
                                                                                FacDetImp.CodigoPorcentaje = codigoPorcentaje
                                                                            Case "tarifa"
                                                                                Dim tarifa = impuesto.InnerText
                                                                                ' NcDetImp._tarifa = tarifa
                                                                                FacDetImp.Tarifa = tarifa
                                                                            Case "baseImponible"
                                                                                Dim baseImponible = impuesto.InnerText
                                                                                ' NcDetImp._baseImponible = baseImponible
                                                                                FacDetImp.BaseImponible = baseImponible
                                                                            Case "valor"
                                                                                Dim valor = impuesto.InnerText
                                                                                ' NcDetImp._valor = valor
                                                                                FacDetImp.Valor = valor
                                                                        End Select
                                                                    Next
                                                                    ListaImpuestoDetalle.Add(FacDetImp)
                                                            End Select
                                                        Next

                                                        If ListaImpuestoDetalle.Count > 0 Then

                                                            FacDet.ENTDetalleNotaCreditoImpuesto = ListaImpuestoDetalle.ToArray

                                                        End If

                                                End Select
                                            Next

                                            XDetalles.Add(FacDet)

                                    End Select
                                Next

                                If XDetalles.Count > 0 Then


                                    Ncredito.ENTDetalleNotaCredito = XDetalles.ToArray

                                End If


                        End Select
                    Next

                End If
                'GuardaLog("NC clave de acceso: " + Nc.NotaCreditoCabecera._claveAcceso + " leido Correctamente!!" + " nombre del archivo: " + Path.GetFileName(ruta))

                Utilitario.Util_Log.Escribir_Log("NC clave de acceso: " + Ncredito.ClaveAcceso + " leido Correctamente!!" + " nombre del archivo: ", " OperacionesXML")


                Return Ncredito
            Catch ex As Exception
                ' GuardaLog("Error al leer XML NC  ==> " + ex.Message + " con nombre: " + Path.GetFileName(ruta))

                Utilitario.Util_Log.Escribir_Log("Error al leer XML NC  ==> " + ex.Message + " con nombre: ", " OperacionesXML")



                Return Nothing
            End Try

        End Function

        Public Shared Function validaSiEmpiezaPunto(ByVal valor As String) As String
            Dim _valor As String = valor

            Try
                If valor.StartsWith(".") Then
                    _valor = "0" & valor

                End If
            Catch ex As Exception

                _valor = Nothing
            End Try

            Return _valor
        End Function
        Public Shared Function LeerXMLRetencion2(ruta As String) As Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion

            Try

                'GuardaLog("Consultando XML Retencion a procesar..")

                Utilitario.Util_Log.Escribir_Log("Consultando XML Retencion a procesar..", " OperacionesXML")



                'Dim _RUTA As String = "C:\Users\David Macias\Documents\ECUADOR\ProyectoServicioXMLRecepcion\FUNCIONES LEER XML\ret.xml"

                Dim mensaje As String = ""

                'Dim ms As New MemoryStream


                Dim m_xmld As XmlDocument
                Dim m_nodelist As XmlNodeList

                m_xmld = New XmlDocument()

                m_xmld.Load(ruta)




                Dim nodoAutSchema = m_xmld.SelectSingleNode("autorizacion")


                Dim nodoAut = m_xmld.SelectSingleNode("autorizacion")

                Dim _nodoAut = m_xmld.SelectSingleNode("respuestaComprobante/autorizaciones/autorizacion")

                If (Not (nodoAutSchema) Is Nothing) Then

                    Dim Retencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion


                    For Each facAut As XmlNode In nodoAut.ChildNodes
                        Select Case facAut.Name
                            Case "numeroAutorizacion"
                                Retencion.AutorizacionSRI = facAut.InnerText.ToString
                            Case "fechaAutorizacion"
                                Retencion.FechaAutorizacion = facAut.InnerText
                            Case "comprobante"
                                For Each comp As XmlNode In facAut.ChildNodes
                                    Select Case comp.Name
                                        Case "comprobanteRetencion"
                                            For Each compRet As XmlNode In comp.ChildNodes
                                                Select Case compRet.Name
                                                    Case "infoTributaria"
                                                        For Each infoTri As XmlNode In compRet.ChildNodes
                                                            Select Case infoTri.Name
                                                                Case "razonSocial"
                                                                    Dim razonSocial = infoTri.InnerText.ToString
                                                                    ' Retencion.RetCabecera._RazonSocial = razonSocial
                                                                    Retencion.RazonSocial = razonSocial
                                                                Case "ruc"
                                                                    Dim ruc = infoTri.InnerText.ToString
                                                                    'Retencion.RetCabecera._ruc = ruc
                                                                    Retencion.Ruc = ruc
                                                                Case "estab"
                                                                    Dim estab = infoTri.InnerText.ToString
                                                                    'Retencion.RetCabecera._estab = estab
                                                                    Retencion.Establecimiento = estab
                                                                Case "ptoEmi"
                                                                    Dim ptoEmi = infoTri.InnerText.ToString
                                                                    'Retencion.RetCabecera._ptoEmi = ptoEmi
                                                                    Retencion.PuntoEmision = ptoEmi
                                                                Case "secuencial"
                                                                    Dim secuencial = infoTri.InnerText.ToString
                                                                    ' Retencion.RetCabecera._secuencial = secuencial
                                                                    Retencion.Secuencial = secuencial
                                                                Case "claveAcceso"
                                                                    Dim claveAcceso = infoTri.InnerText.ToString
                                                                    ' Retencion.RetCabecera._claveAcceso = claveAcceso
                                                                    Retencion.ClaveAcceso = claveAcceso
                                                            End Select
                                                        Next

                                                    Case "infoCompRetencion"
                                                        For Each infoRt As XmlNode In compRet.ChildNodes
                                                            Select Case infoRt.Name
                                                                Case "fechaEmision"
                                                                    ' Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                                                    Retencion.FechaEmision = infoRt.InnerText.ToString
                                                                Case "dirEstablecimiento"
                                                                    'Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                                                    Retencion.DireccionEstablecimiento = infoRt.InnerText.ToString
                                                                Case "razonSocialSujetoRetenido"
                                                                    'Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                                    Retencion.RazonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                                Case "identificacionSujetoRetenido"
                                                                    ' Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                                                    Retencion.IdentificacionSujetoRetenido = infoRt.InnerText.ToString
                                                                Case "periodoFiscal"
                                                                    Retencion.PeriodoFiscal = infoRt.InnerText.ToString
                                                                Case "impuestos"
                                                            End Select
                                                        Next

                                                    Case "impuestos"
                                                        Dim DetallesImpuestosRetencion = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion)

                                                        For Each impuestos As XmlNode In compRet.ChildNodes
                                                            Select Case impuestos.Name
                                                                Case "impuesto"
                                                                    'Dim RTDetImp As New RetDetalleImpuestos
                                                                    Dim DetImpuestoRetencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion

                                                                    For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                        Select Case impuesto.Name
                                                                            Case "codigo"
                                                                                ' RTDetImp._codigo = CInt(impuesto.InnerText)
                                                                                DetImpuestoRetencion.Codigo = CInt(impuesto.InnerText)
                                                                            Case "codigoRetencion"
                                                                                ' RTDetImp._codigoRetencion = impuesto.InnerText
                                                                                DetImpuestoRetencion.CodigoRetencion = impuesto.InnerText
                                                                            Case "baseImponible"
                                                                                ' RTDetImp._baseImponible = impuesto.InnerText
                                                                                DetImpuestoRetencion.BaseImponible = impuesto.InnerText
                                                                            Case "porcentajeRetener"
                                                                                ' RTDetImp._porcentajeRetener = impuesto.InnerText
                                                                                DetImpuestoRetencion.PorcentajeRetener = impuesto.InnerText
                                                                            Case "valorRetenido"
                                                                                ' RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                                                DetImpuestoRetencion.ValorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)

                                                                                Retencion.TotalRetencion = Retencion.TotalRetencion + DetImpuestoRetencion.ValorRetenido
                                                                            Case "codDocSustento"
                                                                                ' RTDetImp._codDocSustento = impuesto.InnerText
                                                                                DetImpuestoRetencion.CodDocRetener = impuesto.InnerText
                                                                            Case "numDocSustento"
                                                                                ' RTDetImp._numDocSustento = impuesto.InnerText
                                                                                DetImpuestoRetencion.NumDocRetener = impuesto.InnerText
                                                                            Case "fechaEmisionDocSustento"
                                                                                ' RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                                                DetImpuestoRetencion.FechaEmisionDocRetener = impuesto.InnerText
                                                                        End Select
                                                                    Next
                                                                    'Retencion.RetDetalleImp.Add(RTDetImp)
                                                                    DetallesImpuestosRetencion.Add(DetImpuestoRetencion)
                                                            End Select
                                                        Next

                                                        If DetallesImpuestosRetencion.Count > 0 Then

                                                            Retencion.ENTDetalleRetencion = DetallesImpuestosRetencion.ToArray

                                                        End If

                                                End Select


                                            Next
                                    End Select
                                Next

                        End Select
                    Next


                    If Not IsNothing(Retencion.RazonSocial) Then
                        Return Retencion
                    End If

                End If

                If (Not (_nodoAut) Is Nothing) Then

                    Dim Retencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
                    'Retencion.RetCabecera = New RetCabecera
                    'Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

                    For Each facAut As XmlNode In _nodoAut.ChildNodes
                        Select Case facAut.Name
                            Case "numeroAutorizacion"
                                'Retencion.RetCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                                Retencion.AutorizacionSRI = facAut.InnerText.ToString
                            Case "fechaAutorizacion"
                                ' Retencion.RetCabecera._FechaAutorizacion = facAut.InnerText.ToString
                                Retencion.FechaAutorizacion = facAut.InnerText.ToString
                        End Select
                    Next

                    Dim nodo = m_xmld.SelectSingleNode("respuestaComprobante/autorizaciones/autorizacion/comprobante")

                    Dim comprobante = nodo.InnerText

                    Dim nodoRet As New XmlDocument()
                    nodoRet.LoadXml(comprobante)


                    For Each comp As XmlNode In nodoRet.ChildNodes
                        Select Case comp.Name
                            Case "comprobanteRetencion"
                                For Each info As XmlNode In comp.ChildNodes
                                    Select Case info.Name
                                        Case "infoTributaria"
                                            For Each infoTri As XmlNode In info.ChildNodes
                                                Select Case infoTri.Name
                                                    Case "razonSocial"
                                                        Dim razonSocial = infoTri.InnerText.ToString
                                                        ' Retencion.RetCabecera._RazonSocial = razonSocial
                                                        Retencion.RazonSocial = razonSocial
                                                    Case "ruc"
                                                        Dim ruc = infoTri.InnerText.ToString
                                                        'Retencion.RetCabecera._ruc = ruc
                                                        Retencion.Ruc = ruc
                                                    Case "estab"
                                                        Dim estab = infoTri.InnerText.ToString
                                                        ' Retencion.RetCabecera._estab = estab
                                                        Retencion.Establecimiento = estab

                                                    Case "ptoEmi"
                                                        Dim ptoEmi = infoTri.InnerText.ToString
                                                        'Retencion.RetCabecera._ptoEmi = ptoEmi
                                                        Retencion.PuntoEmision = ptoEmi
                                                    Case "secuencial"
                                                        Dim secuencial = infoTri.InnerText.ToString
                                                        'Retencion.RetCabecera._secuencial = secuencial
                                                        Retencion.Secuencial = secuencial
                                                    Case "claveAcceso"
                                                        Dim claveAcceso = infoTri.InnerText.ToString
                                                        'Retencion.RetCabecera._claveAcceso = claveAcceso
                                                        Retencion.ClaveAcceso = claveAcceso
                                                End Select
                                            Next

                                        Case "infoCompRetencion"
                                            For Each infoRt As XmlNode In info.ChildNodes
                                                Select Case infoRt.Name
                                                    Case "fechaEmision"
                                                        'Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                                        Retencion.FechaEmision = infoRt.InnerText.ToString
                                                    Case "dirEstablecimiento"
                                                        ' Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                                        Retencion.DireccionEstablecimiento = infoRt.InnerText.ToString

                                                    Case "razonSocialSujetoRetenido"
                                                        'Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                        Retencion.RazonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                    Case "identificacionSujetoRetenido"
                                                        ' Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                                        Retencion.IdentificacionSujetoRetenido = infoRt.InnerText.ToString
                                                    Case "periodoFiscal"
                                                        ' Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                                        Retencion.PeriodoFiscal = infoRt.InnerText.ToString
                                                    Case "impuestos"
                                                End Select
                                            Next

                                        Case "impuestos"
                                            Dim DetallesImpuestosRetencion = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion)

                                            For Each impuestos As XmlNode In info.ChildNodes
                                                Select Case impuestos.Name
                                                    Case "impuesto"
                                                        'Dim RTDetImp As New RetDetalleImpuestos
                                                        Dim DetImpuestoRetencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion

                                                        For Each impuesto As XmlNode In impuestos.ChildNodes
                                                            Select Case impuesto.Name
                                                                Case "codigo"
                                                                    ' RTDetImp._codigo = CInt(impuesto.InnerText)
                                                                    DetImpuestoRetencion.Codigo = CInt(impuesto.InnerText)
                                                                Case "codigoRetencion"
                                                                    ' RTDetImp._codigoRetencion = impuesto.InnerText
                                                                    DetImpuestoRetencion.CodigoRetencion = impuesto.InnerText
                                                                Case "baseImponible"
                                                                    ' RTDetImp._baseImponible = impuesto.InnerText
                                                                    DetImpuestoRetencion.BaseImponible = impuesto.InnerText
                                                                Case "porcentajeRetener"
                                                                    ' RTDetImp._porcentajeRetener = impuesto.InnerText
                                                                    DetImpuestoRetencion.PorcentajeRetener = impuesto.InnerText
                                                                Case "valorRetenido"
                                                                    ' RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                                    DetImpuestoRetencion.ValorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)

                                                                    Retencion.TotalRetencion = Retencion.TotalRetencion + DetImpuestoRetencion.ValorRetenido
                                                                Case "codDocSustento"
                                                                    ' RTDetImp._codDocSustento = impuesto.InnerText
                                                                    DetImpuestoRetencion.CodDocRetener = impuesto.InnerText
                                                                Case "numDocSustento"
                                                                    ' RTDetImp._numDocSustento = impuesto.InnerText
                                                                    DetImpuestoRetencion.NumDocRetener = impuesto.InnerText
                                                                Case "fechaEmisionDocSustento"
                                                                    ' RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                                    DetImpuestoRetencion.FechaEmisionDocRetener = impuesto.InnerText
                                                            End Select
                                                        Next
                                                        'Retencion.RetDetalleImp.Add(RTDetImp)
                                                        DetallesImpuestosRetencion.Add(DetImpuestoRetencion)

                                                End Select
                                            Next

                                            If DetallesImpuestosRetencion.Count > 0 Then

                                                Retencion.ENTDetalleRetencion = DetallesImpuestosRetencion.ToArray

                                            End If

                                    End Select
                                Next
                        End Select
                    Next

                    Utilitario.Util_Log.Escribir_Log("Retencion clave de acceso: " + Retencion.AutorizacionSRI + " leido Correctamente!!", "OperacionesXML")
                    Return Retencion

                End If


                If (Not (nodoAut) Is Nothing) Then

                    Dim Retencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
                    'Retencion.RetCabecera = New RetCabecera
                    'Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)
                    Dim DetallesImpuestosRetencion = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion)

                    For Each facAut As XmlNode In nodoAut.ChildNodes
                        Select Case facAut.Name
                            Case "numeroAutorizacion"
                                'Retencion.RetCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                                Retencion.AutorizacionSRI = facAut.InnerText.ToString
                            Case "fechaAutorizacion"
                                ' Retencion.RetCabecera._FechaAutorizacion = facAut.InnerText.ToString
                                Retencion.FechaAutorizacion = facAut.InnerText.ToString
                        End Select
                    Next

                    Dim nodo = m_xmld.SelectSingleNode("autorizacion/comprobante")

                    Dim comprobante = nodo.InnerText

                    Dim nodoRet As New XmlDocument()
                    nodoRet.LoadXml(comprobante) 'loadxml leo el xml guardado en una vriable

                    'Dim razonSocial As String = ""
                    'Dim ruc As String = ""
                    'Dim estab As String = ""
                    'Dim ptoEmi As String = ""
                    'Dim secuencial As String = ""


                    For Each Ret As XmlNode In nodoRet.ChildNodes
                        Select Case Ret.Name
                            Case "comprobanteRetencion"
                                For Each nodoInfTri As XmlNode In Ret.ChildNodes
                                    Select Case nodoInfTri.Name
                                        Case "infoTributaria"
                                            For Each infoTri As XmlNode In nodoInfTri.ChildNodes
                                                Select Case infoTri.Name
                                                    Case "razonSocial"
                                                        Dim razonSocial = infoTri.InnerText.ToString
                                                        ' Retencion.RetCabecera._RazonSocial = razonSocial
                                                        Retencion.RazonSocial = razonSocial
                                                    Case "ruc"
                                                        Dim ruc = infoTri.InnerText.ToString
                                                        'Retencion.RetCabecera._ruc = ruc
                                                        Retencion.Ruc = ruc
                                                    Case "estab"
                                                        Dim estab = infoTri.InnerText.ToString
                                                        ' Retencion.RetCabecera._estab = estab
                                                        Retencion.Establecimiento = estab

                                                    Case "ptoEmi"
                                                        Dim ptoEmi = infoTri.InnerText.ToString
                                                        'Retencion.RetCabecera._ptoEmi = ptoEmi
                                                        Retencion.PuntoEmision = ptoEmi
                                                    Case "secuencial"
                                                        Dim secuencial = infoTri.InnerText.ToString
                                                        'Retencion.RetCabecera._secuencial = secuencial
                                                        Retencion.Secuencial = secuencial
                                                    Case "claveAcceso"
                                                        Dim claveAcceso = infoTri.InnerText.ToString
                                                        'Retencion.RetCabecera._claveAcceso = claveAcceso
                                                        Retencion.ClaveAcceso = claveAcceso
                                                End Select
                                            Next
                                        Case "infoCompRetencion"
                                            For Each infoRt As XmlNode In nodoInfTri.ChildNodes
                                                Select Case infoRt.Name
                                                    Case "fechaEmision"
                                                        Retencion.FechaEmision = infoRt.InnerText.ToString
                                                    Case "dirEstablecimiento"
                                                        Retencion.DireccionEstablecimiento = infoRt.InnerText.ToString
                                                    Case "razonSocialSujetoRetenido"
                                                        Retencion.RazonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                    Case "identificacionSujetoRetenido"
                                                        Retencion.IdentificacionSujetoRetenido = infoRt.InnerText.ToString
                                                    Case "periodoFiscal"
                                                        Retencion.PeriodoFiscal = infoRt.InnerText.ToString
                                                    Case "impuestos"
                                                End Select
                                            Next

                                        Case "docsSustento"
                                            Dim codDocSustento As String = ""
                                            Dim _numDocSustento As String = ""
                                            Dim fechaEmisionDocSustento As String = ""
                                            For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                                Select Case impuestos.Name
                                                    Case "docSustento"
                                                        For Each docSustento As XmlNode In impuestos.ChildNodes
                                                            Select Case docSustento.Name

                                                                Case "codDocSustento"
                                                                    codDocSustento = docSustento.InnerText
                                                                Case "numDocSustento"
                                                                    _numDocSustento = docSustento.InnerText
                                                                Case "fechaEmisionDocSustento"
                                                                    fechaEmisionDocSustento = docSustento.InnerText

                                                                Case "retenciones"
                                                                    For Each retenciones As XmlNode In docSustento.ChildNodes
                                                                        Select Case retenciones.Name
                                                                            Case "retencion"
                                                                                'Dim RTDetImp As New RetDetalleImpuestos
                                                                                Dim DetImpuestoRetencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion

                                                                                For Each _retencion As XmlNode In retenciones.ChildNodes
                                                                                    Select Case _retencion.Name
                                                                                        Case "codigo"
                                                                                            'RTDetImp._codigo = CInt(_retencion.InnerText)
                                                                                            DetImpuestoRetencion.Codigo = CInt(_retencion.InnerText)
                                                                                        Case "codigoRetencion"
                                                                                            ' RTDetImp._codigoRetencion = _retencion.InnerText
                                                                                            DetImpuestoRetencion.CodigoRetencion = _retencion.InnerText
                                                                                        Case "baseImponible"
                                                                                            'RTDetImp._baseImponible = _retencion.InnerText
                                                                                            DetImpuestoRetencion.BaseImponible = _retencion.InnerText

                                                                                        Case "porcentajeRetener"
                                                                                            'RTDetImp._porcentajeRetener = _retencion.InnerText
                                                                                            DetImpuestoRetencion.PorcentajeRetener = _retencion.InnerText
                                                                                        Case "valorRetenido"
                                                                                            'RTDetImp._valorRetenido = validaSiEmpiezaPunto(_retencion.InnerText)
                                                                                            DetImpuestoRetencion.ValorRetenido = validaSiEmpiezaPunto(_retencion.InnerText)
                                                                                            Retencion.TotalRetencion = Retencion.TotalRetencion + DetImpuestoRetencion.ValorRetenido
                                                                                            'Case "codDocSustento"
                                                                                            '    RTDetImp._codDocSustento = _retencion.InnerText
                                                                                            'Case "numDocSustento"
                                                                                            '    RTDetImp._numDocSustento = _retencion.InnerText
                                                                                            'Case "fechaEmisionDocSustento"
                                                                                            '    RTDetImp._fechaEmisionDocSustento = _retencion.InnerText.ToString
                                                                                            DetImpuestoRetencion.CodDocRetener = codDocSustento
                                                                                            DetImpuestoRetencion.NumDocRetener = _numDocSustento
                                                                                            DetImpuestoRetencion.FechaEmisionDocRetener = fechaEmisionDocSustento
                                                                                    End Select
                                                                                Next
                                                                                DetallesImpuestosRetencion.Add(DetImpuestoRetencion)
                                                                        End Select

                                                                    Next


                                                            End Select
                                                        Next

                                                End Select
                                            Next

                                        Case "impuestos"

                                            For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                                Select Case impuestos.Name
                                                    Case "impuesto"
                                                        'Dim RTDetImp As New RetDetalleImpuestos
                                                        Dim DetImpuestoRetencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion

                                                        For Each impuesto As XmlNode In impuestos.ChildNodes
                                                            Select Case impuesto.Name
                                                                Case "codigo"
                                                                    'RTDetImp._codigo = CInt(impuesto.InnerText)
                                                                    DetImpuestoRetencion.Codigo = CInt(impuesto.InnerText)

                                                                Case "codigoRetencion"
                                                                    ' RTDetImp._codigoRetencion = impuesto.InnerText
                                                                    DetImpuestoRetencion.CodigoRetencion = impuesto.InnerText
                                                                Case "baseImponible"
                                                                    'RTDetImp._baseImponible = impuesto.InnerText
                                                                    DetImpuestoRetencion.BaseImponible = impuesto.InnerText

                                                                Case "porcentajeRetener"
                                                                    'RTDetImp._porcentajeRetener = impuesto.InnerText
                                                                    DetImpuestoRetencion.PorcentajeRetener = impuesto.InnerText
                                                                Case "valorRetenido"
                                                                    'RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                                    DetImpuestoRetencion.ValorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                                    Retencion.TotalRetencion = Retencion.TotalRetencion + DetImpuestoRetencion.ValorRetenido
                                                                Case "codDocSustento"
                                                                    ' RTDetImp._codDocSustento = impuesto.InnerText
                                                                    DetImpuestoRetencion.CodDocRetener = impuesto.InnerText
                                                                Case "numDocSustento"
                                                                    ' RTDetImp._numDocSustento = impuesto.InnerText
                                                                    DetImpuestoRetencion.NumDocRetener = impuesto.InnerText
                                                                Case "fechaEmisionDocSustento"
                                                                    'RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                                    DetImpuestoRetencion.FechaEmisionDocRetener = impuesto.InnerText.ToString
                                                            End Select
                                                        Next
                                                        DetallesImpuestosRetencion.Add(DetImpuestoRetencion)
                                                End Select
                                            Next

                                    End Select
                                Next

                        End Select
                    Next

                    If DetallesImpuestosRetencion.Count > 0 Then

                        Retencion.ENTDetalleRetencion = DetallesImpuestosRetencion.ToArray

                    End If

                    Utilitario.Util_Log.Escribir_Log("Retencion clave de acceso: " + Retencion.AutorizacionSRI + " leido Correctamente!!", "OperacionesXML")
                    Return Retencion

                Else

                    Dim rt = m_xmld.SelectSingleNode("comprobanteRetencion")
                    If (Not (nodoAut) Is Nothing) Then

                        Dim Retencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
                        'Retencion.RetCabecera = New RetCabecera
                        'Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

                        For Each nodoInfTri As XmlNode In rt.ChildNodes
                            Select Case nodoInfTri.Name
                                Case "infoTributaria"
                                    For Each infoTri As XmlNode In nodoInfTri.ChildNodes
                                        Select Case infoTri.Name
                                            Case "razonSocial"
                                                Dim razonSocial = infoTri.InnerText.ToString
                                                ' Retencion.RetCabecera._RazonSocial = razonSocial
                                                Retencion.RazonSocial = razonSocial
                                            Case "ruc"
                                                Dim ruc = infoTri.InnerText.ToString
                                                'Retencion.RetCabecera._ruc = ruc
                                                Retencion.Ruc = ruc
                                            Case "estab"
                                                Dim estab = infoTri.InnerText.ToString
                                                ' Retencion.RetCabecera._estab = estab
                                                Retencion.Establecimiento = estab

                                            Case "ptoEmi"
                                                Dim ptoEmi = infoTri.InnerText.ToString
                                                'Retencion.RetCabecera._ptoEmi = ptoEmi
                                                Retencion.PuntoEmision = ptoEmi
                                            Case "secuencial"
                                                Dim secuencial = infoTri.InnerText.ToString
                                                'Retencion.RetCabecera._secuencial = secuencial
                                                Retencion.Secuencial = secuencial
                                            Case "claveAcceso"
                                                Dim claveAcceso = infoTri.InnerText.ToString
                                                'Retencion.RetCabecera._claveAcceso = claveAcceso
                                                Retencion.ClaveAcceso = claveAcceso
                                        End Select
                                    Next
                                Case "infoCompRetencion"
                                    For Each infoRt As XmlNode In nodoInfTri.ChildNodes
                                        Select Case infoRt.Name
                                            Case "fechaEmision"
                                                'Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                                Retencion.FechaEmision = infoRt.InnerText.ToString
                                            Case "dirEstablecimiento"
                                                ' Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                                Retencion.DireccionEstablecimiento = infoRt.InnerText.ToString

                                            Case "razonSocialSujetoRetenido"
                                                'Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                Retencion.RazonSocialSujetoRetenido = infoRt.InnerText.ToString
                                            Case "identificacionSujetoRetenido"
                                                ' Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                                Retencion.IdentificacionSujetoRetenido = infoRt.InnerText.ToString
                                            Case "periodoFiscal"
                                                ' Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                                Retencion.PeriodoFiscal = infoRt.InnerText.ToString
                                            Case "impuestos"
                                        End Select
                                    Next
                                Case "impuestos"
                                    Dim DetallesImpuestosRetencion = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion)

                                    For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                        Select Case impuestos.Name
                                            Case "impuesto"
                                                'Dim RTDetImp As New RetDetalleImpuestos
                                                Dim DetImpuestoRetencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion

                                                For Each impuesto As XmlNode In impuestos.ChildNodes
                                                    Select Case impuesto.Name
                                                        Case "codigo"
                                                            ' RTDetImp._codigo = CInt(impuesto.InnerText)
                                                            DetImpuestoRetencion.Codigo = CInt(impuesto.InnerText)
                                                        Case "codigoRetencion"
                                                            ' RTDetImp._codigoRetencion = impuesto.InnerText
                                                            DetImpuestoRetencion.CodigoRetencion = impuesto.InnerText
                                                        Case "baseImponible"
                                                            ' RTDetImp._baseImponible = impuesto.InnerText
                                                            DetImpuestoRetencion.BaseImponible = impuesto.InnerText
                                                        Case "porcentajeRetener"
                                                            ' RTDetImp._porcentajeRetener = impuesto.InnerText
                                                            DetImpuestoRetencion.PorcentajeRetener = impuesto.InnerText
                                                        Case "valorRetenido"
                                                            ' RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                            DetImpuestoRetencion.ValorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                            Retencion.TotalRetencion = Retencion.TotalRetencion + DetImpuestoRetencion.ValorRetenido

                                                        Case "codDocSustento"
                                                            ' RTDetImp._codDocSustento = impuesto.InnerText
                                                            DetImpuestoRetencion.CodDocRetener = impuesto.InnerText
                                                        Case "numDocSustento"
                                                            ' RTDetImp._numDocSustento = impuesto.InnerText
                                                            DetImpuestoRetencion.NumDocRetener = impuesto.InnerText
                                                        Case "fechaEmisionDocSustento"
                                                            ' RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                            DetImpuestoRetencion.FechaEmisionDocRetener = impuesto.InnerText
                                                    End Select
                                                Next
                                                DetallesImpuestosRetencion.Add(DetImpuestoRetencion)
                                        End Select
                                    Next

                                    If DetallesImpuestosRetencion.Count > 0 Then

                                        Retencion.ENTDetalleRetencion = DetallesImpuestosRetencion.ToArray

                                    End If

                            End Select
                        Next
                        Utilitario.Util_Log.Escribir_Log("Retencion clave de acceso: " + Retencion.AutorizacionSRI + " leido Correctamente!!", "OperacionesXML")
                        Return Retencion
                    Else

                        Dim Retencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTRetencion
                        'Retencion.RetCabecera = New RetCabecera
                        'Retencion.RetDetalleImp = New List(Of RetDetalleImpuestos)

                        Dim lector = New XmlDocument()
                        lector.Load(ruta)

                        Dim cvev As String = lector.GetNamespaceOfPrefix("http://ec.gob.sri.ws.autorizacion")
                        Dim xmlns As New Xml.XmlNamespaceManager(lector.NameTable)
                        xmlns.AddNamespace("q1", "http://ec.gob.sri.ws.autorizacion")
                        Dim xnodo As Xml.XmlNode
                        xnodo = lector.SelectSingleNode("/q1:respuestaComprobante/autorizaciones/autorizacion", xmlns)

                        If (Not (xnodo) Is Nothing) Then

                            For Each facAut As XmlNode In xnodo.ChildNodes
                                Select Case facAut.Name
                                    Case "numeroAutorizacion"
                                        'Retencion.RetCabecera._NumeroAutorizacion = facAut.InnerText.ToString
                                        Retencion.AutorizacionSRI = facAut.InnerText.ToString
                                    Case "fechaAutorizacion"
                                        ' Retencion.RetCabecera._FechaAutorizacion = facAut.InnerText.ToString
                                        Retencion.FechaAutorizacion = facAut.InnerText.ToString
                                End Select
                            Next

                            Dim nodo = lector.SelectSingleNode("/q1:respuestaComprobante/autorizaciones/autorizacion/comprobante", xmlns)

                            Dim comprobante = nodo.InnerText

                            Dim nodoRet As New XmlDocument()
                            nodoRet.LoadXml(comprobante)

                            For Each Ret As XmlNode In nodoRet.ChildNodes
                                Select Case Ret.Name
                                    Case "comprobanteRetencion"
                                        For Each nodoInfTri As XmlNode In Ret.ChildNodes
                                            Select Case nodoInfTri.Name
                                                Case "infoTributaria"
                                                    For Each infoTri As XmlNode In nodoInfTri.ChildNodes
                                                        Select Case infoTri.Name
                                                            Case "razonSocial"
                                                                Dim razonSocial = infoTri.InnerText.ToString
                                                                ' Retencion.RetCabecera._RazonSocial = razonSocial
                                                                Retencion.RazonSocial = razonSocial
                                                            Case "ruc"
                                                                Dim ruc = infoTri.InnerText.ToString
                                                                'Retencion.RetCabecera._ruc = ruc
                                                                Retencion.Ruc = ruc
                                                            Case "estab"
                                                                Dim estab = infoTri.InnerText.ToString
                                                                ' Retencion.RetCabecera._estab = estab
                                                                Retencion.Establecimiento = estab

                                                            Case "ptoEmi"
                                                                Dim ptoEmi = infoTri.InnerText.ToString
                                                                'Retencion.RetCabecera._ptoEmi = ptoEmi
                                                                Retencion.PuntoEmision = ptoEmi
                                                            Case "secuencial"
                                                                Dim secuencial = infoTri.InnerText.ToString
                                                                'Retencion.RetCabecera._secuencial = secuencial
                                                                Retencion.Secuencial = secuencial
                                                            Case "claveAcceso"
                                                                Dim claveAcceso = infoTri.InnerText.ToString
                                                                'Retencion.RetCabecera._claveAcceso = claveAcceso
                                                                Retencion.ClaveAcceso = claveAcceso
                                                        End Select
                                                    Next

                                                Case "infoCompRetencion"
                                                    For Each infoRt As XmlNode In nodoInfTri.ChildNodes
                                                        Select Case infoRt.Name
                                                            Case "fechaEmision"
                                                                'Retencion.RetCabecera._fechaEmision = infoRt.InnerText.ToString
                                                                Retencion.FechaEmision = infoRt.InnerText.ToString
                                                            Case "dirEstablecimiento"
                                                                ' Retencion.RetCabecera._dirEstablecimiento = infoRt.InnerText.ToString
                                                                Retencion.DireccionEstablecimiento = infoRt.InnerText.ToString

                                                            Case "razonSocialSujetoRetenido"
                                                                'Retencion.RetCabecera._razonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                                Retencion.RazonSocialSujetoRetenido = infoRt.InnerText.ToString
                                                            Case "identificacionSujetoRetenido"
                                                                ' Retencion.RetCabecera._identificacionSujetoRetenido = infoRt.InnerText.ToString
                                                                Retencion.IdentificacionSujetoRetenido = infoRt.InnerText.ToString
                                                            Case "periodoFiscal"
                                                                ' Retencion.RetCabecera._periodoFiscal = infoRt.InnerText.ToString
                                                                Retencion.PeriodoFiscal = infoRt.InnerText.ToString
                                                            Case "impuestos"
                                                        End Select
                                                    Next

                                                Case "impuestos"
                                                    Dim DetallesImpuestosRetencion = New List(Of Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion)

                                                    For Each impuestos As XmlNode In nodoInfTri.ChildNodes
                                                        Select Case impuestos.Name
                                                            Case "impuesto"
                                                                'Dim RTDetImp As New RetDetalleImpuestos
                                                                Dim DetImpuestoRetencion As New Entidades.wsEDoc_ConsultaRecepcion.ENTDetalleRetencion

                                                                For Each impuesto As XmlNode In impuestos.ChildNodes
                                                                    Select Case impuesto.Name
                                                                        Case "codigo"
                                                                            ' RTDetImp._codigo = CInt(impuesto.InnerText)
                                                                            DetImpuestoRetencion.Codigo = CInt(impuesto.InnerText)
                                                                        Case "codigoRetencion"
                                                                            ' RTDetImp._codigoRetencion = impuesto.InnerText
                                                                            DetImpuestoRetencion.CodigoRetencion = impuesto.InnerText
                                                                        Case "baseImponible"
                                                                            ' RTDetImp._baseImponible = impuesto.InnerText
                                                                            DetImpuestoRetencion.BaseImponible = impuesto.InnerText
                                                                        Case "porcentajeRetener"
                                                                            ' RTDetImp._porcentajeRetener = impuesto.InnerText
                                                                            DetImpuestoRetencion.PorcentajeRetener = impuesto.InnerText
                                                                        Case "valorRetenido"
                                                                            ' RTDetImp._valorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                                            DetImpuestoRetencion.ValorRetenido = validaSiEmpiezaPunto(impuesto.InnerText)
                                                                            Retencion.TotalRetencion = Retencion.TotalRetencion + DetImpuestoRetencion.ValorRetenido

                                                                        Case "codDocSustento"
                                                                            ' RTDetImp._codDocSustento = impuesto.InnerText
                                                                            DetImpuestoRetencion.CodDocRetener = impuesto.InnerText
                                                                        Case "numDocSustento"
                                                                            ' RTDetImp._numDocSustento = impuesto.InnerText
                                                                            DetImpuestoRetencion.NumDocRetener = impuesto.InnerText
                                                                        Case "fechaEmisionDocSustento"
                                                                            ' RTDetImp._fechaEmisionDocSustento = impuesto.InnerText.ToString
                                                                            DetImpuestoRetencion.FechaEmisionDocRetener = impuesto.InnerText
                                                                    End Select
                                                                Next
                                                                'Retencion.RetDetalleImp.Add(RTDetImp)
                                                                DetallesImpuestosRetencion.Add(DetImpuestoRetencion)

                                                        End Select
                                                    Next
                                                    'si existen impuestos lleno el detalle

                                                    If DetallesImpuestosRetencion.Count > 0 Then

                                                        Retencion.ENTDetalleRetencion = DetallesImpuestosRetencion.ToArray

                                                    End If

                                            End Select
                                        Next
                                End Select
                            Next


                        End If
                        ' GuardaLog("Retencion clave de acceso: " + Retencion.RetCabecera._claveAcceso + " leido Correctamente" + " - nombre del archivo: " + Path.GetFileName(ruta))

                        Utilitario.Util_Log.Escribir_Log("Retencion clave de acceso: " + Retencion.AutorizacionSRI + " leido Correctamente!!", "OperacionesXML")


                        Return Retencion

                    End If




                End If


            Catch ex As Exception

                Utilitario.Util_Log.Escribir_Log("Error al leer XML retencion: " + ex.Message.ToString + " con nombre: ", " OperacionesXML")


                ' GuardaLog("Error al leer XML retencion: " + ex.Message.ToString + " con nombre: " + Path.GetFileName(ruta))
                Return Nothing
            End Try

        End Function


    End Class


End Namespace


