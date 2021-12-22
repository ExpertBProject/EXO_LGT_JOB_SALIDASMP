Imports System.IO
Imports System.Xml

Public Class Procesos
    Public Shared Sub LimpiarHcos(ByRef oLog As EXO_Log.EXO_Log)
        Try
            Dim sPath As String = Conexiones.Datos_Confi("GUARDAR", "RUTA")
            If Right(sPath, 1) <> "\" Then
                sPath &= "\HISTORIC\"
            Else
                sPath &= "HISTORIC\"
            End If

            If Not System.IO.Directory.Exists(sPath) Then
                System.IO.Directory.CreateDirectory(sPath)
            End If

            'Borramos los ficheros más antiguos a X días
            Dim Fecha As DateTime = DateTime.Now
            Dim sDias = Conexiones.Datos_Confi("GUARDAR", "DIASGUARDAR")
            For Each archivo As String In My.Computer.FileSystem.GetFiles(sPath, FileIO.SearchOption.SearchTopLevelOnly)
                Dim Fecha_Archivo As DateTime = My.Computer.FileSystem.GetFileInfo(archivo).LastWriteTime
                Dim diferencia = (CType(Fecha, DateTime) - CType(Fecha_Archivo, DateTime)).TotalDays

                If diferencia >= CDbl(sDias) Then ' Nº de días
                    File.Delete(archivo)
                    oLog.escribeMensaje("la diferencia de días es " & diferencia.ToString & ". Se borra el fichero " & archivo, EXO_Log.EXO_Log.Tipo.advertencia)
                End If
            Next
        Catch exCOM As System.Runtime.InteropServices.COMException
            Dim sError As String = "Limpiar Hcos - " & exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            Dim sError As String = "Limpiar Hcos - " & ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            oLog.escribeMensaje("Fin del proceso.", EXO_Log.EXO_Log.Tipo.advertencia)
        End Try
    End Sub
    Public Shared Sub AjusteStockSalida(ByRef oLog As EXO_Log.EXO_Log)
#Region "Variables"
        Dim sSQL As String = ""
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sDelimitador As String = "2"
        Dim sNomFichero As String = ""
        Dim sFecha As String = ""
        Dim oOIGE As SAPbobsCOM.Documents = Nothing
        Dim sArticulo As String = ""
        Dim sUnidades As String = ""
        Dim sImporte As String = ""
        Dim sBatchNumber As String = ""
        Dim bGestionLotes As Boolean = False
        Dim sWhsCode As String = ""
        Dim sAcctCodeEx As String = ""
        Dim sCeCo As String = ""
        Dim sDocEntryOIGE As String = "" : Dim sDocNumOIGE As String = ""
#End Region
        Try

            Dim sPath As String = Conexiones.Datos_Confi("GUARDAR", "RUTA")
            If Not System.IO.Directory.Exists(sPath) Then
                System.IO.Directory.CreateDirectory(sPath)
            End If
            Conexiones.Connect_Company(oCompany, "DI", oLog)
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim bEncuentraFich As Boolean = False
            For Each archivo As String In My.Computer.FileSystem.GetFiles(sPath, FileIO.SearchOption.SearchTopLevelOnly)
                sNomFichero = IO.Path.GetFileName(archivo)
                If Left(sNomFichero.ToUpper, 9) = "SALIDASMP" Then 'Ajuste Salida de stock
                    If File.Exists(archivo) Then
                        oLog.escribeMensaje("Tratando Fichero - " & archivo & " - ", EXO_Log.EXO_Log.Tipo.informacion)
                        bEncuentraFich = True
                        If oCompany.InTransaction = True Then
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        oCompany.StartTransaction()

                        oOIGE = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit), SAPbobsCOM.Documents)
                        Using MyReader As New Microsoft.VisualBasic.
                    FileIO.TextFieldParser(archivo, System.Text.Encoding.UTF7)
                            MyReader.TextFieldType = FileIO.FieldType.Delimited
                            Select Case sDelimitador
                                Case "1" : MyReader.SetDelimiters(vbTab)
                                Case "2" : MyReader.SetDelimiters(";")
                                Case "3" : MyReader.SetDelimiters(",")
                                Case "4" : MyReader.SetDelimiters("-")
                                Case Else : MyReader.SetDelimiters(vbTab)
                            End Select

                            Dim currentRow As String()
                            Dim bPrimeraLinea As Boolean = True

                            While Not MyReader.EndOfData
                                Try
                                    If bPrimeraLinea = True Then
                                        currentRow = MyReader.ReadFields() : currentRow = MyReader.ReadFields()
                                        bPrimeraLinea = False
                                        'Creo cabecera
                                        sFecha = Right(Left(archivo, Len(archivo) - 4), 8)
                                        oOIGE.DocDate = CDate(Right(sFecha, 2) & "/" & Mid(sFecha, 5, 2) & "/" & Left(sFecha, 4))
                                        oOIGE.TaxDate = CDate(Right(sFecha, 2) & "/" & Mid(sFecha, 5, 2) & "/" & Left(sFecha, 4))
                                        oOIGE.Comments = "Ajuste de stock salida recibido - " & archivo
                                        oOIGE.UserFields.Fields.Item("U_EXO_FILEINTERFACE").Value = archivo
                                    Else
                                        currentRow = MyReader.ReadFields()
                                        oOIGE.Lines.Add()
                                    End If

                                    Dim currentField As String
                                    Dim scampos(1) As String
                                    Dim iCampo As Integer = 0
                                    For Each currentField In currentRow
                                        iCampo += 1
                                        ReDim Preserve scampos(iCampo)
                                        scampos(iCampo) = currentField
                                        'SboApp.MessageBox(scampos(iCampo))
                                    Next
                                    sArticulo = scampos(1)
                                    sUnidades = scampos(2)
                                    sImporte = scampos(3)
                                    Try
                                        sBatchNumber = scampos(4)
                                    Catch ex As Exception
                                        sBatchNumber = ""
                                        oLog.escribeMensaje("AjusteStockSalida - Artículo " & sArticulo & " No tiene indicado Lote.", EXO_Log.EXO_Log.Tipo.advertencia)
                                    End Try
                                    'Comprobaciones por artículo
                                    ComprobarDatosEntradasInventarioMachos(oCompany, sArticulo)

                                    'Datos generales del artículo
                                    bGestionLotes = GestionLotes(oCompany, sArticulo)
                                    sWhsCode = AlmacenPorDefecto(oCompany, sArticulo)
                                    sSQL = "SELECT COALESCE(t1.""U_EXO_ACCTCODEEX"", '') ""U_EXO_ACCTCODEEX"", " &
                                        "COALESCE(t1.""U_EXO_OCRCODE"", '') ""U_EXO_OCRCODE"" " &
                                        "FROM ""OITB"" t1 INNER JOIN " &
                                        """OITM"" t2 ON t1.""ItmsGrpCod"" = t2.""ItmsGrpCod"" " &
                                        "WHERE t2.""ItemCode"" = '" & sArticulo & "'"
                                    oLog.escribeMensaje("AjusteStockSalida - SQL " & sSQL, EXO_Log.EXO_Log.Tipo.advertencia)
                                    oRs.DoQuery(sSQL)

                                    If oRs.RecordCount > 0 Then
                                        sAcctCodeEx = oRs.Fields.Item("U_EXO_ACCTCODEEX").Value.ToString
                                        sCeCo = oRs.Fields.Item("U_EXO_OCRCODE").Value.ToString
                                    Else
                                        sAcctCodeEx = ""
                                        sCeCo = ""
                                    End If

                                    oOIGE.Lines.ItemCode = sArticulo
                                    oOIGE.Lines.WarehouseCode = sWhsCode
                                    oOIGE.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, sUnidades)
                                    oOIGE.Lines.UnitPrice = EXO_GLOBALES.DblTextToNumber(oCompany, sImporte)
                                    If sCeCo.Trim <> "" Then
                                        oOIGE.Lines.CostingCode = sCeCo
                                    End If

                                    If sAcctCodeEx.Trim <> "" Then
                                        oOIGE.Lines.AccountCode = sAcctCodeEx
                                    End If

                                    If bGestionLotes = True And sBatchNumber <> "" Then
                                        oOIGE.Lines.BatchNumbers.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, sUnidades)
                                        oOIGE.Lines.BatchNumbers.BatchNumber = sBatchNumber
                                        oOIGE.Lines.BatchNumbers.Add()
                                    End If

                                Catch ex As Microsoft.VisualBasic.
                                    FileIO.MalformedLineException
                                    oLog.escribeMensaje("AjusteStockSalida - Línea " & ex.Message & " no es válida y se omitirá.", EXO_Log.EXO_Log.Tipo.error)
                                End Try
                            End While
                        End Using
                        'Cerramos el acceso al fichero para luego poder moverlo a la ruta de histórico
                        If Reader IsNot Nothing Then
                            Reader.Close() : Reader.Dispose() : Reader = Nothing
                        End If


                        If oOIGE.Add() <> 0 Then
                            oLog.escribeMensaje("AjusteStockSalida - " & oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", ""), EXO_Log.EXO_Log.Tipo.error)
                        Else
                            sDocEntryOIGE = oCompany.GetNewObjectKey
                            If oCompany.InTransaction = True Then
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            End If
                            oRs.DoQuery("SELECT ""DocNum"" FROM ""OIGE"" WHERE ""DocEntry"" =" & sDocEntryOIGE)
                            If oRs.RecordCount > 0 Then
                                sDocNumOIGE = oRs.Fields.Item("DocNum").Value.ToString
                                oLog.escribeMensaje("AjusteStockSalida - Se ha generado la Salida Nº " & sDocNumOIGE, EXO_Log.EXO_Log.Tipo.informacion)
                            Else
                                sDocNumOIGE = ""
                                oLog.escribeMensaje("AjusteStockSalida - No se ha podido generar la Salida del fichero" & archivo, EXO_Log.EXO_Log.Tipo.error)
                            End If
                            'Copiamos al Hco
                            My.Computer.FileSystem.MoveFile(archivo, sPath & "\HISTORIC\" & sNomFichero)
                            oLog.escribeMensaje("AjusteStockSalida - Se ha guardado el fichero al hco " & sPath & "\HISTORIC\" & sNomFichero, EXO_Log.EXO_Log.Tipo.informacion)
                        End If
                    Else
                        oLog.escribeMensaje("AjusteStockSalida - No se ha encontrado el fichero a cargar.", EXO_Log.EXO_Log.Tipo.error)
                        Exit Sub
                    End If
                End If
            Next
            If bEncuentraFich = False Then
                oLog.escribeMensaje("No se encuentra ficheros de ajuste Stock de salida.", EXO_Log.EXO_Log.Tipo.informacion)
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Dim sError As String = "AjusteStockSalida - " & exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            oLog.escribeMensaje("AjusteStockSalida - " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            myStream = Nothing
            If Reader IsNot Nothing Then
                Reader.Close() : Reader.Dispose() : Reader = Nothing
            End If

            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)
            oRs = Nothing
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Conexiones.Disconnect_Company(oCompany)
        End Try
    End Sub
    Private Shared Sub ComprobarDatosEntradasInventarioMachos(ByRef oCompany As SAPbobsCOM.Company, ByVal sItemCode As String)
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            'Artículo
            oRs.DoQuery("SELECT ""ItemCode"" " &
                        "FROM ""OITM"" " &
                        "WHERE ""ItemCode"" = '" & sItemCode & "'")

            If oRs.RecordCount = 0 Then
                Throw New Exception("El artículo " & sItemCode & " no está dado de alta en SAP.")
            End If

            'Cuenta de existencias y CeCo
            oRs.DoQuery("SELECT COALESCE(t1.""U_EXO_ACCTCODEEX"", '') ""U_EXO_ACCTCODEEX"", COALESCE(t1.""U_EXO_OCRCODE"", '') ""U_EXO_OCRCODE"", t1.""ItmsGrpNam"" " &
                        "FROM ""OITB"" t1 INNER JOIN " &
                        """OITM"" t2 ON t1.""ItmsGrpCod"" = t2.""ItmsGrpCod"" " &
                        "WHERE t2.""ItemCode"" = '" & sItemCode & "'")

            If oRs.RecordCount > 0 Then
                If oRs.Fields.Item("U_EXO_ACCTCODEEX").Value.ToString = "" Then
                    Throw New Exception("El grupo de artículos " & oRs.Fields.Item("ItmsGrpNam").Value.ToString & " no tiene definida la cuenta de existencias en SAP.")
                ElseIf oRs.Fields.Item("U_EXO_OCRCODE").Value.ToString = "" Then
                    Throw New Exception("El grupo de artículos " & oRs.Fields.Item("ItmsGrpNam").Value.ToString & " no tiene definido el CeCo en SAP.")
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)
            oRs = Nothing
        End Try
    End Sub
    Private Shared Function GestionLotes(ByRef oCompany As SAPbobsCOM.Company, ByVal sItemCode As String) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        GestionLotes = False

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT ""ItemCode"" " &
                        "FROM ""OITM"" " &
                        "WHERE ""ItemCode"" = '" & sItemCode & "' " &
                        "AND ""ManBtchNum"" = 'Y'")

            If oRs.RecordCount > 0 Then
                GestionLotes = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)
            oRs = Nothing
        End Try
    End Function
    Private Shared Function AlmacenPorDefecto(ByRef oCompany As SAPbobsCOM.Company, ByVal sItemCode As String) As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sWhsCode As String = ""

        AlmacenPorDefecto = ""

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT ""DfltWH"" " &
                        "FROM ""OITM"" " &
                        "WHERE ""ItemCode"" = '" & sItemCode & "'")

            If oRs.RecordCount > 0 Then
                'Almacén por defecto del artículo
                sWhsCode = oRs.Fields.Item("DfltWH").Value.ToString
            End If

            If sWhsCode = "" Then
                oRs.DoQuery("SELECT ""DfltWhs"" " &
                            "FROM ""OADM"" " &
                            "WHERE ""Code"" = 1")

                If oRs.RecordCount > 0 Then
                    'Almacén por defecto parametrizaciones generales
                    sWhsCode = oRs.Fields.Item("DfltWhs").Value.ToString
                End If
            End If

            AlmacenPorDefecto = sWhsCode

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)
            oRs = Nothing
        End Try
    End Function
End Class
