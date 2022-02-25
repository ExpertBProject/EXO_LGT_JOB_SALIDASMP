Public Module JOB
#Region "Método principal"
    Public Sub Main()
        Dim iCountExeJOB As Integer = 0
        Dim sIDMax As String = "0"
        Dim oLog As EXO_Log.EXO_Log = Nothing
        Dim sError As String
        Dim sPath As String = ""
        Dim oFiles() As String = Nothing
        'Comprobamos si el JOB está en ejecución y en caso afirmativo no lanzamos ningún proceso del JOB.
        For Each oProcess As Process In Process.GetProcesses()
            If Left(oProcess.ProcessName.ToString, 21) = "EXO_LGT_JOB_SALIDASMP" Then
                iCountExeJOB += 1
            End If
        Next
        Try
            'sPath = My.Application.Info.DirectoryPath.ToString
            sPath = Conexiones.Datos_Confi("RUTAS", "LOG")

            If Not System.IO.Directory.Exists(sPath & "\SALIDASMP") Then
                System.IO.Directory.CreateDirectory(sPath & "\SALIDASMP")
            End If
            oLog = New EXO_Log.EXO_Log(sPath & "\SALIDASMP\EXO_LOG_INTERSALIDAS_", 10, EXO_Log.EXO_Log.Nivel.todos, 4, "", EXO_Log.EXO_Log.GestionFichero.dia)
            oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("###################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("###################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#####         INICIO PROCESO SALIDAS MP       #####", EXO_Log.EXO_Log.Tipo.informacion)

            'en cliente descomentamos
            If iCountExeJOB = 1 Then
                oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("Procedimiento 1. Limpieza de Hcos", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.LimpiarHcos(oLog)

                oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("Procedimiento 2. Salidas de Stock.", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.AjusteStockSalida(oLog)
            Else
                sError = "iCountExeJOB<>1"
                oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.advertencia)
            End If
        Catch ex As Exception
            If ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message <> "" Then
                sError = ex.InnerException.Message
            Else
                sError = ex.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            oLog.escribeMensaje("#####          FIN PROCESO SALIDAS MP         #####", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("###################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("###################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
        End Try
    End Sub
#End Region
End Module
