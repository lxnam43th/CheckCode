Module Module1

    Sub Main()
        Dim oQCEmail As New QCEmail
        Dim mConnectionString As String
        Dim mConnectionString_stb As String
        Dim mConnectionstring_gdis As String
        Dim m_pddate As String = "50"

        Dim bolTrace As Boolean
        Try
            'Get argument from the parameter
            Dim arrArgs() As String = Command.Split(",")
            Dim i As Integer
            Dim strItemListFromArgs As String = String.Empty

            If Not arrArgs(0) Is Nothing Then
                For i = LBound(arrArgs) To UBound(arrArgs)
                    'Console.Write("Parameter " & i & " is " & arrArgs(i) & vbNewLine)
                    If strItemListFromArgs = String.Empty Then
                        strItemListFromArgs = arrArgs(i).Replace("'", "''")
                    Else
                        strItemListFromArgs &= "," & arrArgs(i).Replace("'", "''")
                    End If
                Next
            Else
                'Console.Write("No parameter passed")
                strItemListFromArgs = String.Empty
            End If

            'Console.Write(strItemListFromArgs)
            'Exit Sub

            'Get some setting from config file
            mConnectionString = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")

            mConnectionString_stb = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString_stb")
            mConnectionstring_gdis = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString_gdis")

            'Transfer connection to the DLL
            oQCEmail.sConnectionString = mConnectionString
            oQCEmail.sConnectionString_stb = mConnectionString_stb
            oQCEmail.ConnectionString_gdis = mConnectionstring_gdis

            If System.Configuration.ConfigurationSettings.AppSettings("Trace") = "True" Then
                bolTrace = True
            Else
                bolTrace = False
            End If

            oQCEmail.TRACE_ENABLE = bolTrace

            'Define the listener file
            Dim strPath_Log_File As String = System.AppDomain.CurrentDomain.BaseDirectory() & "QCDailyInformation.text"
            Dim fsTextWriter As System.IO.TextWriter = New System.IO.StreamWriter(strPath_Log_File, False)

            'Dim myQCListenerFile As TextWriterTraceListener = New TextWriterTraceListener(strPath_Log_File)
            Dim myQCListenerFile As TextWriterTraceListener = New TextWriterTraceListener(fsTextWriter)
            'c:\QCEmailLog.txt
            Trace.Listeners.Add(myQCListenerFile)
            Trace.AutoFlush = True
            Trace.WriteLineIf(bolTrace, "Start at " & Now)

            Dim strImageFolder As String = System.AppDomain.CurrentDomain.BaseDirectory() & "ExportImage"
            'At First we should delete all file from generated image before
            oQCEmail.DeleteAllFiles(strImageFolder)

            'Get Item list 

            Dim strItemList As String = "" '= System.Configuration.ConfigurationSettings.AppSettings("ItemList")
            Dim strGroupItemList As String = ""
            'CAL Service position from app config file
            Dim strHostname As String
            Dim servicepos As Integer
            Dim tablename As String

            strHostname = System.Net.Dns.GetHostName()

            servicepos = System.Configuration.ConfigurationSettings.AppSettings("ServicePos")
            Trace.WriteLineIf(bolTrace, " Position: " & servicepos)
            tablename = System.Configuration.ConfigurationSettings.AppSettings("TableName")
            If strItemListFromArgs = String.Empty Then

                Try
                    oQCEmail.GetItemlist(mConnectionString_stb, servicepos, strItemList, strGroupItemList)
                Catch ex As Exception
                    Trace.WriteLineIf(bolTrace, "error on running get list item" & ex.InnerException.Message)
                End Try
            Else
                'get item from parameter
                strItemList = strItemListFromArgs
            End If

            'Trace.WriteLineIf(bolTrace, "List of Item to Generate : " & strItemList & " at " & Now.ToString())

            Trace.WriteLineIf(bolTrace, "List of GrpItem to Generate : " & strGroupItemList & " at " & Now.ToString())
            'Console.Write(strItemList)
            'Exit Sub
            Dim strImageFolder_Delete As String = System.AppDomain.CurrentDomain.BaseDirectory() & "ExportImage\GrpImage"
            'At First we should delete all file from generated image before

            Trace.WriteLineIf(bolTrace, "Delete old images: ")
            oQCEmail.DeleteAllFiles(strImageFolder_Delete)
            'First insert into database 1 record to confirm this service will run
            'oQCEmail.SaveDB(mConnectionString)

            'Init the service
            oQCEmail.InitControl()
            'strItemList = "GN6A"
            'Run to generate for all item
            oQCEmail.ServicePos = servicepos
            ' strGroupItemList = "GN6A"
            ' insert data
            oQCEmail.GenerateItem_GroupItem(strItemList, 0, tablename)
            oQCEmail.GenerateItem_GroupItem(strGroupItemList, 1, tablename)

            'Run to generate for All Group Item
            'oQCEmail.GenerateItem_GroupItem(strGroupItemList, 1)


        Catch ex As Exception
            Trace.WriteLineIf(bolTrace, "Error on GenerateItem " & ex.ToString)
            'oQCEmail.SendEmailOnError("Generate QC Daily Information", ex.ToString())
        Finally
            'End of service , close all trace listener 
            Trace.WriteLineIf(bolTrace, "Stop at " & Now)
            For i As Integer = 0 To Trace.Listeners.Count - 1
                Trace.Listeners(i).Close()
            Next
            End
        End Try
    End Sub

End Module
