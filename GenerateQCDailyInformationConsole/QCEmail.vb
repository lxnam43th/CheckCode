Imports System.Math
Imports System.Data
Imports System.Data.OleDb
Imports ChartFX.WinForms
Imports System.String
Imports System.Drawing
Imports System.IO.Path
Imports System.Drawing.Imaging
Imports System.IO
Imports System.IO.File
Imports System.Net.Mail


Public Class QCEmail
    Public sConnectionString As String = "Provider=OraOLEDB.Oracle.1;User ID=qcinput;Data Source=hoyav3;Password=qcinputfactorypwd"
    Public sConnectionString_stb As String = "Provider=OraOLEDB.Oracle.1;User ID=qcinput;Data Source=HYN_bk ;Password=qcinputfactorypwd"
    Public sConnectionString_gdis As String = "Provider=OraOLEDB.Oracle.1;User ID=gdis;Data Source=gdisweb ;Password=gdis"
    Dim SpectNothing As Integer = 999999
    Dim get_pddate As String = "50"

    Dim rTarget, rLSL, rUSL, rLCL, rUCL, rARUL, rARLL As Double
    Dim Isnumeric As Integer
    Dim arrParam As New ArrayList
    Dim dtData As New DataTable
    Dim dsQC As New DataSet
    Dim adapter As OleDbDataAdapter = New OleDbDataAdapter
    Dim ChartTrend As New ChartFX.WinForms.Chart
    'Dim sShippingItem As String
    'Dim ShippingItemOID As String
    'Dim sGroupItem As String
    'Dim sGroupItemOID As String
    Public TRACE_ENABLE As Boolean
    Dim iServicePos As Integer
    Dim sFullPath As String = "\\172.25.9.61\ExportImage"
    Public Property ServicePos() As Integer
        Get
            Return iServicePos

        End Get
        Set(ByVal value As Integer)
            iServicePos = value
        End Set
    End Property
    Public Property FullPath() As String
        Get
            Return sFullPath
        End Get
        Set(ByVal value As String)
            sFullPath = value
        End Set
    End Property
    Public Property ConnectionString() As String
        Get
            Return sConnectionString
        End Get
        Set(ByVal Value As String)
            sConnectionString = Value
        End Set
    End Property
    Public Property ConnectionString_stb() As String
        Get
            Return sConnectionString_stb
        End Get
        Set(ByVal Value As String)
            sConnectionString_stb = Value
        End Set
    End Property
    Public Property ConnectionString_gdis() As String
        Get
            Return sConnectionString_gdis
        End Get
        Set(ByVal Value As String)
            sConnectionString_gdis = Value
        End Set
    End Property
    Public Sub InitControl()
        Try
            Trace.WriteLineIf(TRACE_ENABLE, "Start QCEmail.InitControl " & Now)
            ChartTrend.Width = 600
            ChartTrend.Height = 200
            With Me.ChartTrend
                ' .BackColor = Color.WhiteSmoke
                .AxisY.Grids.Interlaced = True

                .AxisY.Grids.InterlacedColor = Color.Beige

            End With
            LoadShippingItem() 'On Hoyav3_stb
            LoadGroupItem()

            Trace.WriteLineIf(TRACE_ENABLE, "End QCEmail.InitControl " & Now)
        Catch ex As Exception
            Trace.WriteLineIf(TRACE_ENABLE, " ===============Error of QCEmail.InitControl : " & Now)
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub
    ' option is 0 : Generate Item, 1 : Generate GroupItem
    Public Sub GenerateItem_GroupItem(ByVal Str_Lst_Item As String, ByVal options As Integer, tableName As String)

        Dim sShippingItem As String
        Dim ShippingItemOID As String
        Dim sGroupItem As String
        Dim sGroupItemOID As String
        Try
            Dim tempItem() As String = Str_Lst_Item.Split(",")
            Dim str As String
            If options = 0 Then
                For Each str In tempItem
                    'If str = "WD95-6" Then
                    sShippingItem = str.Trim
                    ShippingItemOID = dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID") 'On Hoyav3_stb
                    'If CheckExistData(ShippingItemOID) Then 'On Hoyav3_stb
                    Trace.WriteLineIf(TRACE_ENABLE, " Start Generate for Item :" & sShippingItem & " AT :" & Now)
                    DrawingForEachSPItem(sShippingItem, ShippingItemOID, Today.ToString("yyyyMMdd"), options, tableName)
                    Trace.WriteLineIf(TRACE_ENABLE, " End Generate for Item :" & sShippingItem & " AT :" & Now)
                    Trace.WriteLineIf(TRACE_ENABLE, GC.GetTotalMemory(True))
                Next
            Else
                For Each str In tempItem
                    sGroupItem = str.Trim
                    sGroupItemOID = dsQC.Tables("GroupItem").Rows.Find(sGroupItem).Item("OID").ToString 'On Hoyav3_stb
                    'If CheckExistData(sGroupItemOID) Then 'On Hoyav3_stb
                    Trace.WriteLineIf(TRACE_ENABLE, " Start Generate for Group Item :" & sGroupItem & " AT :" & Now)
                    DrawingForEachSPItem(sGroupItem, sGroupItemOID, Today.ToString("yyyyMMdd"), options, tableName)
                    Trace.WriteLineIf(TRACE_ENABLE, " End Generate for Item :" & sGroupItem & " AT :" & Now)
                    ' End If
                Next
            End If

        Catch ex As Exception
            Trace.WriteLineIf(TRACE_ENABLE, " =====Error : QCMail.GenerateItem :" & Now)
            Trace.WriteLineIf(TRACE_ENABLE, ex.ToString())
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub
    'Public Sub GenerateImageHAL()
    '    'DrawingAllShippingItem()
    '    sShippingItem = "HALONG"
    '    ShippingItemOID = dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID")
    '    If CheckExistData(ShippingItemOID) Then
    '        DrawingForEachSPItem("HALONG", Today.ToString("yyyyMMdd"))
    '    End If
    'End Sub
    'Public Sub GenerateImageNHA()
    '    sShippingItem = "NHA"
    '    ShippingItemOID = dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID")
    '    If CheckExistData(ShippingItemOID) Then
    '        DrawingForEachSPItem("NHA", Today.ToString("yyyyMMdd"))
    '    End If
    'End Sub
    'Public Sub GenerateImageSAP()
    '    sShippingItem = "SAP"
    '    ShippingItemOID = dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID")
    '    If CheckExistData(ShippingItemOID) Then
    '        DrawingForEachSPItem("SAP", Today.ToString("yyyyMMdd"))
    '    End If
    'End Sub
    'Public Sub GenerateImageVENUS()
    '    sShippingItem = "VENUS"
    '    ShippingItemOID = dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID")
    '    If CheckExistData(ShippingItemOID) Then
    '        DrawingForEachSPItem("VENUS", Today.ToString("yyyyMMdd"))
    '    End If
    'End Sub
    'Public Sub GenerateImageMN4()
    '    sShippingItem = "MN4"
    '    ShippingItemOID = dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID")
    '    If CheckExistData(ShippingItemOID) Then
    '        DrawingForEachSPItem("MN4", Today.ToString("yyyyMMdd"))
    '    End If

    'End Sub
    'Public Sub GenerateImageSRB()
    '    sShippingItem = "SRB"
    '    ShippingItemOID = dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID")
    '    If CheckExistData(ShippingItemOID) Then
    '        DrawingForEachSPItem("SRB", Today.ToString("yyyyMMdd"))
    '    End If
    'End Sub
    'Public Sub GenerateImageCRS()
    '    sShippingItem = "SRB"
    '    ShippingItemOID = dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID")
    '    If CheckExistData(ShippingItemOID) Then
    '        DrawingForEachSPItem("SRB", Today.ToString("yyyyMMdd"))
    '    End If
    'End Sub

    Public Sub GetItemlist(ByVal mConnectionString As String, ByVal servicepos As Integer, ByRef listitem As String, ByRef listGroupitem As String)

        'Dim mConnectionString As String
        Dim moldDBCon As OleDbConnection

        Dim i As Integer
        Try


            'mConnectionString = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
            moldDBCon = New OleDbConnection(mConnectionString)
            moldDBCon.Open()

            Dim tempDate As String
            tempDate = Now.ToString("yyyyMMdd")

            Dim cmd As OleDbCommand
            Dim strSQL As String
            strSQL = "select distinct ITEMNAME from mqcidscanspec" & _
            " where statusflag = 1 "
            If servicepos = 4 Then
                strSQL &= " AND DISPLAYONDAILY = 1 " & _
                            " and ITEMOID in (select SPITEMOID from mshippingitem_generated where last_pddate >= to_number(to_char(sysdate - 7,'YYYYMMDD'))) "
            Else
                strSQL &= " AND DISPLAYONDAILY = 1 AND SERVICEPOS = " & servicepos & _
                            " and ITEMOID in (select SPITEMOID from mshippingitem_generated where last_pddate >= to_number(to_char(sysdate - 7,'YYYYMMDD'))) "
            End If

            'namlx.
            ' lam moi. 27-02-2017
            strSQL = " select distinct SPITEMOID as ITEMNAME from mshippingitem_generated where last_pddate >= to_number(to_char(sysdate - " + get_pddate + ",'YYYYMMDD')) "


            ''            strSQL = "select itemname from mqcidscanspec " & _
            ''" where statusflag=1" & _
            ''" and DISPLAYONDAILY = 1" & _
            ''" and itemname not in (select distinct shippingitemname as ITEMNAME from qctrendimage" & _
            ''" where qcdate = 20100930)"
            cmd = New OleDbCommand(strSQL, moldDBCon)
            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim ds As DataSet = New DataSet("ITEMLIST")
            adapter.Fill(ds, "ITEMLIST")
            For i = 0 To ds.Tables("Itemlist").Rows.Count - 1
                If i = 0 Then
                    listitem = ds.Tables("Itemlist").Rows(i).Item("ItemName")
                    '  listGroupitem = ds.Tables("Itemlist").Rows(i).Item("ITEMGROUP_NAME")
                Else
                    listitem = listitem & "," & ds.Tables("Itemlist").Rows(i).Item("ItemName")
                    ' listGroupitem = listGroupitem & "," & ds.Tables("Itemlist").Rows(i).Item("ITEMGROUP_NAME")
                End If
            Next

            '-------Group Item 

            Dim strSqlGrp As String

            strSqlGrp = "SELECT distinct ITEMGROUP_NAME FROM MQCIDSCANSPEC WHERE STATUSFLAG =1 AND DISPLAYONDAILY = 1  " & _
            " and ITEMGROUP_NAME  in (select distinct NAME" & _
             " from mshippingitem_grp t" & _
             " where t.statusflag = 1) and ITEMOID in (select SPITEMOID from mshippingitem_generated where last_pddate >= to_number(to_char(sysdate - 7,'YYYYMMDD'))) order by ITEMGROUP_NAME"

            strSqlGrp = "SELECT distinct name as ITEMGROUP_NAME FROM mshippingitem_grp  WHERE oid in ( "
            strSqlGrp += " select distinct SPITEMOID from mshippingitem_generated where last_pddate >= to_number(to_char(sysdate - " + get_pddate + ",'YYYYMMDD')) "
            strSqlGrp += " ) "

            cmd = New OleDbCommand(strSqlGrp, moldDBCon)
            Dim adapter1 As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim ds1 As DataSet = New DataSet("GrpItemList")
            adapter1.Fill(ds1, "GrpItemList")
            For i = 0 To ds1.Tables("GrpItemList").Rows.Count - 1
                If i = 0 Then
                    listGroupitem = ds1.Tables("GrpItemList").Rows(i).Item("ITEMGROUP_NAME")
                Else
                    listGroupitem = listGroupitem & "," & ds1.Tables("GrpItemList").Rows(i).Item("ITEMGROUP_NAME")
                End If
            Next
            moldDBCon.Close()

        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub
    Public Sub InsParamNoData(mConnectionString As String, itemname As String, operationoid As String, operationname As String, parametername As String)
        Dim str As String
        'Dim mConnectionString As String
        Dim moldDBCon As OleDbConnection
        Try

            'mConnectionString = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
            moldDBCon = New OleDbConnection(mConnectionString)
            moldDBCon.Open()

            Dim tempDate As String
            tempDate = Now.ToString("yyyyMMdd")

            Dim cmd As New OleDbCommand()
            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim strSQL As String
            strSQL = "select * from qctrendimage_nodataparam where itemname ='" & itemname & "' and parametername = '" & parametername & "'"
            cmd.CommandText = strSQL
            cmd.Connection = moldDBCon
            cmd.CommandType = CommandType.Text
            Dim ds As DataSet = New DataSet("qctrendimage_nodataparam")
            adapter.Fill(ds, "qctrendimage_nodataparam")

            If ds.Tables("qctrendimage_nodataparam").Rows.Count > 0 Then
                'update qcdate, lastupdate (keep startdate)
                strSQL = "update qctrendimage_nodataparam set qcdate = " & tempDate & ", lastupdate = sysdate"
                cmd.ExecuteNonQuery()
            Else
                'insert new to table
                strSQL = "insert into qctrendimage_nodataparam (itemname, operationoid, operationname, parametername, qcdate,startdate, lastupdate) values('" & itemname & "','" & operationoid & "','" & operationname & "','" & parametername & "'," & tempDate & ", sysdate, sysdate" & ") "
                cmd.ExecuteNonQuery()
            End If

            moldDBCon.Close()
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub

    Public Sub SaveDB(ByVal mConnectionString As String)
        Dim str As String
        'Dim mConnectionString As String
        Dim moldDBCon As OleDbConnection
        Try

            'mConnectionString = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
            moldDBCon = New OleDbConnection(mConnectionString)
            moldDBCon.Open()

            Dim tempDate As String
            tempDate = Now.ToString("yyyyMMdd")

            Dim cmd As OleDbCommand
            Dim strSQL As String
            strSQL = "insert into QCEmail(sendDate,LastUpdate) values('" & tempDate & "', sysdate) "
            cmd = New OleDbCommand(strSQL, moldDBCon)
            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim ds As DataSet = New DataSet("SENT_LIST")
            adapter.Fill(ds, "SENT_LIST")

            moldDBCon.Close()
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub

    Public Sub SendEmailOnError(ByVal strSubject As String, ByVal strContent As String)
        'Dim objEmail As New MailMessage("thien_ha@sngw.els.hoya.co.jp", "thien_ha@sngw.els.hoya.co.jp", strSubject, strContent)
        'objEmail.Priority = MailPriority.High
        ''SmtpMail.SmtpServer  = "localhost"
        'Try

        '    Dim smtp As SmtpClient = New SmtpClient
        '    smtp.Host = "172.25.6.41" ' "172.25.10.22"
        '    smtp.Port = "25"
        '    smtp.Send(objEmail)
        '    'smtp.DeliveryMethod = SmtpDeliveryMethod.Network
        '    'smtp.Credentials = New Net.NetworkCredential("thien", "matkhaumail")

        '    '// Smtp configuration
        '    '       SmtpClient smtp = new SmtpClient();
        '    '       smtp.Host = "smtp.gmail.com";

        '    '       smtp.Credentials = new System.Net.NetworkCredential("xx", "xx");
        '    '       smtp.EnableSsl = true;   

        'Catch exc As Exception
        '    Throw New Exception(exc.ToString, exc)
        'End Try

    End Sub

    Public Function DeleteAllFiles(ByVal strFolder As String) As Boolean
        Try
            Dim S() As String
            If Directory.Exists(strFolder) Then
                S = Directory.GetFiles(strFolder)
                Dim DELFILE As String
                For Each DELFILE In S
                    File.Delete(DELFILE)
                Next
            End If
        Catch ex As Exception
        End Try
    End Function


#Region "Load master"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load all shipping item -- hien tai co HAL, NHA, SAP
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[dmquyen]	07/26/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadShippingItem()
        Try
            Dim strSQL As String
            Dim cmd As OleDbCommand
            Dim con As OleDbConnection

            Dim dsShippingItem As DataSet = New DataSet
            Dim strCon As String = sConnectionString_stb
            con = New OleDbConnection(strCon)
            con.Open()

            If IsNothing(dsQC.Tables("ShippingItem")) Then  'Loaded yet
                strSQL = "select distinct t.SPITEMOID as oid, t.SPITEMNAME as name" & _
                 " from mshippingitem_generated t" & _
                 " where t.LAST_PDDATE >= " & Today.AddDays(-Convert.ToInt16(get_pddate)).ToString("yyyyMMdd") & _
                 " order by t.SPITEMNAME"
                cmd = New OleDbCommand(strSQL, con)
                cmd.CommandType = CommandType.Text
                adapter = New OleDbDataAdapter(cmd)
                If Not IsNothing(dsQC.Tables("ShippingItem")) Then
                    dsQC.Tables("ShippingItem").Clear()
                End If
                adapter.Fill(dsQC, "ShippingItem")
                dsQC.Tables("ShippingItem").PrimaryKey = New DataColumn() {dsQC.Tables("ShippingItem").Columns("NAME")}
            End If
            con.Close()
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub

    Private Sub LoadGroupItem()
        Try
            Dim strSQL As String
            Dim cmd As OleDbCommand
            Dim con As OleDbConnection

            Dim dsShippingItem As DataSet = New DataSet
            Dim strCon As String = sConnectionString_stb
            con = New OleDbConnection(strCon)
            con.Open()
            '" inner join mreportparam t1 on t.oid = t1.shippingoid" & _
            If IsNothing(dsQC.Tables("GroupItem")) Then  'Loaded yet

                'strSQL = "select distinct t.oid, t.name" & _
                ' " from mshippingitem_grp t" & _
                ' " where t.statusflag = 1" & _
                ' " order by t.name"

                strSQL = "SELECT distinct ITEMGROUP_NAME, ITEMGROUP_OID as OID FROM MQCIDSCANSPEC WHERE STATUSFLAG =1 AND DISPLAYONDAILY = 1 " & _
            " and ITEMGROUP_NAME  in (select distinct NAME" & _
             " from mshippingitem_grp t" & _
             " where t.statusflag = 1) and ITEMOID in (select SPITEMOID from mshippingitem_generated a where a.last_pddate >= to_number(to_char(sysdate - 7,'YYYYMMDD'))) order by ITEMGROUP_NAME"

                strSQL = " SELECT distinct name as  ITEMGROUP_NAME,  OID FROM mshippingitem_grp WHERE STATUSFLAG =1 and oid in ( "
                strSQL += " select grp_oid from mshippingitem a where oid in ( "
                strSQL += " select SPITEMOID from mshippingitem_generated a where a.last_pddate >= to_number(to_char(sysdate - 7,'YYYYMMDD')) "
                strSQL += " ) "
                strSQL += " ) "
                'groupitemname, groupitemoid
                strSQL = " select distinct groupitemname as ITEMGROUP_NAME, groupitemoid as OID from mshippingitem_generated a where a.last_pddate >= to_number(to_char(sysdate - " + get_pddate + ",'YYYYMMDD')) "

                cmd = New OleDbCommand(strSQL, con)
                cmd.CommandType = CommandType.Text
                adapter = New OleDbDataAdapter(cmd)
                If Not IsNothing(dsQC.Tables("GroupItem")) Then
                    dsQC.Tables("GroupItem").Clear()
                End If
                adapter.Fill(dsQC, "GroupItem")
                dsQC.Tables("GroupItem").PrimaryKey = New DataColumn() {dsQC.Tables("GroupItem").Columns("ITEMGROUP_NAME")}
            End If
            con.Close()
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load section by shipping item
    ''' </summary>
    ''' <param name="strShippingOID"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[dmquyen]	7/1/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function LoadSection(ByVal strShippingOID As String, ByVal options As Integer) As Boolean
        Dim strSQL, strCon As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection
        Try


            strCon = sConnectionString_stb
            con = New OleDbConnection(strCon)
            con.Open()

            strSQL = "select distinct MO.oid, Mo.oname" & _
             " from moperation MO" & _
             " inner join mreportparam MR on MO.oid = MR.operationoid"
            If options = 0 Then
                strSQL &= " where MR.shippingoid = '" & strShippingOID & "'" &
                   " order by to_Number(MO.oid)"
            Else
                strSQL &= " where MR.shippingoid in (select OID from mshippingitem where grp_oid= '" & strShippingOID & "')" &
                        " order by to_Number(MO.oid)"
            End If


            cmd = New OleDbCommand(strSQL, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)
            If Not IsNothing(dsQC.Tables("OPERATION")) Then
                dsQC.Tables("OPERATION").Clear()
            End If
            'Dim dsOPERATION As DataSet = New DataSet
            adapter.Fill(dsQC, "OPERATION")
            dsQC.Tables("Operation").PrimaryKey = New DataColumn() {dsQC.Tables("Operation").Columns("OID")}
            con.Close()
            Return True
        Catch ex As Exception
            con.Close()
            Throw New Exception(ex.Message, ex)
        End Try
    End Function

    Private Function LoadParameter(ByVal strShippingOID As String, ByVal options As String) As Boolean
        Dim strSQL, strCon As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection

        Try


            strCon = sConnectionString_stb
            con = New OleDbConnection(strCon)
            con.Open()

            '----------------------------------------------------
            If options = 0 Then
                strSQL = "select RPTPARAMOID as oid,RPTNAME as reportparaname, facode, operationoid, operationname" &
                " from mshippingitem_param_generated " &
                " where SPITEMOID ='" & strShippingOID & "'" &
                " and last_pddate >= " & Today.AddDays(-Convert.ToInt16(get_pddate)).ToString("yyyyMMdd") '& _
                '"   and ISNUMERIC = 0"

            Else
                strSQL = "select distinct 'GRP' as OID, RPTNAME as reportparaname, facode, operationoid, operationname" &
                " from mshippingitem_param_generated " &
                " where GRPITEMOID ='" & strShippingOID & "'" &
                " and last_pddate >= " & Today.AddDays(-Convert.ToInt16(get_pddate)).ToString("yyyyMMdd")
            End If
            strSQL &= " Order by RPTNAME"

            'strSQL = "select oid, reportparaname,Facode from mreportparam " & _
            '" where statusflag = 1 and shippingoid= 'a9ecd06d-6778-4e5e-9ad3-d79a92e86edb' and operationoid = '410'" & _
            '" Order by reportparaname"

            cmd = New OleDbCommand(strSQL, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)

            'Dim dsQCPARAMETER As DataSet = New DataSet
            If Not IsNothing(dsQC.Tables("QCPARAMETER")) Then
                dsQC.Tables("QCPARAMETER").Clear()
            End If
            adapter.Fill(dsQC, "QCPARAMETER")
            ' ''If options = 0 Then
            ' ''    dsQC.Tables("QCPARAMETER").PrimaryKey = New DataColumn() {dsQC.Tables("QCPARAMETER").Columns("OID")}
            ' ''End If

            con.Close()
            Return True
        Catch ex As Exception
            con.Close()
            Throw New Exception(ex.Message, ex)
        End Try
    End Function

    Private Sub LoadInfoParameter(ByVal strQCParamOID As String, ByVal paramname As String, ByVal options As Integer, Optional ByVal strSpitemoid As String = "", Optional ByVal stroperationoid As String = "")
        Dim strSQL, strCon As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection
        Dim strTPDDate, strFPDDate As String
        Try

            strCon = sConnectionString_stb
            con = New OleDbConnection(strCon)
            con.Open()

            Dim dr As DataRow
            '' ''strTPDDate = Today.ToString("yyyyMMdd")
            '' ''strFPDDate = Today.AddDays(-7).ToString("yyyyMMdd")

            '' ''strSQL = "select distinct t1.oid, t1.name,replace(t1.nickname,'  ',' ') as nickname ,t4.CUL, t4.UL as USL, t4.SUL, t4.ARUL, " & _
            '' ''"t4.ARLL, t4.LL as LSL, t4.sll, t4.cll, t4.target, t4.loc, t4.uoc,t3.begindate, t3.beginshift " & _
            '' ''" from mqcstdparam t1 inner join mreport_qcstdparam t2 on t1.oid = t2.mqcstdparamoid" & _
            '' ''" inner join dqcparameterversion t3 on t1.oid = t3.qcstdparamoid " & _
            '' ''"left join dqcparameterspec t4 on t3.oid = t4.dqcparameterversionoid " & _
            '' ''"where t2.mreportparamoid = '" & strQCParamOID & "' " & _
            '' ''" and t3.spitemoid='" & strSpitemoid & "' " & _
            '' ''" and t3.operationoid = '" & stroperationoid & "' " & _
            '' ''" and t3.PDITEMOID ='00c90c93-32de-404d-a967-078c65f2a5bb' " & _
            '' ''" order by t3.begindate desc, t3.beginshift desc "

            ' '' ''strSQL = " select distinct t4.CUL, t4.UL as USL, t4.SUL, t4.ARUL, " & _
            ' '' ''"t4.ARLL, t4.LL as LSL, t4.sll, t4.cll, t4.target, t4.loc, t4.uoc,t3.begindate, t3.beginshift" & _
            ' '' ''"from dpd_qc_info t4 where t4.STDPARAMREPORTOID='" & strQCParamOID & "' and t4.SPITEMOID='" & strSpitemoid & "' and t4.operationoid ='" & stroperationoid & "'"


            ' ''cmd = New OleDbCommand(strSQL, con)
            ' ''cmd.CommandType = CommandType.Text
            ' ''adapter = New OleDbDataAdapter(cmd)

            '
            '' ''Dim dsQCPARAMETER As DataSet = New DataSet
            ' ''If Not IsNothing(dsQC.Tables("QCPARAMETERINFO")) Then
            ' ''    dsQC.Tables("QCPARAMETERINFO").Clear()
            ' ''End If

            ' ''adapter.Fill(dsQC, "QCPARAMETERINFO")

            ' ''For Each dr In dsQC.Tables("QCPARAMETERINFO").Rows
            ' ''    If CInt(strFPDDate) = CInt(dr.Item("Begindate")) Then
            ' ''        'If cboPDShift.SelectedItem.Text = "ALL" Then
            ' ''        If dr.Item("beginshift") = "DAY" Then
            ' ''            Exit For
            ' ''        End If
            ' ''        'Else
            ' ''        'If cboPDShift.SelectedItem.Text >= dr.Item("beginshift") Then
            ' ''        '    Exit For
            ' ''        'End If
            ' ''        'End If
            ' ''    ElseIf CInt(strFPDDate) > CInt(dr.Item("Begindate")) Then
            ' ''        Exit For
            ' ''    End If
            ' ''Next
            ' ''If Not dr Is Nothing Then
            ' ''    If Not IsDBNull(dr.Item("TARGET")) Then
            ' ''        rTarget = Mid(dr.Item("TARGET"), 2)
            ' ''    Else
            ' ''        rTarget = SpectNothing
            ' ''    End If
            ' ''    If Not IsDBNull(dr.Item("LSL")) Then
            ' ''        If Asc(dr.Item("LSL")(1)) <> 61 Then  'in case < or >
            ' ''            rLSL = Mid(dr.Item("LSL"), 2)
            ' ''        Else
            ' ''            rLSL = Mid(dr.Item("LSL"), 3)
            ' ''        End If
            ' ''    Else
            ' ''        rLSL = SpectNothing
            ' ''    End If
            ' ''    If Not IsDBNull(dr.Item("USL")) Then
            ' ''        If Asc(dr.Item("USL")(1)) <> 61 Then  'in case < or >
            ' ''            rUSL = Mid(dr.Item("USL"), 2)
            ' ''        Else
            ' ''            rUSL = Mid(dr.Item("USL"), 3)
            ' ''        End If
            ' ''    Else
            ' ''        rUSL = SpectNothing
            ' ''    End If
            ' ''    If Not IsDBNull(dr.Item("CLL")) Then
            ' ''        If Asc(dr.Item("CLL")(1)) <> 61 Then  'in case < or >
            ' ''            rLCL = Mid(dr.Item("CLL"), 2)
            ' ''        Else
            ' ''            rLCL = Mid(dr.Item("CLL"), 3)
            ' ''        End If
            ' ''    Else
            ' ''        rLCL = SpectNothing
            ' ''    End If
            ' ''    If Not IsDBNull(dr.Item("CUL")) Then
            ' ''        If Asc(dr.Item("CUL")(1)) <> 61 Then  'in case < or >
            ' ''            rUCL = Mid(dr.Item("CUL"), 2)
            ' ''        Else
            ' ''            rUCL = Mid(dr.Item("CUL"), 3)
            ' ''        End If
            ' ''    Else
            ' ''        rUCL = SpectNothing
            ' ''    End If
            ' ''    If Not IsDBNull(dr.Item("ARUL")) Then
            ' ''        If Asc(dr.Item("ARUL")(1)) <> 61 Then  'in case < or >
            ' ''            rARUL = Mid(dr.Item("ARUL"), 2)
            ' ''        Else
            ' ''            rARUL = Mid(dr.Item("ARUL"), 3)
            ' ''        End If
            ' ''    Else
            ' ''        rARUL = SpectNothing
            ' ''    End If
            ' ''    If Not IsDBNull(dr.Item("ARLL")) Then
            ' ''        If Asc(dr.Item("ARLL")(1)) <> 61 Then 'in case < or >
            ' ''            rARLL = Mid(dr.Item("ARLL"), 2)
            ' ''        Else
            ' ''            rARLL = Mid(dr.Item("ARLL"), 3)
            ' ''        End If
            ' ''    Else
            ' ''        rARLL = SpectNothing
            ' ''    End If
            ' ''Else
            ' ''    rTarget = SpectNothing
            ' ''    rLSL = SpectNothing
            ' ''    rUSL = SpectNothing
            ' ''    rLCL = SpectNothing
            ' ''    rUCL = SpectNothing
            ' ''    rARUL = SpectNothing
            ' ''    rARLL = SpectNothing

            ' ''End If

            'get isnumeric info
            If options = 0 Then
                strSQL = "select Isnumeric from mreportparam where oid = '" & strQCParamOID & "'"
            Else
                strSQL = "select distinct Isnumeric from mreportparam where REPORTPARANAME = '" & paramname & "'"
            End If

            cmd = New OleDbCommand(strSQL, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)

            If Not IsNothing(dsQC.Tables("QCPARAMETERINFO")) Then
                dsQC.Tables("QCPARAMETERINFO").Clear()
            End If
            ' dsQC.Tables("QCPARAMETERINFO").Clear()


            adapter.Fill(dsQC, "QCPARAMETERINFO")
            If dsQC.Tables("QCPARAMETERINFO").Rows.Count > 0 Then
                Isnumeric = dsQC.Tables("QCPARAMETERINFO").Rows(0).Item("isnumeric")
            End If

            con.Close()
        Catch ex As Exception
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            Throw New Exception(ex.ToString, ex)
        End Try
    End Sub
    Private Sub LoadInfoIDScanSpec(ByVal ShippingItem As String, ByVal options As Integer)
        Try
            Dim drIDScanSpec As DataRow
            drIDScanSpec = GetIDScanSpecByItemName(ShippingItem, options)
            With drIDScanSpec
                If Not IsDBNull(drIDScanSpec.Item("LSL")) Then
                    'lbLSL.Text = "LSL: " & .Item("LSL")
                    rLSL = .Item("LSL")
                Else
                    rLSL = SpectNothing
                End If
                If Not IsDBNull(drIDScanSpec.Item("Target")) Then
                    'lbTarget.Text = "Target: " & .Item("Target")
                    rTarget = .Item("Target")
                Else
                    rTarget = SpectNothing
                End If
                If Not IsDBNull(drIDScanSpec.Item("USL")) Then
                    'lbUCL.Text = "UCL: " & .Item("UCL")
                    rUSL = .Item("USL")
                Else
                    rUSL = SpectNothing
                End If
                If Not IsDBNull(drIDScanSpec.Item("UCL")) Then
                    'lbUCL.Text = "UCL: " & .Item("UCL")
                    rARUL = .Item("UCL")
                Else
                    rARUL = SpectNothing
                End If
                If Not IsDBNull(drIDScanSpec.Item("LCL")) Then
                    'lbUCL.Text = "UCL: " & .Item("UCL")
                    rARLL = .Item("LCL")
                Else
                    rARLL = SpectNothing
                End If
            End With
            rUCL = SpectNothing
            rLCL = SpectNothing
        Catch ex As Exception

            Throw New Exception(ex.ToString, ex)
        End Try
    End Sub
    Private Function GetIDScanSpecByItemName(ByVal ShippingItem As String, ByVal options As Integer) As DataRow
        Dim strCon As String
        Dim strSQL As String
        Dim cmd As New OleDbCommand
        Dim adapter As OleDbDataAdapter = New OleDbDataAdapter
        Dim con As OleDbConnection
        Dim dsMaster As New DataSet
        Dim drResult As DataRow

        Try
            strCon = sConnectionString
            con = New OleDbConnection(strCon)
            con.Open()
            cmd.Connection = con
            'Get Spec
            If options = 0 Then
                strSQL = "select * from MQCIDSCANSPEC where itemname ='" & ShippingItem & "' and rownum = 1"
            Else
                strSQL = "select max(USL) as usl,min(LSL) as lsl,max(TARGET) as target,Max(UCL) as ucl,Min(LCL) as lcl from MQCIDSCANSPEC where ITEMGROUP_NAME ='" & ShippingItem & "' and rownum = 1"
            End If

            cmd.CommandText = strSQL
            If Not dsQC.Tables("MQCIDSCANSPEC") Is Nothing Then
                dsQC.Tables("MQCIDSCANSPEC").Clear()
            End If
            adapter.SelectCommand = cmd
            adapter.Fill(dsQC, "MQCIDSCANSPEC")
            If dsQC.Tables("MQCIDSCANSPEC").Rows.Count > 0 Then
                drResult = dsQC.Tables("MQCIDSCANSPEC").Rows(0)
            Else
                drResult = Nothing
            End If
            con.Close()
        Catch ex As Exception
            con.Close()
            Throw ex
        End Try
        Return drResult
    End Function
#End Region
#Region "Process data"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check for generated or no
    ''' </summary>
    ''' <param name="ShippingItemOID"></param>
    ''' <param name="dateView"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[dmquyen]	08/10/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function CheckGenerated(ByVal ShippingItemName As String, ByVal dateView As String, ByVal options As Integer) As Boolean
        Dim strSQL, strCon As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection

        Try

            strCon = sConnectionString_stb
            con = New OleDbConnection(strCon)
            con.Open()
            'condition with QC daily information is datatype =0
            strSQL = "Select * From QCTRENDIMAGE Where QCDATE = " & dateView & " and  SHIPPINGITEMNAME = '" & ShippingItemName & "' and rownum=1"

            If options = 0 Then
                strSQL &= " and datatype =0 "
            Else ' datatype =3 Generate follow Group Item QC Daily Information 
                strSQL &= " and datatype = 3 "
            End If

            cmd = New OleDbCommand(strSQL, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)
            If Not IsNothing(dsQC.Tables("QCTRENDIMAGE")) Then
                dsQC.Tables("QCTRENDIMAGE").Clear()
            End If
            'Dim dsQCTRENDIMAGE As DataSet = New DataSet
            adapter.Fill(dsQC, "QCTRENDIMAGE")
            con.Close()
            If dsQC.Tables("QCTRENDIMAGE").Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            con.Close()
            Throw ex
        End Try
    End Function
    Private Function CheckExistData(ByVal ShippingItemOID As String) As Boolean
        Dim strSQL, strCon As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection
        Try

            Dim strFPDDate As String
            Dim strTPDDate As String
            strCon = sConnectionString_stb
            con = New OleDbConnection(strCon)
            con.Open()
            strTPDDate = Today.ToString("yyyyMMdd")
            strFPDDate = Today.AddDays(-7).ToString("yyyyMMdd")
            strSQL = "select * from dqcmaintransaction b " &
            " where b.pddate between " & strFPDDate & " and " & strTPDDate &
            " and b.spitemoid='" & ShippingItemOID & "' AND ROWNUM = 1"
            cmd = New OleDbCommand(strSQL, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)
            If Not IsNothing(dsQC.Tables("QCCHECKDATA")) Then
                dsQC.Tables("QCCHECKDATA").Clear()
            End If
            'Dim dsQCTRENDIMAGE As DataSet = New DataSet
            adapter.Fill(dsQC, "QCCHECKDATA")
            con.Close()
            If dsQC.Tables("QCCHECKDATA").Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            con.Close()
            Throw New Exception(ex.ToString, ex)
        End Try
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load data for parameter (for each shipping item and operation)
    ''' </summary>
    ''' <param name="ShippingItemOID"></param>
    ''' <param name="OperationOiD"></param>
    ''' <param name="Parameter"></param>
    ''' <param name="DateView"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[dmquyen]	07/26/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadDataForParam(ByVal ShippingItemOID As String, ByVal OperationOiD As String, ByVal ReportParameterOID As String, ByVal reportname As String, ByVal DateView As String, ByVal options As Integer, ByVal facode As String, tablename As String)

        Dim strFPDDate As String
        Dim strTPDDate As String
        Dim strOrder As String
        Dim strOperationOID As String

        Dim strSQLParaName As String
        Dim strSQLParamNickName As String
        Dim strSQLParaValue, strSQLCase, strSQLSelect, strSQLWhere As String
        Dim strSQL_AllPara, strSQL As String
        Dim dr As DataRow
        Dim dtResult As DataTable
        Dim strSelect As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection
        Dim adapter As OleDbDataAdapter
        Try

            strTPDDate = Today.ToString("yyyyMMdd")
            'strFPDDate = Today.AddDays(-7).ToString("yyyyMMdd")
            strFPDDate = Today.AddDays(-Convert.ToInt16(get_pddate)).ToString("yyyyMMdd")
            'Added by dmquyen 18/7/2006 --chuyển tạm thời params của texturing sang Final Cleaning
            'Đối với Shipping Item SAP
            '---------------------------------------------------
            If ShippingItemOID = "3318875c-59dd-4e43-8" And OperationOiD = "1000" Then
                OperationOiD = "1100"
            End If
            '----------------------------------------------------
            Dim strCon As String = sConnectionString_gdis
            '--------------------------------------------------------------------
            'Get all paramnickname from mqcstdparam
            strSQLParamNickName = "select distinct replace(t2.nickname,'  ',' ') as nickname" &
            " from mreport_qcstdparam t1" &
            " inner join mqcstdparam t2 on t1.mqcstdparamoid = t2.oid"
            If options = 0 Then
                strSQLParamNickName &= " where t1.mreportparamoid = '" & ReportParameterOID & "'" &
            " order by nickname"
            Else
                strSQLParamNickName &= " where t1.mreportparamoid in (select distinct oid from mreportparam where REPORTPARANAME ='" & reportname & "'" &
                                        "  and SHIPPINGOID in (select OID from mshippingitem where grp_oid= '" & ShippingItemOID & "') " &
                                        " and operationoid = '" & OperationOiD & "'" &
                                        "and facode = '" & facode & "' )" &
            " order by nickname"
            End If
            con = New OleDbConnection(strCon)
            con.Open()
            cmd = New OleDbCommand(strSQLParamNickName, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)
            'Clear data before fill data
            If Not IsNothing(dsQC.Tables("PARANICKNAME")) Then
                dsQC.Tables("PARANICKNAME").Clear()
            End If
            adapter.Fill(dsQC, "PARANICKNAME")
            'Fill into Cache.Item("arrPara") = arrParam
            '-------------------------------------------
            If arrParam.Count > 0 Then
                arrParam.Clear()
            End If

            For Each dr In dsQC.Tables("PARANICKNAME").Rows
                arrParam.Add(dr.Item("NICKNAME"))
            Next
            '---------------------------------------------------------------------------
            If options = 0 Then
                strSQL_AllPara = "Select slipno as SlipNo,pddate as PDDate,pdshift as PDShift,PDtime,diskno, case  when pdshift='NIGHT' and length(pdtime) <4 then pdtime + 2400 else pdtime end  ""pdtime1"",qcchecksheetoid,judgment,ll,ul,arll,arul,target "
                strSQLParaValue = " slipno,pddate,pdshift,PDtime,diskno,replace(stdparamnickname,'  ',' ') as PARANICKNAME,qcchecksheetoid, valueview as Paramvalue,judgment,ll,ul,arll,arul,target "
                strSQLParaName = " slipno,pddate,pdshift,PDtime,diskno,qcchecksheetoid,judgment,ll,ul,arll,arul,target"
            Else
                strSQL_AllPara = "Select slipno as SlipNo,pddate as PDDate,pdshift as PDShift,PDtime,diskno, case  when pdshift='NIGHT' and length(pdtime) <4 then pdtime + 2400 else pdtime end  ""pdtime1"",qcchecksheetoid,judgment,Min(ll) as ll,max(ul) as ul,min(arll) as arll,max(arul) as arul,max(target) as target "
                strSQLParaValue = " slipno,pddate,pdshift,PDtime,diskno,replace(stdparamnickname,'  ',' ') as PARANICKNAME,qcchecksheetoid, valueview as Paramvalue,judgment,ll,ul,arll,arul,target "
                strSQLParaName = " slipno,pddate,pdshift,PDtime,diskno,qcchecksheetoid,judgment"

            End If
            '-----------------------------------------------------------------

            Dim strSQLFrom As String

            'strSQLFrom &= "  from dpd_qc_info_new a "
            'get data from separate tables
            Dim strIndexHint, strIndexHintQC As String
            ''Select Case OperationOiD
            ''    Case "1000", "1500"
            ''        strSQLFrom &= "  from dpd_qc_info_fcl a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_fcl_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_fcl_qcdate) */"
            ''    Case "600"
            ''        strSQLFrom &= "  from dpd_qc_info_2P a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_2p_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_2p_qcdate) */"
            ''    Case "500"
            ''        strSQLFrom &= "  from dpd_qc_info_1P a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_1p_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_1p_qcdate) */"
            ''    Case "800"  '1PVI e damage

            ''        strSQLFrom &= "  from dpd_qc_info_1P a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_1p_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_1p_qcdate) */"
            ''    Case "420"

            ''        strSQLFrom &= "  from dpd_qc_info_OD a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_OD_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_od_qcdate) */"
            ''    Case "410"

            ''        strSQLFrom &= "  from dpd_qc_info_ID a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_id_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_id_qcdate) */"
            ''    Case "300"

            ''        strSQLFrom &= "  from dpd_qc_info_lap2 a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_lap2_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_lap2_qcdate) */"
            ''    Case "220"

            ''        strSQLFrom &= "  from dpd_qc_info_charmfer a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_charmfer_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_charmfer_qcdate) */"
            ''    Case "100"

            ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_lap1_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_lap1_qcdate) */"
            ''    Case "150"

            ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_lap1_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_lap1_qcdate) */"
            ''    Case "200"

            ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_lap1_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_lap1_qcdate) */"
            ''    Case "120"

            ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_lap1_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_lap1_qcdate) */"
            ''    Case "210"

            ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ''        strIndexHint = "/*+ index(a idx_qcinfo_lap1_dateparamround) */"
            ''        strIndexHintQC = "/*+ index(a idx_qcinfo_lap1_qcdate) */"
            ''    Case Else

            ''        strSQLFrom &= "  from dpd_qc_info a "
            ''        strIndexHint = "/*+ index(a IDX_DPDQCINFO_PARAMMAXR) */"
            ''        strIndexHintQC = "/*+ index(a IDX_DPDQINFSTDPISMAXQCDATE) */"

            ''End Select
            strSQLFrom &= "  from " & tablename & " a "
            strIndexHint = ""
            strIndexHintQC = ""
            If options <> 0 Then
                If OperationOiD = "150099" Or OperationOiD = "150199" Then
                    OperationOiD = "1000"
                End If
            End If

            strSQLWhere = " where  "
            '- Set query SHIFT - DATE
            strSQLWhere &= " pddate between " & strFPDDate & " and " & strTPDDate
            'strSQLWhere &= " and   pddate >='" & strFPDDate & "'"
            'strSQLWhere &= " and   pddate  <='" & strTPDDate & "'"
            If options = 0 Then
                strSQLWhere &= " and a.stdparamreportoid = '" & ReportParameterOID & "'"
            Else
                strSQLWhere &= " and a.STDPARAMREPORTNAME= '" & reportname & "'" &
                    " and a.GROUPITEMOID ='" & ShippingItemOID & "'" &
                                                   " and a.facode ='" & facode & "' and a.operationoid ='" & OperationOiD & "'"
            End If

            'for old data and new data
            If strFPDDate >= 20070402 Then
                strSQLWhere &= " and a.Ismaxqcround = '1' "
            Else
                strSQLWhere &= " and a.qcround = 1"
            End If
            'strSQLWhere &= " and (a.calculatebudo = 1 or calculatebudo is null)"
            'strSQLWhere &= " AND PARAMVALUE IS NOT NULL"


            'added by dmquyen 27/6/2006
            strOrder = " pddate,pdshift,""pdtime1"", slipno,diskno " 'PDTIME

            '-------------------------------------------
            'GET dtData - strSQLParavalue
            '--------------------------------------
            'Create cross tab --there're 2 cases: textbox, combobox
            Dim i As Integer
            For i = 0 To arrParam.Count - 1
                strSQLCase &= " ,SUM(Decode(b.PARANICKNAME,'" & arrParam(i) & "' , IS_NUMERIC(replace(replace(b.Paramvalue,',',''),' ','')),null)) """ & arrParam(i) & """"
            Next

            strSQL = strSQL_AllPara & strSQLCase & " From( " & "Select " & strIndexHint & strSQLParaValue & strSQLFrom & strSQLWhere & " ) b" & " Group by " & strSQLParaName & " Order by " & strOrder

            '-------------------------------------------
            'GET dtData - strSQLParavalue
            '--------------------------------------

            cmd.CommandText = strSQL

            If Not IsNothing(dsQC.Tables("PARANAME")) Then
                dsQC.Tables("PARANAME").Clear()
            End If
            adapter.Fill(dsQC, "PARANAME")
            '---------------------------------------------------------------
            dtData = dsQC.Tables("PARANAME")

            con.Close()

        Catch ex As Exception
            con.Close()
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub
    ''' <summary>
    ''' 
    ''' Load data for non-numeric parameter
    ''' </summary>
    ''' <param name="ShippingItemOID"></param>
    ''' <param name="facodeoid"></param>
    ''' <param name="DateView"></param>
    ''' <remarks></remarks>
    Private Sub LoadDataForNotnumeric(ByVal ReportParameterOID As String, OperationOiD As String, ByVal reportname As String, ByVal grpitemoid As String, ByVal facode As String, ByVal options As Integer, tablename As String)
        Dim strFPDDate As String
        Dim strTPDDate As String
        Dim strOrder As String

        Dim strSQLParaName As String

        Dim strMainSQL, strGroupOrderBy, strSQLWhere, strSQLFrom, strSQL As String
        strGroupOrderBy = ""
        Dim dtResult As DataTable

        Dim reportoid As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection
        Dim adapter As OleDbDataAdapter
        Try

            strTPDDate = Today.ToString("yyyyMMdd")
            strFPDDate = Today.AddDays(-Convert.ToInt16(get_pddate)).ToString("yyyyMMdd")
            Dim strCon As String = sConnectionString_gdis

            strMainSQL = "SELECT DATESHIFT,PDDATE, PDSHIFT,SUM(PCS) AS PCS,SUM(SUM(PCS)) OVER (PARTITION BY     DATESHIFT) TOTAL_PCS ," &
                                " ROUND(SUM(PCS) / SUM(SUM(PCS)) OVER (PARTITION BY DATESHIFT),4) AS  PCS_PERCENT,PARAMVALUE " &
                                       "FROM ( "

            strSQLParaName = "SELECT  a.pddate || SUBSTR(a.pdshift,1,1) as DATESHIFT," &
                                     " A.PDDATE,A.PDSHIFT, count(*) as Pcs, UPPER(paramvalue) as paramvalue "


            strSQLFrom = " from " & tablename & " a"

            'get data from separate tables
            ' ''Select Case OperationOiD
            ' ''    Case "1000"
            ' ''        strSQLFrom &= "  from dpd_qc_info_fcl a "
            ' ''    Case "600"
            ' ''        strSQLFrom &= "  from dpd_qc_info_2P a "
            ' ''    Case "500"
            ' ''        strSQLFrom &= "  from dpd_qc_info_1P a "
            ' ''    Case "800"  '1PVI e damage

            ' ''        strSQLFrom &= "  from dpd_qc_info_1P a "
            ' ''    Case "420"

            ' ''        strSQLFrom &= "  from dpd_qc_info_OD a "
            ' ''    Case "410"

            ' ''        strSQLFrom &= "  from dpd_qc_info_ID a "
            ' ''    Case "300"

            ' ''        strSQLFrom &= "  from dpd_qc_info_lap2 a "
            ' ''    Case "220"

            ' ''        strSQLFrom &= "  from dpd_qc_info_charmfer a "
            ' ''    Case "100"

            ' ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ' ''    Case "150"

            ' ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ' ''    Case "200"

            ' ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ' ''    Case "120"

            ' ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ' ''    Case "210"

            ' ''        strSQLFrom &= "  from dpd_qc_info_Lap1 a "
            ' ''    Case Else

            ' ''        strSQLFrom &= "  from dpd_qc_info a "
            ' ''End Select
            ' strSQLFrom &= "  from dpd_qc_info_cal_cpk a where 1=1"

            strSQLWhere = ""
            '- Set query SHIFT - DATE

            '- Set query DATE
            If strFPDDate = strTPDDate Then
                strSQLWhere &= " where a.pddate =" & strFPDDate
            Else
                strSQLWhere &= " where a.pddate between " & strFPDDate & " and " & strTPDDate
            End If


            strOrder = " b.pddate,b.pdshift,b.spitemname,PDProcess, ""pdtime1"", b.slipno,b.diskno " 'PDTIME

            If options = 0 Then
                strSQLWhere &= " and a.stdparamreportoid = '" & ReportParameterOID & "'"
            Else
                strSQLWhere &= " and a.STDPARAMREPORTNAME = '" & reportname & "'" &
                    " and a.groupitemoid ='" & grpitemoid & "'" &
                    " and a.facode ='" & facode & "'"
            End If



            strSQLWhere &= " and a.Ismaxqcround = '1' "

            strGroupOrderBy = " GROUP BY a.pddate || SUBSTR(a.pdshift,1,1)," &
                                           " A.PDDATE,A.PDSHIFT, paramvalue " &
                                           " ) GROUP BY DATESHIFT,PDDATE, PDSHIFT,PARAMVALUE" &
                                           " ORDER BY PDDATE,PDSHIFT,paramvalue"


            strSQL = strMainSQL & strSQLParaName & strSQLFrom & strSQLWhere & strGroupOrderBy

            '-------------------------------------------
            'GET dtData - strSQLParavalue
            '--------------------------------------
            ' cmd.CommandText = strSQL
            con = New OleDbConnection(strCon)
            con.Open()
            cmd = New OleDbCommand(strSQL, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)
            If Not dsQC.Tables("PARANAME") Is Nothing Then
                dsQC.Tables("PARANAME").Clear()
            End If
            ' Adapter.Fill(dsQC, "PARANAME")
            cmd.CommandText = strSQL
            If Not IsNothing(dsQC.Tables("PARANAME")) Then
                dsQC.Tables("PARANAME").Clear()
            End If
            adapter.Fill(dsQC, "PARANAME")

            dtData = dsQC.Tables("PARANAME")

            con.Close()
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub
    Private Sub LoadDataIDScan(ByVal ShippingItemOID As String, ByVal facodeoid As String, ByVal DateView As String, ByVal options As Integer)

        Dim strFPDDate As String
        Dim strTPDDate As String
        Dim strOrder As String
        Dim strOperationOID As String

        Dim strSQLParaName As String
        Dim strSQLParamNickName As String
        Dim strSQLParaValue, strSQLCase, strSQLSelect, strSQLWhere As String
        Dim strSQL_AllPara, strSQL As String
        Dim dr As DataRow
        Dim dtResult As DataTable
        Dim strSelect As String
        Dim i As Integer
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection
        Dim adapter As OleDbDataAdapter
        Try

            strTPDDate = Today.ToString("yyyyMMdd")
            'strFPDDate = Today.AddDays(-7).ToString("yyyyMMdd")
            strFPDDate = Today.AddDays(-Convert.ToInt16(get_pddate)).ToString("yyyyMMdd")

            Dim strCon As String = sConnectionString_stb
            'Fill into Cache.Item("arrPara") = arrParam
            '-------------------------------------------
            If arrParam.Count > 0 Then
                arrParam.Clear()
            End If
            arrParam.Add("MaxIDValue")
            arrParam.Add("MinIDValue")
            arrParam.Add("AVGIDValue")

            '-------------------------IDScan FVI -------------------------------------------
            strSQLSelect = "Select slipno ,PDDate,PDShift,MaxIDValue,MinIDValue,AVGIDvalue,USL as UL, LSL as LL,Target " &
                            " From laserid_summary " &
                            " where pddate between " & strFPDDate & " and " & strTPDDate &
                            " and SHPITEMOID ='" & ShippingItemOID & "'" &
                            " and FACODE ='" & facodeoid & "'" &
                            " and maxIDvalue between 20 and 20.06" &
                            " and minIDvalue between 20 and 20.06" &
                            " and AVGIDvalue between 20 and 20.06"

            '-------------------------IDScan AOI -------------------------------------------
            strSQLSelect &= " Union " &
                            " Select slipno ,PDDate,PDShift,MaxIDValue,MinIDValue,AVGIDvalue,USL as UL, LSL as LL,target " &
                            " from IDScan_AOI_Summary " &
                            " where pddate between " & strFPDDate & " and " & strTPDDate &
                            " and SHPITEMOID ='" & ShippingItemOID & "'" &
                            " and FACCODE ='" & facodeoid & "'" &
                            " and AVGIDvalue is not null" &
                            " and maxIDvalue between 20 and 20.06" &
                            " and minIDvalue between 20 and 20.06" &
                            " and AVGIDvalue between 20 and 20.06" &
                            " Order by PDDate, PDShift, slipno"


            '--------------------------------------------------------------------
            If options = 1 Then
                strSQLSelect = strSQLSelect.Replace(" and SHPITEMOID ='" & ShippingItemOID & "'", " and SHPITEMOID in(select OID from mshippingitem where grp_oid= '" & ShippingItemOID & "')")
            End If


            con = New OleDbConnection(strCon)
            con.Open()
            cmd = New OleDbCommand(strSQLSelect, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)
            'Clear data before fill data
            If Not IsNothing(dsQC.Tables("IDScanData")) Then
                dsQC.Tables("IDScanData").Clear()
            End If
            adapter.Fill(dsQC, "IDScanData")
            '---------------------------------------------------------------
            dtData = dsQC.Tables("IDScanData")

            con.Close()

        Catch ex As Exception
            con.Close()
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub

#End Region
#Region "Draw trend"
    Private Sub CreatTableResult()
        Dim resultTable As New DataTable
        Dim dtCol As DataColumn
        With resultTable
            dtCol = New DataColumn("ShippingItem", GetType(System.String))
            .Columns.Add(dtCol)
            dtCol = New DataColumn("OperationName", GetType(System.String))
            .Columns.Add(dtCol)
            dtCol = New DataColumn("ParameterName", GetType(System.String))
            .Columns.Add(dtCol)
            dtCol = New DataColumn("TrendImage", GetType(System.Byte()))
        End With
        If IsNothing(dsQC.Tables("resultTable")) Then
            dsQC.Tables.Add(resultTable)
        End If
    End Sub
    Private Function InsertImageIntoTableResult(ByVal ShippingItemOID As String, ByVal item_Grpitem As String, ByVal OperationOiD As String, operationname As String, ByVal ParameterOID As String, ByVal parametername As String, ByVal DateView As String, ByVal bGenerated As Boolean, ByVal FacodeOid As String, ByVal options As Integer, tablename As String, Optional ByVal IsIDScan As Boolean = False) As System.IO.Stream
        'Dim Item_GroupItemName As String
        Dim ParamNickname As String
        Dim dtRow As DataRow
        'Dim streamimage As System.IO.Stream
        Dim fStream As System.IO.FileStream
        Dim bitmapStream As System.IO.Stream
        Dim g As System.Drawing.Graphics
        Dim image As System.Drawing.Image
        Dim bloc() As Byte
        Dim strParamvaluelst As String = ""
        Dim strColorlst As String = ""
        Dim strBmpFile As String = System.Windows.Forms.Application.StartupPath & "..\ExportImage\BitmapImage.bmp"
        Try

            If IsIDScan Then 'normal parameter

                operationname = "Final Cleaning"
                parametername = "IDscan"
            End If
            '  End If

            'Load data before drawing trend
            'LoadDataForParam(ShippingItemOID, OperationOiD, ParameterOID, DateView)
            Try
                ChartTrend = New ChartFX.WinForms.Chart
                If Not IsIDScan Then 'normal parameter
                    If Isnumeric = 1 Then
                        LoadDataForParam(ShippingItemOID, OperationOiD, ParameterOID, parametername, DateView, options, FacodeOid, tablename) 'On Hoyav3_stb
                    Else
                        LoadDataForNotnumeric(ParameterOID, OperationOiD, parametername, ShippingItemOID, FacodeOid, options, tablename)
                    End If
                Else 'IDScan parameter
                    LoadDataIDScan(ShippingItemOID, FacodeOid, DateView, options)
                End If

                If dtData.Rows.Count > 0 Then

                    If Isnumeric = 1 Then
                        If IsIDScan Then
                            DrawTrend(ChartTrend, dtData, "PDSHIFT", item_Grpitem, options, "ALL", rTarget, rLSL, rUSL, rARUL, rARLL, operationname, parametername, FacodeOid, True)
                        Else
                            DrawTrend(ChartTrend, dtData, "PDSHIFT", item_Grpitem, options, "ALL", rTarget, rLSL, rUSL, rARUL, rARLL, operationname, parametername, FacodeOid)
                        End If

                    Else
                        '  strParamvaluelst = LoadParamValue(OperationOiD, ParameterOID, parametername, FacodeOid, options, ShippingItemOID)
                        LoadParamValue(OperationOiD, ParameterOID, parametername, FacodeOid, options, ShippingItemOID, strParamvaluelst, strColorlst)
                        DrawBar(ChartTrend, dtData, strParamvaluelst, item_Grpitem, options, parametername, strColorlst)
                    End If



                    If options = 0 Then
                        strBmpFile = System.Windows.Forms.Application.StartupPath & "\ExportImage\ItemImage\"
                    Else
                        strBmpFile = System.Windows.Forms.Application.StartupPath & "\ExportImage\GrpImage\"
                    End If
                    ' namlx. tao folder.
                    If Not Directory.Exists(strBmpFile) Then
                        Directory.CreateDirectory(strBmpFile)
                    End If

                    'strBmpFile &= item_Grpitem & operationname & parametername.Replace(">", "").Replace("<", "").Replace("*", "").Replace(":", "") & FacodeOid & ".bmp"
                    strBmpFile &= "BitmapImage.bmp"
                    ' ''Dim strBmpFile As String = FullPath & "\" & shippingItemName & OperationName & ParameterName.Replace(">", "").Replace("<", "").Replace("*", "") & FacodeOid & ".bmp"
                    bitmapStream = New FileStream(strBmpFile, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write)
                    Try
                        ChartTrend.Export(ChartFX.WinForms.FileFormat.Bitmap, bitmapStream)
                        bitmapStream.Close()
                    Catch ex As Exception
                        Trace.WriteLineIf(TRACE_ENABLE, "Error on exporting chartfx" & vbCrLf & ex.ToString())
                        'bitmapStream = New System.IO.FileStream(strBmpFile, IO.FileMode.Open, IO.FileAccess.Read)
                        bitmapStream.Dispose()

                    End Try


                    Dim strPNGFile As String
                    If options = 0 Then
                        strPNGFile = System.Windows.Forms.Application.StartupPath & "\ExportImage\ItemImage\"
                    Else
                        strPNGFile = System.Windows.Forms.Application.StartupPath & "\ExportImage\GrpImage\"
                    End If

                    'strPNGFile &= item_Grpitem & operationname & parametername.Replace(">", "").Replace("<", "").Replace("*", "").Replace(":", "") & FacodeOid & ".png"
                    strPNGFile &= "BitmapImage.png"
                    ' ''Dim strPNGFile As String = FullPath & "\" & shippingItemName & OperationName & ParameterName.Replace(">", "").Replace("<", "").Replace("*", "") & FacodeOid & ".png"
                    Dim sImagePath As String
                    ' ''If options = 0 Then
                    ' ''    sImagePath = FullPath & "\" & item_Grpitem & operationname & parametername.Replace(">", "").Replace("<", "").Replace("*", "").Replace(":", "") & FacodeOid & ".png"
                    ' ''Else
                    ' ''    sImagePath = FullPath & "\GrpImage" & item_Grpitem & operationname & parametername.Replace(">", "").Replace("<", "").Replace("*", "").Replace(":", "") & FacodeOid & ".png"
                    ' ''End If
                    'Convert from bmp format to png format
                    'ConvertbmpStream(bitmapStream, strPNGFile, ImageFormat.Png)
                    ConvertBMP(strBmpFile, ImageFormat.Png)
                    'Open the stream to the PNG File
                    fStream = New System.IO.FileStream(strPNGFile, IO.FileMode.Open, IO.FileAccess.Read)

                    bloc = New Byte(fStream.Length) {}
                    fStream.Read(bloc, 0, fStream.Length)
                    fStream.Close()
                    sImagePath = ""
                    '---------------------------
                    'streamimage.ge()
                    'bGenerated = False
                    'MsgBox("update db")
                    SaveImage(ShippingItemOID, item_Grpitem, OperationOiD, operationname, ParameterOID, parametername, DateView, bloc, bGenerated, FacodeOid, options, sImagePath) 'On Hoyav3
                    fStream.Dispose()
                    'MsgBox("End update db")
                    bitmapStream.Dispose()

                End If
            Catch ex As Exception
                Trace.WriteLineIf(TRACE_ENABLE, "===Error : QCEmail.InsertImageIntoTableResult - Parameter : " & parametername & ". Nhung van bo qua de chay tiep parameter sau")
                Trace.WriteLineIf(TRACE_ENABLE, ex.ToString())
                Trace.WriteLineIf(TRACE_ENABLE, "==End Parameter Error : " & parametername)
                'SendEmailOnError("Generate QC Daily Information at QCEmail.InsertImageIntoTableResult  For ITEM :" & shippingItemName, ex.ToString())
                'Delete(strBmpFile)
                If Not bitmapStream Is Nothing Then
                    bitmapStream.Dispose()
                End If
            End Try
        Catch ex As Exception
            Throw New Exception(ex.ToString, ex)
        End Try
    End Function
    Public Function ConvertbmpStream(ByVal BmpStream As FileStream, pngFile As String,
        ByVal imgFormat As ImageFormat) As Boolean

        Dim bAns As Boolean

        Try

            'bitmap class in system.drawing.imaging
            Dim objBmp As New Bitmap(BmpStream, False)
            Dim Eps As New EncoderParameters(1)
            Dim ici As ImageCodecInfo
            ici = GetEncoderInfo("image/png") '"image/png"

            Dim lCompression As Long
            Eps.Param(0) = New EncoderParameter(Encoder.Quality, lCompression)

            'below 2 functions in system.io.path
            objBmp.Save(pngFile, ici, Eps)
            objBmp.Dispose()
            bAns = True 'return true on success
        Catch ex As Exception
            bAns = False 'return false on error
        End Try
        Return bAns

    End Function
    Public Function ConvertBMP(ByVal BMPFullPath As String,
        ByVal imgFormat As ImageFormat) As Boolean

        Dim bAns As Boolean
        Dim sNewFile As String

        Try
            'bitmap class in system.drawing.imaging
            Dim objBmp As New Bitmap(BMPFullPath)
            Dim Eps As New EncoderParameters(1)
            Dim ici As ImageCodecInfo
            ici = GetEncoderInfo("image/png") '"image/png"

            Dim lCompression As Long
            Eps.Param(0) = New EncoderParameter(Encoder.Quality, lCompression)

            'below 2 functions in system.io.path
            sNewFile = GetDirectoryName(BMPFullPath)
            sNewFile &= "\" & GetFileNameWithoutExtension(BMPFullPath)

            sNewFile &= "." & imgFormat.ToString
            objBmp.Save(sNewFile, ici, Eps)
            objBmp.Dispose()
            bAns = True 'return true on success
        Catch ex As Exception
            bAns = False 'return false on error
        End Try
        Return bAns

    End Function
    Public Function GetEncoderInfo(ByVal sType As String)
        Dim i As Integer
        Dim Encodes As ImageCodecInfo()
        Encodes = ImageCodecInfo.GetImageEncoders
        For i = 0 To Encodes.Length - 1
            If Encodes(i).MimeType = sType Then
                Return Encodes(i)
            End If
        Next
        Return Nothing
    End Function
    Private Sub SaveImage(ByVal ShippingItemOID As String, ByVal ShippingItemName As String, ByVal OperationOID As String, ByVal OperationName As String, ByVal ParameterOID As String, ByVal parameterName As String, ByVal dateView As String, ByVal TrendImage As Byte(), ByVal bGenerated As Boolean, ByVal FacodeOid As String, ByVal options As Integer, Optional ByVal ImagePath As String = "")
        Dim strSQL, strCon As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection
        Try

            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter

            strCon = sConnectionString
            con = New OleDbConnection(strCon)
            con.Open()
            cmd = New OleDbCommand
            cmd.Connection = con
            '----DELETE BEFORE INSERT
            If bGenerated Then
                strSQL = "DELETE FROM QCTRENDIMAGE where ShippingItemOID = '" & ShippingItemOID & "' and OperationOID = '" & OperationOID & "' and  qcDate = " & dateView & " and facode = " & FacodeOid
                If options = 0 Then
                    strSQL &= " and ParameterOID = '" & ParameterOID & "' and datatype = 0 "
                Else
                    strSQL &= " and PARAMETERNAME = '" & parameterName & "' and datatype = 3  "
                End If
                cmd.CommandType = CommandType.Text
                cmd.CommandText = strSQL
                cmd.ExecuteNonQuery()
            End If

            If options = 0 Then
                strSQL = "INSERT INTO QCTRENDIMAGE VALUES('" & Guid.NewGuid.ToString & "', '" & ShippingItemOID & "', '" & ShippingItemName & "','" & OperationOID & "', '" & OperationName & "', '" & ParameterOID & "', '" & parameterName & "',to_number( " & Today.ToString("yyyyMMdd") & "), " & " :bTrendImage" & ", '" & Date.Now.ToLongTimeString & "', " & FacodeOid & ",'" & ImagePath & "'," & ServicePos & ",0)"
            Else
                strSQL = "INSERT INTO QCTRENDIMAGE VALUES('" & Guid.NewGuid.ToString & "', '" & ShippingItemOID & "', '" & ShippingItemName & "','" & OperationOID & "', '" & OperationName & "', '" & ParameterOID & "', '" & parameterName & "',to_number( " & Today.ToString("yyyyMMdd") & "), " & " :bTrendImage" & ", '" & Date.Now.ToLongTimeString & "', " & FacodeOid & ",'" & ImagePath & "'," & ServicePos & ",3)"
            End If

            cmd.CommandText = strSQL
            Dim para As New OleDbParameter   '("TrendImage", TrendImage)
            para.OleDbType = OleDbType.Binary
            para.ParameterName = "bTrendImage"
            para.Value = TrendImage
            'para.DbType = DbType.Binary
            cmd.Parameters.Add(para)
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            con.Close()
            Throw New Exception(ex.ToString, ex)
        End Try
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' drawing trend for each parameter
    ''' </summary>
    ''' <param name="ChartTrend"></param>
    ''' <param name="dtDrawTrend"></param>
    ''' <param name="strShiftType"></param>
    ''' <param name="strColumnName"></param>
    ''' <param name="dblTarget"></param>
    ''' <param name="dblLSL"></param>
    ''' <param name="dblUSL"></param>
    ''' <param name="dblLCL"></param>
    ''' <param name="dblUCL"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[dmquyen]	07/26/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawTrend(ByRef ChartTrend As ChartFX.WinForms.Chart, ByVal dtDrawTrend As DataTable, ByVal strShiftType As String, ByVal item_grpItem As String, ByVal options As Integer, Optional ByVal strColumnName As String = "ALL", Optional ByVal dblTarget As Double = Nothing, Optional ByVal dblLSL As Double = 0, Optional ByVal dblUSL As Double = 0, Optional ByVal dblARUL As Double = 0, Optional ByVal dblARLL As Double = 0, Optional ByVal OperationName As String = "", Optional ByVal ParameterName As String = "", Optional ByVal FacodeOid As String = "6", Optional ByVal IsIDScan As Boolean = False)
        Dim i As Integer = -1
        Dim j As Integer = -1
        Dim dblValue, rmax, rmin, rDown, rUp As Double

        Dim iLine As Integer = 0

        Dim sColName As String = ""
        Dim TotalSeries As Integer = 0
        Dim iCount As Integer
        Dim iPointTrend, iCountTrend, iFlag As Integer
        Dim SumX, SumY, XAvg, YAvg As Double
        'Dim dARUL, dARLL, dUL, dLL As Double
        dblARUL = SpectNothing
        dblARLL = SpectNothing
        dblUSL = SpectNothing
        dblLSL = SpectNothing
        dblTarget = SpectNothing
        Dim seritrendline As Integer
        Dim iCase As Integer
        Dim iCaseAll, iparamCount As Integer
        Dim k As Integer
        Try
            rmax = -100000
            rmin = 100000
            'ChartTrend = New ChartFX.WinForms.Chart
            With ChartTrend
                .Data.Clear()
                .Reset()
                .AxisY.CustomGridLines.Clear()
                .AxisX.CustomGridLines.Clear()
                .Series.Clear()
                .Points.Clear()
                .Data.Clear()
                .Data.Labels.Clear()
                .AxisX.Labels.Clear()
                .AxisX.Sections.Clear()
                .AxisY.Sections.Clear()
                .AxisY2.Labels.Clear()
                .AxisY2.Sections.Clear()

                .Height = 230 '312
                .Width = 600

            End With

            iCount = dtDrawTrend.Rows.Count

            Dim constSpect As ChartFX.WinForms.CustomGridLine
            Dim strOldDate As String
            Dim strOldShift As String = ""
            Dim strShift As String
            Dim l As Integer = arrParam.Count

            For iCol As Integer = 0 To dtDrawTrend.Rows.Count - 1

                'draw spec and Target
                'add by maidt.it 24/1/2012
                ' ''    'draw Spec
                If Not IsIDScan Then
                    If Not IsDBNull(dtDrawTrend.Rows(iCol).Item("ARUL")) Then
                        dblARUL = dtDrawTrend.Rows(iCol).Item("ARUL")
                        If dblARUL < rmin Then
                            rmin = dblARUL
                        End If
                        If dblARUL > rmax Then
                            rmax = dblARUL
                        End If
                    Else
                        dblARUL = SpectNothing
                    End If
                    'drawing ARLL
                    If Not IsDBNull(dtDrawTrend.Rows(iCol).Item("ARLL")) Then

                        dblARLL = dtDrawTrend.Rows(iCol).Item("ARLL")
                        If dblARLL < rmin Then
                            rmin = dblARLL
                        End If
                        If dblARLL > rmax Then
                            rmax = dblARLL
                        End If
                    Else
                        dblARLL = SpectNothing
                    End If
                End If

                'Drawing USL line
                If Not IsDBNull(dtDrawTrend.Rows(iCol).Item("UL")) Then

                    dblUSL = dtDrawTrend.Rows(iCol).Item("UL")
                    If dblUSL < rmin Then
                        rmin = dblUSL
                    End If
                    If dblUSL > rmax Then
                        rmax = dblUSL
                    End If
                Else
                    dblUSL = SpectNothing
                End If
                'drawing LSL line
                If Not IsDBNull(dtDrawTrend.Rows(iCol).Item("LL")) Then

                    dblLSL = dtDrawTrend.Rows(iCol).Item("LL")
                    If dblLSL < rmin Then
                        rmin = dblLSL
                    End If
                    If dblLSL > rmax Then
                        rmax = dblLSL
                    End If
                Else
                    dblLSL = SpectNothing
                End If
                'drawing target line
                If Not IsDBNull(dtDrawTrend.Rows(iCol).Item("Target")) Then
                    dblTarget = dtDrawTrend.Rows(iCol).Item("Target")
                Else
                    dblTarget = SpectNothing
                End If
                'drawing USL series
                If dblUSL <> SpectNothing Then
                    'TotalSeries = 0
                    ChartTrend.Data(0, iCol) = dblUSL
                    ChartTrend.Points(0, iCol).Text = dblUSL
                    ' ChartTrend.Data(k, iCol) = dblValue
                    ChartTrend.Data.X(0, iCol) = iCol
                    ChartTrend.Points(0, iCol).Color = Color.Red
                    'TotalSeries += 1
                Else
                    ChartTrend.Data(0, iCol) = Chart.Hidden
                End If
                'drawing target series
                If dblTarget <> SpectNothing Then
                    'TotalSeries = 1
                    ChartTrend.Data(1, iCol) = dblTarget
                    ChartTrend.Points(1, iCol).Text = dblTarget
                    ' ChartTrend.Data(k, iCol) = dblValue
                    ChartTrend.Data.X(1, iCol) = iCol
                    ChartTrend.Points(1, iCol).Color = Color.Cyan
                    ' TotalSeries += 1
                Else
                    ChartTrend.Data(1, iCol) = Chart.Hidden
                End If
                'drawing LSL series
                If dblLSL <> SpectNothing Then
                    'TotalSeries = 2
                    ChartTrend.Data(2, iCol) = dblLSL
                    ChartTrend.Points(2, iCol).Text = dblLSL
                    ' ChartTrend.Data(k, iCol) = dblValue
                    ChartTrend.Data.X(2, iCol) = iCol
                    '.Points(TotalSeries, iCol).Color = Color.Red
                    'TotalSeries += 1
                Else
                    ChartTrend.Data(2, iCol) = Chart.Hidden
                End If
                If Not IsIDScan Then
                    'drawing ARUL series
                    If dblARUL <> SpectNothing Then
                        ' TotalSeries = 3
                        ChartTrend.Data(3, iCol) = dblARUL
                        ChartTrend.Points(3, iCol).Text = dblARUL
                        ' ChartTrend.Data(k, iCol) = dblValue
                        ChartTrend.Data.X(3, iCol) = iCol
                        '.Points(TotalSeries, iCol).Color = Color.Red
                        ' TotalSeries += 1
                    Else
                        ChartTrend.Data(3, iCol) = Chart.Hidden
                    End If
                    'drawing ARLL series
                    If dblARLL <> SpectNothing Then
                        ' TotalSeries = 4
                        ChartTrend.Data(4, iCol) = dblARLL
                        ChartTrend.Points(4, iCol).Text = dblARLL

                        ' ChartTrend.Data(k, iCol) = dblValue
                        ChartTrend.Data.X(4, iCol) = iCol
                        '.Points(TotalSeries, iCol).Color = Color.Red
                        'TotalSeries += 1
                    Else
                        ChartTrend.Data(4, iCol) = Chart.Hidden
                    End If
                End If


                'TotalSeries = 0
                ''drawing USL series
                'If dblUSL <> SpectNothing Then
                '    ChartTrend.Data(TotalSeries, iCol) = dblUSL
                '    ChartTrend.Points(TotalSeries, iCol).Text = dblUSL

                '    ' ChartTrend.Data(k, iCol) = dblValue
                '    ChartTrend.Data.X(TotalSeries, iCol) = iCol
                '    ' ChartTrend.Points(TotalSeries, iCol).Color = Color.Red
                '    TotalSeries += 1
                'End If
                ''drawing LSL series
                'If dblLSL <> SpectNothing Then
                '    ChartTrend.Data(TotalSeries, iCol) = dblLSL
                '    ChartTrend.Points(TotalSeries, iCol).Text = dblLSL
                '    ' ChartTrend.Data(k, iCol) = dblValue
                '    ChartTrend.Data.X(TotalSeries, iCol) = iCol
                '    '.Points(TotalSeries, iCol).Color = Color.Red
                '    TotalSeries += 1
                'End If
                ''drawing Target series
                'If dblTarget <> SpectNothing Then
                '    ChartTrend.Data(TotalSeries, iCol) = dblTarget
                '    ChartTrend.Points(TotalSeries, iCol).Text = dblTarget
                '    ChartTrend.Data.X(TotalSeries, iCol) = iCol
                '    TotalSeries += 1
                'End If

                'If Not IsIDScan Then
                '    'drawing ARUL series
                '    If dblARUL <> SpectNothing Then
                '        ChartTrend.Data(TotalSeries, iCol) = dblARUL
                '        ChartTrend.Points(TotalSeries, iCol).Text = dblARUL

                '        ' ChartTrend.Data(k, iCol) = dblValue
                '        ChartTrend.Data.X(TotalSeries, iCol) = iCol
                '        '.Points(TotalSeries, iCol).Color = Color.Red
                '        TotalSeries += 1
                '    End If
                '    'drawing ARLL series
                '    If dblARLL <> SpectNothing Then
                '        ChartTrend.Data(TotalSeries, iCol) = dblARLL
                '        ChartTrend.Points(TotalSeries, iCol).Text = dblARLL

                '        ' ChartTrend.Data(k, iCol) = dblValue
                '        ChartTrend.Data.X(TotalSeries, iCol) = iCol
                '        '.Points(TotalSeries, iCol).Color = Color.Red
                '        TotalSeries += 1
                '    End If
                'End If
                seritrendline = 5 'TotalSeries
                For k = seritrendline + 1 To (arrParam.Count + seritrendline)    '- row in chart
                    sColName = arrParam(k - seritrendline - 1)
                    If Not IsDBNull(dtDrawTrend.Rows(iCol).Item(sColName)) Then

                        dblValue = dtDrawTrend.Rows(iCol).Item(sColName)
                        ChartTrend.Data(k, iCol) = dblValue
                        ChartTrend.Data.X(k, iCol) = iCol + 1

                        ' ''    '- Find max,min
                        If dblValue < rmin Then
                            rmin = dblValue
                        End If
                        If dblValue > rmax Then
                            rmax = dblValue
                        End If

                        'If dblLSL <> dblUSL Then
                        If dblLSL <> SpectNothing Then
                            If dblValue < dblLSL Then
                                ChartTrend.Points(k, iCol).Color = Color.Red
                            End If
                            If dblARLL <> SpectNothing Then
                                If dblValue < dblARLL And dblValue >= dblLSL Then
                                    ChartTrend.Points(k, iCol).Color = Color.Orange
                                End If
                            End If
                        Else
                            If dblARLL <> SpectNothing Then
                                If dblValue < dblARLL Then
                                    ChartTrend.Points(k, iCol).Color = Color.Orange
                                End If
                            End If
                        End If

                        If dblUSL <> SpectNothing Then
                            If dblValue > dblUSL Then
                                ChartTrend.Points(k, iCol).Color = Color.Red
                            End If
                            If dblARUL <> SpectNothing Then
                                If dblValue > dblARUL And dblValue <= dblUSL Then
                                    ChartTrend.Points(k, iCol).Color = Color.Orange
                                End If

                            End If
                        Else
                            If dblARUL <> SpectNothing Then
                                If dblValue > dblARUL Then
                                    ChartTrend.Points(k, iCol).Color = Color.Orange
                                End If

                            End If
                        End If
                    Else
                        ChartTrend.Data(k, iCol) = Chart.Hidden
                        ChartTrend.Data.X(k, iCol) = iCol + 1
                    End If
                Next
                If dblUSL <> SpectNothing Then

                    ChartTrend.Series(0).Text = "USL = " & dblUSL
                    ChartTrend.Series(0).MarkerShape = MarkerShape.None
                    ChartTrend.Series(0).Line.Width = 2
                    ChartTrend.Series(0).Gallery = Gallery.Lines
                    ChartTrend.Series(0).Color = Color.Red
                    ' TotalSeries += 1
                    'Else
                    '    .Data(0, dtDrawTrend.Rows.Count) = Chart.Hidden
                End If
                'drawing target series
                If dblTarget <> SpectNothing Then

                    ChartTrend.Series(1).Text = "Target = " & dblTarget
                    ChartTrend.Series(1).MarkerShape = MarkerShape.None
                    ChartTrend.Series(1).Line.Width = 1
                    ChartTrend.Series(1).Gallery = Gallery.Lines
                    ChartTrend.Series(1).Color = Color.Cyan
                    'TotalSeries += 1
                End If

                'drawing LSL series
                If dblLSL <> SpectNothing Then

                    ChartTrend.Series(2).Text = "LSL = " & dblLSL
                    ChartTrend.Series(2).MarkerShape = MarkerShape.None
                    ChartTrend.Series(2).Line.Width = 2
                    ChartTrend.Series(2).Gallery = Gallery.Lines
                    ChartTrend.Series(2).Color = Color.Red
                    'TotalSeries += 1
                End If

                'drawing ARUL series
                If dblARUL <> SpectNothing Then


                    ChartTrend.Series(3).Text = "ARUL = " & dblARUL
                    ChartTrend.Series(3).MarkerShape = MarkerShape.None
                    ChartTrend.Series(3).Line.Width = 1.5
                    ChartTrend.Series(3).Gallery = Gallery.Lines
                    ChartTrend.Series(3).Color = Color.Orange
                    'TotalSeries += 1
                    'Else
                    '    .Data(1, dtDrawTrend.Rows.Count) = Chart.Hidden
                End If
                'drawing ARLL series
                If dblARLL <> SpectNothing Then

                    ChartTrend.Series(4).Text = "ARLL = " & dblARLL
                    ChartTrend.Series(4).MarkerShape = MarkerShape.None
                    ChartTrend.Series(4).Line.Width = 1.5
                    ChartTrend.Series(4).Gallery = Gallery.Lines
                    ChartTrend.Series(4).Color = Color.Orange
                    ' TotalSeries += 1
                End If

            Next
            ' Don't visible legend on trend
            'Format serial visible is line
            'TotalSeries = 0
            If dblUSL <> SpectNothing Then
                ChartTrend.Data(0, dtDrawTrend.Rows.Count) = dblUSL
                ' ChartTrend.Data(k, iCol) = dblValue
                ChartTrend.Data.X(0, dtDrawTrend.Rows.Count) = dtDrawTrend.Rows.Count
                ChartTrend.Points(0, dtDrawTrend.Rows.Count).Color = Color.Red
                ' ChartTrend.Series(TotalSeries).Text = "USL = " & dblUSL
                'ChartTrend.Series(TotalSeries).MarkerShape = MarkerShape.None
                'ChartTrend.Series(TotalSeries).Line.Width = 2
                'ChartTrend.Series(TotalSeries).Gallery = Gallery.Lines
                'ChartTrend.Series(TotalSeries).Color = Color.Red

                'TotalSeries += 1
            End If
            'drawing LSL series
            If dblLSL <> SpectNothing Then
                ChartTrend.Data(2, dtDrawTrend.Rows.Count) = dblLSL
                ' ChartTrend.Data(k, iCol) = dblValue
                ChartTrend.Data.X(2, dtDrawTrend.Rows.Count) = dtDrawTrend.Rows.Count
                ChartTrend.Points(2, dtDrawTrend.Rows.Count).Color = Color.Red
                ' ChartTrend.Series(TotalSeries).Text = "LSL = " & dblLSL
                'ChartTrend.Series(TotalSeries).MarkerShape = MarkerShape.None
                'ChartTrend.Series(TotalSeries).Line.Width = 2
                'ChartTrend.Series(TotalSeries).Gallery = Gallery.Lines
                'ChartTrend.Series(TotalSeries).Color = Color.Red
                'TotalSeries += 1
            End If


            'drawing target series
            If dblTarget <> SpectNothing Then
                ChartTrend.Data(1, dtDrawTrend.Rows.Count) = dblTarget
                ChartTrend.Data.X(1, dtDrawTrend.Rows.Count) = dtDrawTrend.Rows.Count
                ChartTrend.Points(1, dtDrawTrend.Rows.Count).Color = Color.Cyan

                'ChartTrend.Series(TotalSeries).MarkerShape = MarkerShape.None
                'ChartTrend.Series(TotalSeries).Line.Width = 1
                'ChartTrend.Series(TotalSeries).Gallery = Gallery.Lines
                'ChartTrend.Series(TotalSeries).Color = Color.Cyan
                'TotalSeries += 1
            End If


            If Not IsIDScan Then
                'drawing ARUL series
                If dblARUL <> SpectNothing Then
                    ChartTrend.Data(3, dtDrawTrend.Rows.Count) = dblARUL
                    ' ChartTrend.Data(k, iCol) = dblValue
                    ChartTrend.Data.X(3, dtDrawTrend.Rows.Count) = dtDrawTrend.Rows.Count
                    ChartTrend.Points(3, dtDrawTrend.Rows.Count).Color = Color.Orange

                    ' ChartTrend.Series(TotalSeries).Text = "ARUL = " & dblARUL
                    'ChartTrend.Series(TotalSeries).MarkerShape = MarkerShape.None
                    'ChartTrend.Series(TotalSeries).Line.Width = 1
                    'ChartTrend.Series(TotalSeries).Gallery = Gallery.Lines
                    'ChartTrend.Series(TotalSeries).Color = Color.Red

                    'TotalSeries += 1
                End If
                'drawing ARLL series
                If dblARLL <> SpectNothing Then
                    ChartTrend.Data(4, dtDrawTrend.Rows.Count) = dblARLL
                    ' ChartTrend.Data(k, iCol) = dblValue
                    ChartTrend.Data.X(4, dtDrawTrend.Rows.Count) = dtDrawTrend.Rows.Count
                    ChartTrend.Points(4, dtDrawTrend.Rows.Count).Color = Color.Orange
                    'ChartTrend.Series(TotalSeries).Text = "ARLL = " & dblARLL
                    'ChartTrend.Series(TotalSeries).MarkerShape = MarkerShape.None
                    'ChartTrend.Series(TotalSeries).Line.Width = 1
                    'ChartTrend.Series(TotalSeries).Gallery = Gallery.Lines
                    'ChartTrend.Series(TotalSeries).Color = Color.Red
                    'TotalSeries += 1
                End If
                '' ''------------------------------------------------------------------
            End If


            'Drawing trend line for ALL parmas

            iparamCount = arrParam.Count
            Select Case iCount
                Case iCount <= 30
                    iCaseAll = 0
                Case 31 To 200
                    'Count of Slipno >=31 and <= 200 - drawing trend: 5 into 1
                    iCaseAll = 6
                Case 201 To 1000
                    'Count of Slipno >=201 and <= 1000 - drawing trend: 20 into 1
                    iCaseAll = 20
                Case Is >= 1001
                    'Count of >= 1001 - drawing trend: 30 into 1
                    iCaseAll = 30
            End Select
            Dim checkpointnull As Boolean = False ' Not have point null
            Dim iparamcountnull As Integer = 0 ' With case have point null when iparamcountnull = arrparam.count - 1
            Dim icountpointnotnull As Integer
            Dim icountpointnull As Integer '= iCaseAll

            For iCol As Integer = 0 To dtDrawTrend.Rows.Count - 1
                With ChartTrend
                    iFlag += 1


                    If iFlag = iCaseAll Or iCount - iCol + 1 < iCaseAll Then
                        icountpointnotnull = iCaseAll
                        icountpointnull = iCaseAll
                        For k = seritrendline + 1 To (arrParam.Count + seritrendline)
                            For iCountTrend = iCol - iCaseAll + 1 To iCol

                                SumX += .Data.X(k, iCountTrend)
                                If .Data(k, iCountTrend) <> Chart.Hidden Then
                                    SumY += .Data(k, iCountTrend)
                                    checkpointnull = False
                                    icountpointnotnull = icountpointnotnull + 1
                                Else
                                    checkpointnull = True
                                    iparamcountnull = arrParam.Count - 1
                                    icountpointnull = icountpointnull + 1
                                End If
                            Next
                        Next
                        XAvg = SumX / (iCaseAll * iparamCount)
                        'YAvg = SumY / (iCaseAll * iparamCount)
                        If checkpointnull = True Then
                            YAvg = SumY / (iCaseAll * iparamcountnull)
                        Else
                            If icountpointnull > iCaseAll Then
                                YAvg = SumY / (icountpointnull * iparamcountnull) + SumY / (icountpointnotnull * iparamCount)
                            Else
                                YAvg = SumY / (iCaseAll * iparamCount)
                            End If
                        End If

                        SumX = 0
                        SumY = 0
                        For iCountTrend = iCol - iCaseAll + 1 To iCol

                            .Data(seritrendline, iCountTrend) = YAvg
                            .Data.X(seritrendline, iCountTrend) = XAvg
                        Next

                        iFlag = 0
                    End If

                End With

            Next
            ChartTrend.Series(seritrendline).MarkerSize = 0
            ChartTrend.Series(seritrendline).Line.Width = 4
            ChartTrend.Series(seritrendline).Gallery = Gallery.Lines
            ' ''--------------------------------------------------------------------
            ChartTrend.AxisY.CustomGridLines.Clear()
            ''- Draw line date
            strOldDate = ""
            For iCol As Integer = 0 To dtDrawTrend.Rows.Count - 1   ' iColNo   '- column in chart = row in datatable
                If strOldDate <> dtDrawTrend.Rows(iCol).Item("PDDATE").ToString Then
                    constSpect = New ChartFX.WinForms.CustomGridLine
                    ChartTrend.AxisX.CustomGridLines.Add(constSpect)
                    strOldDate = dtDrawTrend.Rows(iCol).Item("PDDATE")
                    strOldShift = dtDrawTrend.Rows(iCol).Item(strShiftType)
                    With constSpect
                        '.Axis = SoftwareFX.ChartFX.AxisItem.X
                        If iCol = 0 Then
                            .Value = iCol + 0.5
                        Else
                            .Value = iCol + 0.07
                        End If
                        If dtDrawTrend.Rows(iCol).Item(strShiftType) = "DAY" Then
                            .Text = strOldDate & "-D"
                            .TextColor = System.Drawing.Color.Orchid
                            .Color = System.Drawing.Color.Magenta
                        Else    ' NIGHT
                            .Text = strOldDate & "-N"
                            .TextColor = System.Drawing.Color.Black
                            .Color = System.Drawing.Color.Black
                        End If
                        .Alignment = StringAlignment.Far
                        .Width = 1
                        .Style = System.Drawing.Drawing2D.DashStyle.Dash
                    End With
                    iLine += 1
                Else
                    If strOldShift <> dtDrawTrend.Rows(iCol).Item(strShiftType) Then
                        'constSpect = ChartTrend.AxisX.CustomGridLines(iLine)
                        constSpect = New ChartFX.WinForms.CustomGridLine
                        ChartTrend.AxisX.CustomGridLines.Add(constSpect)
                        strOldShift = dtDrawTrend.Rows(iCol).Item(strShiftType)
                        With constSpect
                            '.Axis = SoftwareFX.ChartFX.AxisItem.X
                            .Color = System.Drawing.Color.Magenta
                            If iCol = 0 Then
                                .Value = iCol + 0.5
                            Else
                                .Value = iCol + 0.07
                            End If
                            If strOldShift = "DAY" Then
                                .Text = strOldDate & "-D"
                                .TextColor = System.Drawing.Color.Orchid
                                .Color = System.Drawing.Color.Magenta
                            Else    ' NIGHT
                                .Text = strOldDate & "-N"
                                .TextColor = System.Drawing.Color.Black
                                .Color = System.Drawing.Color.Black
                            End If
                            .Alignment = StringAlignment.Far
                            .Width = 1
                            .Style = System.Drawing.Drawing2D.DashStyle.Dash
                        End With
                        iLine += 1
                    End If
                End If
            Next
            'ChartTrend.CloseData(COD.Values)

            '--------------Begin Modify by Dao Thi Hai 09/11/2007-----------------------------------------
            If strColumnName.ToUpper = "ALL" Then
                If dblLSL <> SpectNothing Then
                    If dblLSL - rmin > 0 Then
                        rDown = rmin - (rmax - rmin) / 2
                    Else
                        rDown = dblLSL - (rmax - rmin) / 2
                    End If
                Else
                    rDown = rmin - (rmax - rmin) / 2
                End If
            Else
                If dblLSL <> SpectNothing Then
                    If rmin > dblLSL Then

                        rDown = dblLSL - (rmax - rmin) / 2
                    Else
                        rDown = rmin - (rmax - rmin) / 2
                    End If
                Else
                    rDown = rmin - (rmax - rmin) / 2
                End If

            End If


            If strColumnName.ToUpper = "ALL" Then
                If dblUSL <> SpectNothing Then
                    If rmax > dblUSL Then
                        rUp = rmax + (rmax - rmin) / 2
                    Else
                        rUp = dblUSL + (rmax - rmin) / 2
                    End If
                Else
                    rUp = rmax + (rmax - rmin) / 2

                End If


            Else
                If dblUSL <> SpectNothing Then
                    If rmax < dblUSL Then
                        rUp = dblUSL + (rmax - rmin) / 2
                    Else
                        rUp = rmax + (rmax - rmin) / 2

                    End If
                Else
                    rUp = rmax + (rmax - rmin) / 2
                End If
            End If


            ' '' '''- Draw Spect
            ' ''If dblTarget <> SpectNothing Then
            ' ''    Dim constTarget As ChartFX.WinForms.CustomGridLine
            ' ''    constTarget = New ChartFX.WinForms.CustomGridLine
            ' ''    ChartTrend.AxisY.CustomGridLines.Add(constTarget)

            ' ''    'constTarget = ChartTrend.AxisY.CustomGridLines(iLine)
            ' ''    With constTarget
            ' ''        .Value = dblTarget
            ' ''        .Color = System.Drawing.Color.Cyan

            ' ''        '.Axis = SoftwareFX.ChartFX.AxisItem.Y
            ' ''        .Text = ""
            ' ''        .Width = 1
            ' ''        .Style = System.Drawing.Drawing2D.DashStyle.Solid
            ' ''    End With
            ' ''    iLine += 1
            ' ''End If
            '' ''- LCL : UCL
            ' ''If dblARLL <> dblARUL Then
            ' ''    If dblARLL <> SpectNothing Then
            ' ''        Dim constSpect1 As ChartFX.WinForms.CustomGridLine
            ' ''        constSpect1 = New ChartFX.WinForms.CustomGridLine
            ' ''        ChartTrend.AxisY.CustomGridLines.Add(constSpect1)
            ' ''        'constSpect1 = ChartTrend.AxisY.CustomGridLines(iLine)
            ' ''        iLine += 1
            ' ''        With constSpect1
            ' ''            .Value = dblARLL
            ' ''            .Color = System.Drawing.Color.OrangeRed
            ' ''            '.Axis = SoftwareFX.ChartFX.AxisItem.Y
            ' ''            .Text = "ARLL = " & dblARLL   'Space(10) &
            ' ''            .TextColor = System.Drawing.Color.OrangeRed
            ' ''            .Width = 1
            ' ''            .Style = System.Drawing.Drawing2D.DashStyle.Solid
            ' ''        End With
            ' ''    End If

            ' ''    If dblARUL <> SpectNothing Then
            ' ''        Dim constSpect2 As ChartFX.WinForms.CustomGridLine
            ' ''        'constSpect2 = ChartTrend.AxisY.CustomGridLines(iLine)
            ' ''        constSpect2 = New ChartFX.WinForms.CustomGridLine
            ' ''        ChartTrend.AxisY.CustomGridLines.Add(constSpect2)
            ' ''        iLine += 1
            ' ''        With constSpect2
            ' ''            .Value = dblARUL
            ' ''            .Color = System.Drawing.Color.OrangeRed
            ' ''            '.Axis = SoftwareFX.ChartFX.AxisItem.Y
            ' ''            .Text = "ARUL = " & dblARUL
            ' ''            .TextColor = System.Drawing.Color.OrangeRed
            ' ''            .Width = 1
            ' ''            .Style = System.Drawing.Drawing2D.DashStyle.Solid
            ' ''        End With
            ' ''    End If
            ' ''End If
            '' ''- LSL : USL
            ' ''If dblLSL <> SpectNothing Then
            ' ''    Dim constSpect3 As ChartFX.WinForms.CustomGridLine
            ' ''    'constSpect3 = ChartTrend.AxisY.CustomGridLines(iLine)
            ' ''    constSpect3 = New ChartFX.WinForms.CustomGridLine
            ' ''    ChartTrend.AxisY.CustomGridLines.Add(constSpect3)
            ' ''    iLine += 1
            ' ''    With constSpect3
            ' ''        .Value = dblLSL + (rmax - rmin) / 200
            ' ''        .Color = System.Drawing.Color.Red
            ' ''        '.Axis = SoftwareFX.ChartFX.AxisItem.Y
            ' ''        .Text = "LSL = " & dblLSL
            ' ''        .TextColor = System.Drawing.Color.Red
            ' ''        .Width = 2
            ' ''        .Style = System.Drawing.Drawing2D.DashStyle.Solid

            ' ''    End With
            ' ''End If
            '' ''- USL
            ' ''If dblUSL <> SpectNothing Then
            ' ''    Dim constSpect4 As ChartFX.WinForms.CustomGridLine
            ' ''    'constSpect4 = ChartTrend.AxisY.CustomGridLines(iLine)
            ' ''    constSpect4 = New ChartFX.WinForms.CustomGridLine
            ' ''    ChartTrend.AxisY.CustomGridLines.Add(constSpect4)
            ' ''    iLine += 1
            ' ''    With constSpect4
            ' ''        .Value = dblUSL
            ' ''        .Color = System.Drawing.Color.Red
            ' ''        '.Axis = SoftwareFX.ChartFX.AxisItem.Y
            ' ''        .Text = "USL = " & dblUSL
            ' ''        .TextColor = System.Drawing.Color.Red
            ' ''        .Width = 2
            ' ''        .Style = System.Drawing.Drawing2D.DashStyle.Solid
            ' ''    End With
            ' ''End If
            '' ''End If
            '- Set chart
            With ChartTrend
                .Titles.Clear()
                .LegendBox.Titles.Clear()
                Dim title, title4 As ChartFX.WinForms.TitleDockable
                title = New ChartFX.WinForms.TitleDockable

                title.Text = OperationName.ToUpper & " -- " & ParameterName
                title.Font = New System.Drawing.Font("Times new Roman", 7, System.Drawing.FontStyle.Bold)
                title.TextColor = Color.DarkBlue
                .Titles.Add(title)

                title4 = New ChartFX.WinForms.TitleDockable

                title4.Text = "From: " & Today.AddDays(-7).ToString("yyyyMMdd") & " To: " & Today.ToString("yyyyMMdd") & " ***** Created: " & Today.ToString("yyyyMMdd") & "--" & Now.ToLongTimeString
                title4.Font = New System.Drawing.Font("Times new Roman", 7, System.Drawing.FontStyle.Bold)
                title4.TextColor = Color.DarkBlue
                .Titles.Add(title4)


                .Gallery = Gallery.Curve
                .LegendBox.Visible = True
                .LegendBox.Font = New System.Drawing.Font("Arial", 7.0!)
                .LegendBox.Style = LegendBoxStyles.Wordbreak
                .LegendBox.PlotAreaOnly = False


                .LegendBox.ItemAttributes(.AxisY.CustomGridLines).Visible = False
                .LegendBox.ItemAttributes(.AxisX.CustomGridLines).Visible = False
                .LegendBox.ContentLayout = ContentLayout.Near
                .LegendBox.Border = DockBorder.External
                '.SerLegBoxObj.AutoSize = True
                '.LegendBox.Width = 100
                '.SerLegBoxObj.SizeToFit()
                '.SerLegBoxObj.AutoSize = True
                Dim title1, title2, title3, title5 As ChartFX.WinForms.TitleDockable
                title1 = New ChartFX.WinForms.TitleDockable
                If options = 0 Then
                    title1.Text = "Item: " & item_grpItem
                Else
                    title1.Text = "GroupItem: " & item_grpItem
                End If

                title1.Dock = DockArea.Top
                .LegendBox.Titles.Add(title1)
                title2 = New ChartFX.WinForms.TitleDockable
                title2.Text = "FDate: " & Today.AddDays(-7).ToString("yyyyMMdd")
                .LegendBox.Titles.Add(title2)

                title3 = New ChartFX.WinForms.TitleDockable
                title3.Text = "TDate: " & Today.ToString("yyyyMMdd")
                .LegendBox.Titles.Add(title3)
                ' .SerLegBoxObj.Titles(2).Alignment = StringAlignment.Near
                title5 = New ChartFX.WinForms.TitleDockable
                title5.Text = "**************"
                .LegendBox.Titles.Add(title5)

                .AxisX.Visible = True
                .AxisX.Style = ChartTrend.AxisX.Style
                .AxisY.Max = rUp
                .AxisY.Min = rDown
                .AxisY.DataFormat.Decimals = 4
                .AxisY.LabelsFormat.Decimals = 4

                If strColumnName = "ALL" Then
                    If arrParam.Count > 4 Then
                        For k = seritrendline + 1 To arrParam.Count + seritrendline     '- row in chart
                            .Series(k).Text = arrParam(k - seritrendline - 1)
                            '.Series(k).MarkerShape = MarkerShape.Diamond
                            .Series(k).MarkerSize = 2
                        Next
                    Else
                        For k = seritrendline + 1 To arrParam.Count + seritrendline  '- row in chart
                            .Series(k).Text = arrParam(k - seritrendline - 1)
                            '.Series(k).MarkerShape = MarkerShape.Diamond
                            .Series(k).MarkerSize = 2
                            ' ''If k = 4 Then
                            ' ''    .Series(k).Color = ColorTranslator.FromHtml("#5179d6") 'Color.Red
                            ' ''End If
                            ' ''If k = 5 Then
                            ' ''    .Series(k).Color = ColorTranslator.FromHtml("#66cc66") 'Color.Hex() = {33, 99, FF}
                            ' ''    'Color.RoyalBlue
                            ' ''End If
                            '' ''If k = 1 Then
                            '' ''    .Series(k).Color = Color.Green
                            '' ''End If
                            '' ''If k = 2 Then
                            '' ''    .Series(k).Color = Color.RoyalBlue
                            '' ''End If
                            ' ''If k = 6 Then
                            ' ''    .Series(k).Color = Color.Magenta
                            ' ''End If
                            ' ''If k = 7 Then
                            ' ''    .Series(k).Color = Color.LightSteelBlue
                            ' ''End If
                        Next
                        ''add by maidt.it
                        If arrParam.Count = 2 Then
                            .Series(seritrendline + 1).Color = ColorTranslator.FromHtml("#5179d6")
                            .Series(seritrendline + 2).Color = ColorTranslator.FromHtml("#66cc66")
                        End If
                        If arrParam.Count = 3 Then
                            .Series(seritrendline + 1).Color = ColorTranslator.FromHtml("#5179d6")
                            .Series(seritrendline + 2).Color = ColorTranslator.FromHtml("#66cc66")
                            .Series(seritrendline + 3).Color = Color.Magenta
                        End If
                        If arrParam.Count = 4 Then
                            .Series(seritrendline + 1).Color = ColorTranslator.FromHtml("#5179d6")
                            .Series(seritrendline + 2).Color = ColorTranslator.FromHtml("#66cc66")
                            .Series(seritrendline + 3).Color = Color.Magenta
                            .Series(seritrendline + 4).Color = Color.LightSteelBlue
                        End If
                    End If
                    'If arrParam.Count * 3 > 30 Then
                    If arrParam.Count = 1 Then
                        .Series(seritrendline + 1).Color = Color.MediumAquamarine
                        .Series(seritrendline + 1).MarkerShape = MarkerShape.Circle
                    End If

                    If dtDrawTrend.Rows.Count > 30 Then
                        .Series(seritrendline).Color = Color.Blue
                        .Series(seritrendline).Gallery = Gallery.Lines
                        .Series(seritrendline).Text = "Trend line"
                        .Series(seritrendline).MarkerShape = MarkerShape.None
                    End If
                Else
                    ' ''.Series(1).Text = strColumnName
                    ' ''.Series(1).MarkerShape = MarkerShape.Diamond
                    '.Series(0).MarkerSize = 2
                    If dtDrawTrend.Rows.Count > 30 Then
                        .Series(seritrendline).Text = "Trend line"
                        .Series(seritrendline).MarkerShape = MarkerShape.None
                        .Series(seritrendline).Color = Color.Blue
                        .Series(seritrendline).Gallery = Gallery.Lines
                    End If
                    '.SerLegBoxObj.BackColor = Color.Yellow
                End If


                If strColumnName.ToUpper <> "ALL" Then
                    .Series(seritrendline + 1).Color = Color.MediumSeaGreen
                End If
                .LegendBox.ContentLayout = ContentLayout.Near
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Sub DrawBar(ByVal ChartTrend As ChartFX.WinForms.Chart, ByVal dtDrawTrend As DataTable, ByVal strParamvaluelst As String, ByVal Item_GrpItem As String, ByVal options As String, ByVal ParameterName As String, ByVal strColorlist As String)

        Dim seriHashtable As Hashtable = New Hashtable
        Dim valueRatioHashtable As Hashtable = New Hashtable
        Dim strPointLabel As String
        Dim iPoint, iSerie, iMaxSerie As Integer
        Dim dblMaxYValue, dblTotalYValueByPoint As Integer
        Dim dbMaxY2, dbTotalValueByPercent As Double
        Dim colDrawBy As String
        Dim i As Integer
        Dim TotalParamValue As Integer
        Dim lstValue, tempValue As String
        Dim arr() As String
        Dim dt As DataTable
        Dim paramcol(0) As String
        Dim dr() As DataRow
        Dim arrcolor() As String
        Dim arrDefaultvalue() As String
        Dim strExpr As String = "paramvalue  in (" & strParamvaluelst.ToUpper() & ")"
        Dim strSort As String = "DATESHIFT ASC"

        Try
            'dr = dtDrawTrend.Select(" PARAMVALUE in ('" & strParamvaluelst.ToUpper() & "')", "DATESHIFT ASC")
            dr = dtDrawTrend.Select(" PARAMVALUE is not null ", "DATESHIFT ASC")

            If dr Is Nothing Then
                Exit Sub
            End If
            If dr.Length > 0 Then
                seriHashtable.Clear()
                valueRatioHashtable.Clear()
                'Array.Clear(arr, 0, arr.Length)
                paramcol(0) = "paramvalue"
                dt = dtDrawTrend.DefaultView.ToTable(True, paramcol)
                If Not IsDBNull(dt.Rows(0).Item("paramvalue")) Then
                    lstValue = dt.Rows(0).Item("paramvalue")
                    tempValue = dt.Rows(0).Item("paramvalue")
                End If

                For i = 0 To dt.Rows.Count - 1
                    If Not IsDBNull(dt.Rows(i).Item("paramvalue")) Then
                        If dt.Rows(i).Item("paramvalue") <> tempValue Then
                            tempValue = dt.Rows(i).Item("paramvalue")
                            lstValue &= "|" & tempValue
                        End If
                    End If
                Next

                lstValue = lstValue.ToUpper()
                arr = lstValue.Split("|")


                'arr = lstValue.Split("|")
                Array.Sort(arr)
                TotalParamValue = arr.Length
                For i = 0 To arr.Length - 1
                    seriHashtable.Add(arr(i), i)
                    valueRatioHashtable.Add(arr(i), 0)
                Next
                'ChartTrend = New ChartFX.WinForms.Chart

                If strColorlist <> "" Then
                    arrDefaultvalue = strParamvaluelst.Replace("'", "").Split(",")
                    arrcolor = strColorlist.Split("|")
                End If

                With ChartTrend
                    .Reset()

                    .Data.Clear()
                    .Titles.Clear()
                    .Extensions.Clear()

                    .AxisX.Sections.Clear()
                    .AxisX.Labels.Clear()
                    .AxisY2.Sections.Clear()
                    .LegendBox.CustomItems.Clear()
                    .LegendBox.Titles.Clear()

                    .AxisY.CustomGridLines.Clear()

                    .Series.Clear()
                    .Points.Clear()

                    .Data.Labels.Clear()

                    ChartTrend.Height = 330 '312
                    ChartTrend.Width = 600

                    .Gallery = ChartFX.WinForms.Gallery.Bar
                    .AxisY.LabelsFormat.Format = AxisFormat.Percentage
                    .AxisY.LabelsFormat.Decimals = 2
                    .AllSeries.Stacked = ChartFX.WinForms.Stacked.Normal

                    strPointLabel = ""
                    iPoint = -1
                    'If ChkNotNumeric.Checked Then

                    colDrawBy = "DATESHIFT"

                    ' Drawing chart for all data for Axis Y1 
                    For Each drow As DataRow In dr
                        If strPointLabel <> drow(colDrawBy).ToString() Then
                            iPoint += 1
                            strPointLabel = drow(colDrawBy).ToString()


                            dblTotalYValueByPoint = CType(drow("TOTAL_PCS"), Integer)
                            dbTotalValueByPercent = CType(drow("PCS_PERCENT"), Double)
                        Else
                            dblTotalYValueByPoint = CType(drow("TOTAL_PCS"), Integer)
                            dbTotalValueByPercent += CType(drow("PCS_PERCENT"), Double)
                        End If


                        'Luu lai gia tri pcs cho tung value (vi du: value A co 10 pcs ...)
                        ' valueRatioHashtable(drow("paramvalue").ToString().Trim()) = drow("pcs_percent")
                        ' valueRatioHashtable(drow("paramvalue").ToString().Trim()) = drow("TOTAL_PCS")

                        iSerie = CType(seriHashtable(drow("paramvalue").ToString().Trim()), Integer)
                        .Data.Y(iSerie, iPoint) = CType(drow("pcs_percent"), Double)


                        .AxisX.Labels(iPoint) = strPointLabel '.Substring(4, 5)

                        'define the max y value for all chart points
                        If dblMaxYValue < dbTotalValueByPercent Then
                            dblMaxYValue = dbTotalValueByPercent
                        End If

                        If dbMaxY2 < dblTotalYValueByPoint Then
                            dbMaxY2 = dblTotalYValueByPoint
                        End If
                    Next


                    ' Show Series (Label) in botton Corner
                    For i = 0 To TotalParamValue - 1
                        .Series(i).Text = arr(i) 'Show Series
                        .Series(i).Stacked = True
                        If strColorlist <> "" Then
                            For j As Integer = 0 To arrDefaultvalue.Length - 1
                                If arrDefaultvalue(j) = arr(i) Then
                                    .Series(i).Color = Color.FromArgb(CType(arrcolor(j), Integer))
                                    Exit For
                                End If
                            Next
                            ' .Series(i).Color = Color.FromArgb(CType(arrcolor(i), Integer))
                        End If
                    Next

                    ' Drawing chart for all data for Axis Y2 
                    iPoint = -1
                    strPointLabel = ""
                    For Each drow As DataRow In dr
                        If strPointLabel <> drow(colDrawBy).ToString() Then
                            iPoint += 1
                            strPointLabel = drow(colDrawBy).ToString()
                            dblTotalYValueByPoint = CType(drow("TOTAL_PCS"), Integer)
                        Else
                            dblTotalYValueByPoint = CType(drow("TOTAL_PCS"), Integer)
                        End If
                        .Data.Y(TotalParamValue + 1, iPoint) = CType(drow("TOTAL_PCS"), Integer) ' .Data.Y(SeriesIndex, pointIndex) =Value
                    Next

                    '  show lalbel in botton corner
                    .AxisX.LabelAngle = 45
                    .AxisX.Step = 1
                    .Series(TotalParamValue + 1).Gallery = Gallery.Curve
                    .Series(TotalParamValue + 1).AxisY = .AxisY2   'Gan toa do ve theo truc Y2
                    .Series(TotalParamValue + 1).Text = "Total pcs"      '  Set Series for Y2
                    'For i = 0 To TotalParamValue - 1
                    '    .Series(i).Stacked = True
                    '    .Series(i).Text = arr(i)
                    'Next

                    ''For i = TotalParamValue To 2 * TotalParamValue - 1
                    ''    .Series(i).Gallery = Gallery.Curve
                    ''    .Series(i).AxisY = .AxisY2
                    ''    .Series(i).Text = "% " & arr(i - TotalParamValue)
                    ''    .Series(i).MarkerSize = 3
                    ''Next

                    .AxisY2.Visible = True

                    .AxisY2.Visible = True
                    .AxisY2.LabelsFormat.Format = AxisFormat.Number

                    .AxisY2.AutoScale = False
                    .AxisY2.Min = 0
                    .AxisY2.Max = dbMaxY2 + 50

                    .AxisY2.Grids.Minor.Visible = True
                    .AxisY2.Grids.Interlaced = True
                    .AxisY2.Grids.InterlacedColor = Color.Ivory


                    .AxisY.Title.Text = "Ratio"
                    .AxisY.Title.Font = New System.Drawing.Font("Times new Roman", 10, System.Drawing.FontStyle.Bold)
                    .AxisY.Max = 1
                    .AxisY.Min = 0
                    .AxisY.Step = 0.2


                    .AxisY2.Title.Text = "Total pcs"
                    .AxisY2.Title.Font = New System.Drawing.Font("Times new Roman", 10, System.Drawing.FontStyle.Bold)
                    Dim title As ChartFX.WinForms.TitleDockable
                    title = New ChartFX.WinForms.TitleDockable
                    If options = 0 Then
                        title.Text = ParameterName & " *** Item:" & Item_GrpItem
                    Else
                        title.Text = ParameterName & " *** GrpItem:" & Item_GrpItem
                    End If

                    title.Font = New System.Drawing.Font("Times new Roman", 8, System.Drawing.FontStyle.Bold)
                    title.TextColor = Color.DarkBlue
                    .Titles.Add(title)
                    Dim title4 As ChartFX.WinForms.TitleDockable
                    title4 = New ChartFX.WinForms.TitleDockable

                    title4.Text = "From: " & Today.AddDays(-7).ToString("yyyyMMdd") & " To: " & Today.ToString("yyyyMMdd") & " ***** Created: " & Today.ToString("yyyyMMdd") & "--" & Now.ToLongTimeString
                    title4.Font = New System.Drawing.Font("Times new Roman", 7, System.Drawing.FontStyle.Bold)
                    title4.TextColor = Color.DarkBlue
                    .Titles.Add(title4)

                    ' ''Dim title1, title2, title3, title5 As ChartFX.WinForms.TitleDockable
                    ' ''title1 = New ChartFX.WinForms.TitleDockable
                    ' ''title1.Text = "Item: " & sShippingItem
                    ' ''title1.Dock = DockArea.Top
                    ' ''title1.Font = New System.Drawing.Font("Times new Roman", 5, System.Drawing.FontStyle.Regular)
                    ' ''.LegendBox.Titles.Add(title1)
                    ' ''title2 = New ChartFX.WinForms.TitleDockable
                    ' ''title2.Text = "FDate: " & Today.AddDays(-7).ToString("yyyyMMdd")
                    ' ''title2.Font = New System.Drawing.Font("Times new Roman", 5, System.Drawing.FontStyle.Regular)
                    ' ''.LegendBox.Titles.Add(title2)

                    ' ''title3 = New ChartFX.WinForms.TitleDockable
                    ' ''title3.Text = "TDate: " & Today.ToString("yyyyMMdd")
                    ' ''title3.Font = New System.Drawing.Font("Times new Roman", 5, System.Drawing.FontStyle.Regular)
                    ' ''.LegendBox.Titles.Add(title3)

                    '' '' .SerLegBoxObj.Titles(2).Alignment = StringAlignment.Near
                    ' ''title5 = New ChartFX.WinForms.TitleDockable
                    ' ''title5.Text = "**************"

                    ' ''title5.Font = New System.Drawing.Font("Times new Roman", 5, System.Drawing.FontStyle.Regular)
                    ' ''.LegendBox.Titles.Add(title5)


                    .LegendBox.SizeToFit()
                    .LegendBox.ContentLayout = ContentLayout.Near
                    .LegendBox.Visible = True
                    .LegendBox.Dock = DockArea.Bottom
                    .LegendBox.Border = DockBorder.External
                    .LegendBox.Font = New Font("Times new Roman", 6)


                End With
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadParamValue(ByVal strOperationoid As String, ByVal strParameteroid As String, ByVal parametername As String, ByVal strFacodeOID As String, ByVal options As String, ByVal grpitemoid As String, ByRef listparamvalue As String, ByRef listcolor As String)
        Dim strSQL As String = ""
        Dim ds As New DataSet
        Dim qcstdparamoid, strCon As String
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection
        Dim adapter As OleDbDataAdapter
        Dim strsql1 As String = ""
        '       Dim strDEFAULTVALUE As String = ""
        Try

            strCon = sConnectionString_stb
            con = New OleDbConnection(strCon)
            con.Open()

            'Get QCstdparamoid from reportoid
            If options = 0 Then
                strSQL = "SELECT * FROM mreport_qcstdparam C" & _
                           " WHERE C.MREPORTPARAMOID in (select a.oid from mreportparam a where a.OID ='" & strParameteroid & "') and  rownum = 1"
            Else
                strSQL = "SELECT * FROM mreport_qcstdparam C" & _
                           " WHERE C.MREPORTPARAMOID in (select a.oid from mreportparam a where a.REPORTPARANAME ='" & parametername & "' and a.facode = '" & strFacodeOID & "' " & _
                           " and SHIPPINGOID in (select OID from mshippingitem where grp_oid= '" & grpitemoid & "')" & _
                           ") and  rownum = 1"

            End If


            cmd = New OleDbCommand(strSQL, con)
            cmd.CommandType = CommandType.Text
            adapter = New OleDbDataAdapter(cmd)

            'Dim dsQCPARAMETER As DataSet = New DataSet
            If Not IsNothing(dsQC.Tables("mreport_qcstdparam")) Then
                dsQC.Tables("mreport_qcstdparam").Clear()
            End If
            adapter.Fill(dsQC, "mreport_qcstdparam")

            If dsQC.Tables("mreport_qcstdparam").Rows.Count > 0 Then
                qcstdparamoid = dsQC.Tables("mreport_qcstdparam").Rows(0).Item(2)
                'get combobox value
                strsql1 = "select      DEFAULTVALUE" & _
                        " from (" & _
                        " select DEFAULTVALUE, length(DEFAULTVALUE) as leng" & _
                        " from                       dqcchecksheet       " & _
                        " where mqcstdparamoid= '" & qcstdparamoid & "'" & _
                        " and length(DEFAULTVALUE) is not null" & _
                        " order by length(DEFAULTVALUE) desc" & _
                        " )" & _
                        " where rownum = 1"

                strSQL = " select DEFAULTVALUE,DEFAULTCOLOR from mqcstdparam where oid ='" & qcstdparamoid & "'"

                cmd.CommandText = strsql1
                If Not IsNothing(dsQC.Tables("dqcchecksheet")) Then
                    dsQC.Tables("dqcchecksheet").Clear()
                End If
                adapter.Fill(dsQC, "dqcchecksheet")

                ' truong hop neu chua set trong standardparameter thi lay trong checksheet
                cmd.CommandText = strSQL
                If Not IsNothing(dsQC.Tables("mqcstdparam")) Then
                    dsQC.Tables("mqcstdparam").Clear()
                End If
                adapter.Fill(dsQC, "mqcstdparam")


                If dsQC.Tables("mqcstdparam").Rows.Count > 0 Then
                    If (Not IsDBNull(dsQC.Tables("mqcstdparam").Rows(0).Item(0))) AndAlso (Not IsDBNull(dsQC.Tables("mqcstdparam").Rows(0).Item(1))) Then
                        listparamvalue = dsQC.Tables("mqcstdparam").Rows(0).Item(0)
                        listparamvalue = listparamvalue.Replace("|", "','")
                        'fill default value into combobox
                        'strDEFAULTVALUE = strDEFAULTVALUE.Replace("|", "','")
                        listcolor = dsQC.Tables("mqcstdparam").Rows(0).Item(1)
                    Else ' If not setting in stdparameter then get from dqcchecksheet
                        If dsQC.Tables("dqcchecksheet").Rows.Count > 0 Then
                            listparamvalue = dsQC.Tables("dqcchecksheet").Rows(0).Item(0)
                            listparamvalue = listparamvalue.Replace("|", "','")
                        End If
                    End If
                End If

            End If
            con.Close()
            ' Return strDEFAULTVALUE
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "drawing trend for ShippingItem"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' drawing trend for each Shipping Item
    ''' </summary>
    ''' <param name="ShippingItemOID"></param>
    ''' <param name="dateView"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[dmquyen]	07/26/2006	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    '''For run
    Private Sub DrawingForEachSPItem(ByVal sShippingItem As String, ByVal ShippingItemOID As String, ByVal dateView As String, ByVal options As Integer, tablename As String)
        Dim drOperation, drParameter As DataRow
        Dim bGenerated As Boolean
        Dim FacodeA, FacodeB, FacodeV2 As String
        Try
            Trace.WriteLineIf(TRACE_ENABLE, " --- Load Section AT :" & Now)

            LoadSection(ShippingItemOID, options) 'On Hoyav3_stb
            'bGenerated = False
            Trace.WriteLineIf(TRACE_ENABLE, " --- Check if generated :" & Now)
            bGenerated = CheckGenerated(sShippingItem, Today.ToString("yyyyMMdd"), options) 'On Hoyav3
            Trace.WriteLineIf(TRACE_ENABLE, " --- Start Loop on Operation :" & Now)
            'For Each drOperation In dsQC.Tables("OPERATION").Rows
            'get parameter for all sections
            LoadParameter(ShippingItemOID, options) 'On Hoyav3_stb
            Trace.WriteLineIf(TRACE_ENABLE, " ------ Start Loop on Parameter :" & Now)
            FacodeV2 = ""
            FacodeA = ""
            FacodeB = ""
            For Each drParameter In dsQC.Tables("QCPARAMETER").Rows
                'If drParameter.Item("reportparaname") = "CSCL Dub off" Then
                Select Case drParameter.Item("Facode")
                    Case "1"
                        FacodeV2 = "1"
                    Case "6"
                        FacodeA = "6"
                    Case "8"
                        FacodeB = "8"
                End Select
                If drParameter("reportparaname").ToString() = "BLANK AUTO INSPECTION DEFECT" Then
                    FacodeA = "6"
                End If

                ' ActualFacode = drParameter.Item("facode")
                LoadInfoParameter(drParameter.Item("OID"), drParameter("reportparaname").ToString(), options, ShippingItemOID, drParameter.Item("OPERATIONOID")) 'On Hoyav3_stb
                Trace.WriteLineIf(TRACE_ENABLE, " --------- Start QCEmail.InsertImageIntoTableResult on Parameter :" & drParameter("reportparaname").ToString() & " AT " & Now)

                InsertImageIntoTableResult(ShippingItemOID, sShippingItem, drParameter.Item("OPERATIONOID"), drParameter.Item("OPERATIONNAME"), drParameter.Item("OID"), drParameter("reportparaname").ToString(), dateView, bGenerated, drParameter.Item("Facode"), options, tablename) 'On Hoyav3

                Trace.WriteLineIf(TRACE_ENABLE, " --------- End QCEmail.InsertImageIntoTableResult on Parameter :" & drParameter("reportparaname").ToString() & " AT " & Now)
                'End If
            Next

            Trace.WriteLineIf(TRACE_ENABLE, " ------ End Loop on Parameter :" & Now)
            'Next
            Trace.WriteLineIf(TRACE_ENABLE, " --- End Loop on Operation :" & Now)

        Catch ex As Exception
            Throw New Exception(ex.ToString, ex)
        End Try
    End Sub

    '''For test
    'Private Sub DrawingForEachSPItem(ByVal sShippingItem As String, ByVal dateView As String)
    '    Dim drOperation, drParameter As DataRow
    '    Dim bGenerated As Boolean
    '    Try
    '        LoadSection(ShippingItemOID) 'On Hoyav3_stb        '        bGenerated = False
    '        'bGenerated = CheckGenerated(sShippingItem, Today.ToString("yyyyMMdd")) 'On Hoyav3
    '        LoadParameter("420", ShippingItemOID)
    '        'For Each drParameter In dsQC.Tables("QCPARAMETER").Rows
    '        LoadInfoParameter("6E639801E9B380D6E043AC19062880D6", ShippingItemOID, "420") 'On Hoyav3_stb
    '        InsertImageIntoTableResult(ShippingItemOID, "420", "6E639801E9B380D6E043AC19062880D6", dateView, bGenerated, "6") 'On Hoyav3
    '        'Next
    '        'For Each drOperation In dsQC.Tables("OPERATION").Rows
    '        '    LoadParameter(drOperation.Item("OID"), ShippingItemOID) 'On Hoyav3_stb
    '        '    For Each drParameter In dsQC.Tables("QCPARAMETER").Rows
    '        '        LoadInfoParameter(drParameter.Item("OID"), ShippingItemOID, drOperation.Item("OID")) 'On Hoyav3_stb
    '        '        InsertImageIntoTableResult(ShippingItemOID, drOperation.Item("OID"), drParameter.Item("OID"), dateView, bGenerated, drParameter.Item("Facode")) 'On Hoyav3
    '        '    Next
    '        'Next
    '        'Generate IDScan trend image
    '        'Try
    '        '    LoadInfoIDScanSpec(sShippingItem)
    '        '    InsertImageIntoTableResult(ShippingItemOID, "1000", "IDScan", dateView, bGenerated, "6", True) 'HOGV1
    '        '    InsertImageIntoTableResult(ShippingItemOID, "1000", "IDScan", dateView, bGenerated, "8", True) 'HOGV2

    '        'Catch ex As Exception

    '        'End Try
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub DrawingForEachSPItem(ByVal sShippingItem As String, ByVal dateView As String)
    '    Dim drOperation, drParameter As DataRow
    '    Dim bGenerated As Boolean
    '    Try

    '        LoadShippingItem()
    '        ShippingItemOID = "db14220a-590c-42a3-8" 'dsQC.Tables("ShippingItem").Rows.Find(sShippingItem).Item("OID")
    '        LoadSection(ShippingItemOID)
    '        bGenerated = CheckGenerated(sShippingItem, Today.ToString("yyyyMMdd"))
    '        LoadParameter(1500, ShippingItemOID)
    '        LoadInfoParameter("2DBEB8FE6418218CE043AC190628218C", ShippingItemOID, 1500)
    '        InsertImageIntoTableResult(ShippingItemOID, 1500, "2DBEB8FE6418218CE043AC190628218C", dateView, bGenerated)
    '        'For Each drOperation In dsQC.Tables("OPERATION").Rows
    '        '    LoadParameter(drOperation.Item("OID"), ShippingItemOID)
    '        '    For Each drParameter In dsQC.Tables("QCPARAMETER").Rows
    '        '        LoadInfoParameter(drParameter.Item("OID"), ShippingItemOID, drOperation.Item("OID"))
    '        '        InsertImageIntoTableResult(ShippingItemOID, drOperation.Item("OID"), drParameter.Item("OID"), dateView, bGenerated)
    '        '    Next
    '        'Next
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub DrawingForEachSPItem(ByVal sShippingItem As String, ByVal dateView As String)
    '    Dim drOperation, drParameter As DataRow
    '    Dim bGenerated As Boolean
    '    Try
    '        'LoadShippingItem()
    '        ShippingItemOID = "db14220a-590c-42a3-8"
    '        LoadSection(ShippingItemOID)
    '        bGenerated = CheckGenerated(sShippingItem, Today.ToString("yyyyMMdd"))
    '        'For Each drOperation In dsQC.Tables("OPERATION").Rows
    '        LoadParameter(600, ShippingItemOID)
    '        For Each drParameter In dsQC.Tables("QCPARAMETER").Rows
    '            LoadInfoParameter(drParameter.Item("OID"), ShippingItemOID, 600)
    '            InsertImageIntoTableResult(ShippingItemOID, 600, drParameter.Item("OID"), dateView, bGenerated)
    '        Next
    '        'Next
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

#End Region
End Class
