Imports System.Xml
Imports System.IO
Imports System.Data.SqlClient
Imports System.Text

Public Class FormASNReader
    Dim ASNReader As String()
    Dim ASNProcessing As New List(Of FileProcessing)
    Dim htmlLineFileProcessing As String = ""
    Dim SuccessFileProcess As String = ""
    Dim FailedFileProcess As String = ""
    Dim CountSuccess As Integer = 0
    Dim CountFailed As Integer = 0

    Public Structure FileProcessing
        Dim FileName As String
        Dim ProcessingTime As String
        Dim Process As String
        Dim Status As String
        Dim Remarks As String

        Public Sub New(ByVal FileName As String, ByVal ProcessingTime As String, ByVal Process As String, ByVal Status As String, ByVal Remarks As String)
            Me.FileName = FileName
            Me.ProcessingTime = ProcessingTime
            Me.Process = Process
            Me.Status = Status
            Me.Remarks = Remarks
        End Sub
    End Structure

    Private Function GetMailRecipients(ByVal autoEmailCode As Integer) As DataSet
        Dim dsEmail As New DataSet
        Dim strSQL As String = "usp_SPT_AutoEmail_GetRecipients"
        Dim sql_handler As New SQLHandler
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@AutoEmailCode", SqlDbType.NVarChar, autoEmailCode)

        If (sql_handler.FillDataSet(strSQL, dsEmail, CommandType.StoredProcedure)) Then
        End If
        sql_handler = Nothing
        Return dsEmail
    End Function

    Private Function Save_ASN_Details(ByVal ListInvoiceDetails As Array, ByVal filename As String) As String
        Dim result As String = ""
        Dim remarks As String = CheckContentValidity(ListInvoiceDetails)

        If remarks = "" Then
            Dim ReceiveCode As Integer = Check_Existing_Invoice(ListInvoiceDetails(0))
            If ReceiveCode = 0 Then
                If Save_Invoice(ListInvoiceDetails(0)) Then
                    ReceiveCode = Check_Existing_Invoice(ListInvoiceDetails(0))
                End If
            End If

            If Not ReceiveCode = 0 Then
                If Not ifExisting_LotID(ReceiveCode, ListInvoiceDetails(5)) Then
                    If Save_Invoice_Details(ReceiveCode, ListInvoiceDetails(5), ListInvoiceDetails(3), Val(ListInvoiceDetails(8)), Val(ListInvoiceDetails(7)), ListInvoiceDetails(25)) Then
                        If Not Save_ASN_XML_Details(ListInvoiceDetails) Then  '<-- Save content of XML to database
                            remarks &= "Failed saving XML Details; "
                        End If
                    Else
                        remarks &= "Failed saving Invoice Details; "
                    End If
                Else
                    remarks &= "Item already exist; "
                End If
            End If
        End If

        Dim ProcessType As String
        If ListInvoiceDetails(4).ToString = "REWORK" Then
            ProcessType = "REWORK"
        ElseIf ListInvoiceDetails(23).ToString.Contains("R") Then
            ProcessType = "TEST ONLY"
        Else
            ProcessType = "WAFER"
        End If

        result = createHTML_Body(ListInvoiceDetails(0), ListInvoiceDetails(5), ListInvoiceDetails(3), ListInvoiceDetails(8), ListInvoiceDetails(7), remarks, filename, ListInvoiceDetails(0), ProcessType, ListInvoiceDetails(6))



        Return result
    End Function

    Private Function createHTML_Header(ByVal Status As String) As String
        Dim result As String = ""
        'result &= "New ON SEMI ASN file has been downloaded from sFTP for your review and verification.<br><br>"
        result &= "<br/ ><table style='border:solid 1px black; font:0.7em/1.5 arial,sans-serif;'>"
        result &= "<tr style='background-color: blue; font-weight:bold; text-transform:uppercase; text-align: center'><td colspan='10' style='border:solid 1px black;'><a name='STATUS'>" & Status & " Transaction</a></td></tr>"
        result &= "<tr style='font-weight:bold'>"
        result &= "<th style='border:solid 1px black; width:80px; text-align:center'>Status</td>"
        result &= "<th style='border:solid 1px black; width:150px; text-align:center'>Invoice No.</td>"
        result &= "<th style='border:solid 1px black; width:50px; text-align:center'>Custom Source</td>"
        result &= "<th style='border:solid 1px black; width:150px; text-align:center'>Lot</td>"
        result &= "<th style='border:solid 1px black; width:180px; text-align:center'>Part No.</td>"
        result &= "<th style='border:solid 1px black; width:50px; text-align:center'>No. of wafers</td>"
        result &= "<th style='border:solid 1px black; width:100px; text-align:center'>Wafer Qty</td>"
        result &= "<th style='border:solid 1px black; width:100px; text-align:center'>Process Type</td>"
        result &= "<th style='border:solid 1px black; width:200px'>Remarks</td>"
        result &= "<th style='border:solid 1px black; width:150px'>Filename</td>"
        result &= "</tr>"
        Return result
    End Function

    Private Function createHTML_Body(ByVal Invoice As String, ByVal WaferLot As String, ByVal device As String, ByVal NoofWafers As String, _
                                 ByVal WaferQty As String, ByVal Remarks As String, ByVal filename As String, ByVal Source As String, _
                                 ByVal ProcessType As String, Optional ByVal CustomSource As String = "") As String
        Dim result As String = ""
        Dim StatusColor As String = IIf(Remarks = "", "blue", "red")
        Dim Status As String = IIf(Remarks = "", "Success", "Failed")

        result &= "<tr>"
        result &= "<td style='border:solid 1px black; color:" & StatusColor & "'>" & Status & "</td>"
        result &= "<td style='border:solid 1px black;'>" & Invoice & "</td>"
        result &= "<td style='border:solid 1px black;'>" & CustomSource & "</td>"
        result &= "<td style='border:solid 1px black;'>" & WaferLot & "</td>"
        result &= "<td style='border:solid 1px black;'>" & device & "</td>"
        result &= "<td style='border:solid 1px black;'>" & NoofWafers & "</td>"
        result &= "<td style='border:solid 1px black;'>" & WaferQty & "</td>"
        result &= "<td style='border:solid 1px black;'>" & ProcessType & "</td>"
        result &= "<td style='border:solid 1px black;'>" & Remarks & "</td>"
        result &= "<td style='border:solid 1px black;'>" & filename & "</td>"
        result &= "</tr>"

        If Status = "Success" Then
            CountSuccess += 1
            SuccessFileProcess &= result
        Else
            CountFailed += 1
            FailedFileProcess &= result
        End If

        Return result
    End Function

    Private Function CheckContentValidity(ByVal ListInvoiceDetails As Array) As String
        Dim result As String = ""

        Dim material As String
        If ListInvoiceDetails(3).ToString.Contains("22528-003-ASY") Then
            material = ListInvoiceDetails(3).ToString
        ElseIf ListInvoiceDetails(3).ToString.Contains("-ASY") Then
            material = ListInvoiceDetails(3).ToString.Replace("-", "_").Replace("_ASY", "")
        ElseIf ListInvoiceDetails(3).ToString.Contains("-MEF") Then
            material = ListInvoiceDetails(3).ToString
        ElseIf ListInvoiceDetails(3).ToString.Contains("22528-003-WDQ") _
        Or ListInvoiceDetails(3).ToString.Contains("22922-903-WDQ") _
        Or ListInvoiceDetails(3).ToString.Contains("22922-001-WDQ") Then
            material = ListInvoiceDetails(3).ToString
        Else
            material = ListInvoiceDetails(3).ToString.Replace("-", "_").Replace("_WDQ", "")
        End If

        If Not check_Material(material) Then
            result &= "Part Number not found; "
        End If

        If Val(ListInvoiceDetails(7)) <= 0 Then
            result &= "Invalid Wafer Qty; "
        End If

        If ListInvoiceDetails(6).ToString.Trim = "" Then
            result &= "Invalid Custom Source; "
        End If
        'If Val(ListInvoiceDetails(8)) <= 0 Then
        '    result &= "Invalid number of wafers; "
        'End If

        Return result
    End Function


    Private Function check_Material(ByVal MaterialID As String) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        Dim ds As New DataSet
        Dim sql As String = "SELECT MaterialCode FROM PS_Material WHERE MaterialID=@MaterialID AND Active=1"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@MaterialID", SqlDbType.NVarChar, MaterialID)

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataSet(sql, ds, CommandType.Text) Then
                If Not ds.Tables(0).Rows.Count > 0 Then
                    result = False
                End If
            Else
                result = True
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        ds = Nothing

        Return result
    End Function

    Private Function Save_Invoice(ByVal Invoice As String) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        Dim sql As String = "usp_TRN_Receive_Save"
        sql_handler.CreateParameter(7)
        sql_handler.SetParameterValues(0, "@ReceiveCode", SqlDbType.BigInt, 0, ParameterDirection.Output)
        sql_handler.SetParameterValues(1, "@CustomerCode", SqlDbType.BigInt, 18)
        sql_handler.SetParameterValues(2, "@CustomerSO", SqlDbType.NVarChar, "")
        sql_handler.SetParameterValues(3, "@CustomerPO", SqlDbType.NVarChar, "")
        sql_handler.SetParameterValues(4, "@StoreCode", SqlDbType.BigInt, 1)
        sql_handler.SetParameterValues(5, "@InvoiceNo", SqlDbType.NVarChar, Invoice)
        sql_handler.SetParameterValues(6, "@UserCode", SqlDbType.BigInt, 1)

        If sql_handler.OpenConnection Then
            If Not sql_handler.ExecuteNonQuery(sql, CommandType.StoredProcedure) Then
                result = False
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function


    Private Function ifExisting_LotID(ByVal ReceiveCode As Integer, ByVal LotID As String) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        Dim ds As New DataSet
        Dim sql As String = "SELECT LotCode FROM TRN_Receive_Material WHERE ReceiveCode=@ReceiveCode AND LotID=@LotID"

        sql_handler.CreateParameter(2)
        sql_handler.SetParameterValues(0, "@ReceiveCode", SqlDbType.NVarChar, ReceiveCode)
        sql_handler.SetParameterValues(1, "@LotID", SqlDbType.NVarChar, LotID)

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataSet(sql, ds, CommandType.Text) Then
                If Not ds.Tables(0).Rows.Count > 0 Then
                    result = False
                End If
            Else
                result = True
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        ds = Nothing
        Return result
    End Function

    Private Function Save_Invoice_Details(ByVal ReceiveCode As Integer, ByVal LotID As String, ByVal PartNumber As String, _
                                          ByVal MaterialCount As Integer, ByVal MaterialQty As Integer, ByVal LotType As String) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        If PartNumber.ToString.Contains("22528-003-WDQ") _
            Or PartNumber.ToString.Contains("22922-903-WDQ") _
                Or PartNumber.ToString.Contains("22922-001-WDQ") Then
            PartNumber = PartNumber
        Else
            PartNumber = PartNumber.ToString.Replace("-", "_").Replace("_WDQ", "")
        End If

        Dim strSQL As String = "usp_TRN_Receive_Material_Save"

        sql_handler.CreateParameter(9)
        sql_handler.SetParameterValues(0, "@ReceiveCode", SqlDbType.BigInt, ReceiveCode)
        sql_handler.SetParameterValues(1, "@CustomerCode", SqlDbType.BigInt, 18)
        sql_handler.SetParameterValues(2, "@StoreCode", SqlDbType.BigInt, 1)
        sql_handler.SetParameterValues(3, "@LotID", SqlDbType.NVarChar, LotID)
        sql_handler.SetParameterValues(4, "@MaterialID", SqlDbType.NVarChar, PartNumber)
        sql_handler.SetParameterValues(5, "@MaterialCount", SqlDbType.Int, IIf(MaterialCount = 0, 1, 0))
        sql_handler.SetParameterValues(6, "@MaterialQty", SqlDbType.Int, MaterialQty)
        sql_handler.SetParameterValues(7, "@ReceivedBy", SqlDbType.BigInt, 1)
        sql_handler.SetParameterValues(8, "@Remarks", SqlDbType.NVarChar, LotType)

        If sql_handler.OpenConnection Then
            If Not sql_handler.ExecuteNonQuery(strSQL, CommandType.StoredProcedure) Then
                result = False
            End If
            sql_handler.CloseConnection()
        Else
            MsgBox(sql_handler.GetErrorMessage)
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function


    Private Function Check_Existing_Invoice(ByVal Invoice As String) As Integer
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim ds As New DataSet
        Dim sql As String = "SELECT ReceiveCode FROM TRN_Receive WHERE InvoiceNo=@InvoiceNo"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@InvoiceNo", SqlDbType.NVarChar, Invoice)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataSet(sql, ds, CommandType.Text) Then
                If ds.Tables(0).Rows.Count > 0 Then
                    result = Val(ds.Tables(0).Rows(0).Item("ReceiveCode").ToString)
                End If
            Else
                result = 0
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        ds = Nothing
        Return result
    End Function

    Private Function Save_ASN_XML_Details(ByVal ListInvoiceDetails As Array) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        Dim sql As String = "INSERT INTO CST_ONSEMI_ASN_Receive (invoicenumber, intransitdate, eta, device, lotcategory, lotnbr, customsource, lotqty, " & _
                            " numberofwafer, waferdiameter, waferthickness, micnumber, probesite, probedate, assyloc, testloc, datecode1, " & _
                            " datecode2, dispodate, moonumber, sealcode, topsdmk, backsdmk, materialstatus, pti, lottype, Wafer_Attributes, DateTrans )" & _
                            "VALUES (@invoicenumber, @intransitdate, @eta, @device, @lotcategory, @lotnbr, @customsource, @lotqty, " & _
                            " @numberofwafer, @waferdiameter, @waferthickness, @micnumber, @probesite, @probedate, @assyloc, @testloc, @datecode1, " & _
                            " @datecode2, @dispodate, @moonumber, @sealcode, @topsdmk, @backsdmk, @materialstatus, @pti, @lottype, @Wafer_Attributes, getDate() )"

        sql_handler.CreateParameter(27)
        sql_handler.SetParameterValues(0, "@invoicenumber", SqlDbType.NVarChar, ListInvoiceDetails(0))
        sql_handler.SetParameterValues(1, "@intransitdate", SqlDbType.NVarChar, ListInvoiceDetails(1))
        sql_handler.SetParameterValues(2, "@eta", SqlDbType.NVarChar, ListInvoiceDetails(2))
        sql_handler.SetParameterValues(3, "@device", SqlDbType.NVarChar, ListInvoiceDetails(3))
        sql_handler.SetParameterValues(4, "@lotcategory", SqlDbType.NVarChar, ListInvoiceDetails(4))
        sql_handler.SetParameterValues(5, "@lotnbr", SqlDbType.NVarChar, ListInvoiceDetails(5))
        sql_handler.SetParameterValues(6, "@customsource", SqlDbType.NVarChar, ListInvoiceDetails(6))
        sql_handler.SetParameterValues(7, "@lotqty", SqlDbType.NVarChar, ListInvoiceDetails(7))
        sql_handler.SetParameterValues(8, "@numberofwafer", SqlDbType.NVarChar, ListInvoiceDetails(8))
        sql_handler.SetParameterValues(9, "@waferdiameter", SqlDbType.NVarChar, ListInvoiceDetails(9))
        sql_handler.SetParameterValues(10, "@waferthickness", SqlDbType.NVarChar, ListInvoiceDetails(10))
        sql_handler.SetParameterValues(11, "@micnumber", SqlDbType.NVarChar, ListInvoiceDetails(11))
        sql_handler.SetParameterValues(12, "@probesite", SqlDbType.NVarChar, ListInvoiceDetails(12))
        sql_handler.SetParameterValues(13, "@probedate", SqlDbType.NVarChar, ListInvoiceDetails(13))
        sql_handler.SetParameterValues(14, "@assyloc", SqlDbType.NVarChar, ListInvoiceDetails(14))
        sql_handler.SetParameterValues(15, "@testloc", SqlDbType.NVarChar, ListInvoiceDetails(15))
        sql_handler.SetParameterValues(16, "@datecode1", SqlDbType.NVarChar, ListInvoiceDetails(16))
        sql_handler.SetParameterValues(17, "@datecode2", SqlDbType.NVarChar, ListInvoiceDetails(17))
        sql_handler.SetParameterValues(18, "@dispodate", SqlDbType.NVarChar, ListInvoiceDetails(18))
        sql_handler.SetParameterValues(19, "@moonumber", SqlDbType.NVarChar, ListInvoiceDetails(19))
        sql_handler.SetParameterValues(20, "@sealcode", SqlDbType.NVarChar, ListInvoiceDetails(20))
        sql_handler.SetParameterValues(21, "@topsdmk", SqlDbType.NVarChar, ListInvoiceDetails(21))
        sql_handler.SetParameterValues(22, "@backsdmk", SqlDbType.NVarChar, ListInvoiceDetails(22))
        sql_handler.SetParameterValues(23, "@materialstatus", SqlDbType.NVarChar, ListInvoiceDetails(23))
        sql_handler.SetParameterValues(24, "@pti", SqlDbType.NVarChar, ListInvoiceDetails(24))
        sql_handler.SetParameterValues(25, "@lottype", SqlDbType.NVarChar, ListInvoiceDetails(25))
        sql_handler.SetParameterValues(26, "@Wafer_Attributes", SqlDbType.NVarChar, ListInvoiceDetails(26))

        If sql_handler.OpenConnection Then
            If Not sql_handler.ExecuteNonQuery(sql, CommandType.Text) Then
                result = False
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing

        Return result
    End Function

    Private Sub FormASNReader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim em As New EmailHandler
        Dim xmlDoc As XmlDocument
        Dim foldername As String = Now.Date.ToString("MMddyyyy")
        'Dim destinationfolder As String = "\\192.168.5.20\fsc_cap\ONSEMI_ASN_Backup\"
        Dim destinationfolder As String = "D:\intransit\Backup\"

        Dim ds As DataSet
        Dim reader As StringReader
        Dim ListInvoice(26) As String
        Dim ctr As Integer = 0
        Dim htmlString As String = ""
        Dim emailString As String = ""
        Dim dsEmail As DataSet = GetMailRecipients(15)
        'Dim ASNPath As String = "\\192.168.5.20\sftp\FSC\intransit\"       '<-- Live Environment
        Dim ASNPath As String = "D:\intransit\"

        ASNReader = IO.Directory.GetFiles(ASNPath, "*.xml")
        Dim BackupFile As Boolean = False
        Dim BackupFileRemarks As String = ""
        Dim movedSuccessCount As Integer = 0
        CountSuccess = 0
        CountFailed = 0
        SuccessFileProcess = ""
        FailedFileProcess = ""

        If Not IO.Directory.Exists(destinationfolder & "\" & foldername) Then
            IO.Directory.CreateDirectory(destinationfolder & foldername)
        End If

        For i As Integer = 0 To ASNReader.Length - 1
            Dim ASNFilename As String = Path.GetFileName(ASNReader(i))
            ASNProcessing.Add(New FileProcessing("", Now, "New ASN File found: " & ASNFilename, "", ""))

            xmlDoc = New XmlDocument()
            xmlDoc.Load(ASNPath & ASNFilename)

            ds = New DataSet()
            reader = New StringReader(xmlDoc.InnerXml)
            ds.ReadXml(reader)

            For Each t As DataTable In ds.Tables
                Console.WriteLine(String.Format("{0}: {1}", t.TableName, t.Rows.Count))
                If t.TableName = "header" Or t.TableName = "Wafer_Attributes" Or t.TableName = "Wafer_Information" Then
                    'Skip header node in XML
                    GoTo skip
                End If

                For Each r As DataRow In t.Rows
                    Dim NodeValue As String = ""
                    ctr = 0
                    For Each c As DataColumn In t.Columns
                        If c.ColumnName = "message_Id" Then
                            GoTo nextNode
                        End If
                        'Get message node inside XML
                        NodeValue = r(c.ColumnName)
                        ListInvoice(ctr) = NodeValue '<-- Add node value per invoice to array string
                        ctr += 1
nextNode:           Next

                    If ctr > 0 Then
                        htmlString &= Save_ASN_Details(ListInvoice, ASNFilename)
                    End If

                Next
skip:       Next

            Try
                File.Move(ASNPath & ASNFilename, destinationfolder & foldername & "\" & ASNFilename) '<-- Move current file to backup folder
                movedSuccessCount += 1
            Catch ex As Exception
                ASNProcessing.Add(New FileProcessing(ASNFilename, Now, "Moved to backup folder", "Failed", ASNFilename & ": " & ex.Message))
            End Try
        Next

        If ASNReader.Length > 0 Then
            ASNProcessing.Add(New FileProcessing("", Now, "Moved to backup folder", "Success", movedSuccessCount & " file(s) successfully back up"))
            ASNProcessing.Add(New FileProcessing("", Now, "Get No. of successful transaction", "Success", CountSuccess & " lot(s) found"))
            ASNProcessing.Add(New FileProcessing("", Now, "Get No. of failed transaction", "Success", CountFailed & " lot(s) found"))
            ASNProcessing.Add(New FileProcessing("", Now, "See below details", "", ""))

            htmlLineFileProcessing = GetFileProcessingHTML(ASNProcessing)
            htmlLineFileProcessing &= "<table style='font:0.7em/1.5 arial,sans-serif;'>" & createHTML_Header("Success") & SuccessFileProcess & "</table><a href='#top'>go to top</a><br/><br/>"
            htmlLineFileProcessing &= "<table style='font:0.7em/1.5 arial,sans-serif;'>" & createHTML_Header("Failed") & FailedFileProcess & "</table><a href='#top'>go to top</a><br/><br/>"

            'emailString = "<html><body style='font:0.8em/1.5 arial,sans-serif; text-align:left;'>" & createHTML_Header() & htmlString & "</table></body></html>"
            em.SendEmail("ONSEMI Receive Alert", htmlLineFileProcessing, "", dsEmail)
        End If
        End
    End Sub



    Private Function GetFileProcessingHTML(ByVal _fileProcessing As List(Of FileProcessing), Optional ByVal filename As String = "") As String
        Dim htmlLine As String = ""
        Dim list As List(Of FileProcessing) = _fileProcessing.FindAll(Function(r) r.FileName = filename OrElse r.FileName = "")

        If list.Count > 0 Then
            htmlLine &= "<table style='border:solid 1px black; font:0.7em/1.5 arial,sans-serif;'>"
            htmlLine &= "<tr style='background-color: blue; font-weight:bold; text-transform:uppercase; text-align: center'><td colspan='4' style='border:solid 1px black;'><a name='STATUS'>ASN File Processing Status</a></td></tr>"
            htmlLine &= "<tr style='font-weight:bold'>"
            htmlLine &= "<th style='border:solid 1px black; width:150px; text-align:center'>Time</td>"
            htmlLine &= "<th style='border:solid 1px black; width:200px'>Process Description</td>"
            htmlLine &= "<th style='border:solid 1px black; width:50px; text-align:center'>Status</td>"
            htmlLine &= "<th style='border:solid 1px black; width:200px'>Remarks</td>"
            htmlLine &= "</tr>"

            For Each fprocess As FileProcessing In list
                htmlLine &= "<tr>"
                htmlLine &= "<td style='border:solid 1px black;'>" & fprocess.ProcessingTime & "</td>"
                htmlLine &= "<td style='border:solid 1px black;'>" & fprocess.Process & "</td>"
                htmlLine &= "<td style='border:solid 1px black; color:blue'>" & fprocess.Status & "</td>"
                htmlLine &= "<td style='border:solid 1px black;'>" & fprocess.Remarks & "</td>"
                htmlLine &= "</tr>"
            Next
            htmlLine &= "</table><a href='#top'>go to top</a><br/><br/>"
        End If
        Return htmlLine
    End Function

End Class
