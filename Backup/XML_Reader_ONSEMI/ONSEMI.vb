Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data

Public Class ONSEMI
    Dim OptionCode As String = ""

#Region "Functions"
    Public Function InsertTransactionLogs(ByVal LotNumberCode As Integer, ByVal OldValue As String, _
                                          ByVal NewValue As String, ByVal Type As String, ByVal Remarks As String, _
                                          ByVal UserCode As String) As Boolean
        Dim result As Boolean = False
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "INSERT INTO  SUP_Transaction_Logs (LotCode, OldValue, NewValue, Type, Remarks, UserCode, DateTime) " & _
                            "VALUES (@LotCode, @OldValue, @NewValue, @Type, @Remarks, @UserCode, getdate())"
        sql_handler.CreateParameter(6)
        sql_handler.SetParameterValues(0, "@LotCode", SqlDbType.NVarChar, LotNumberCode)
        sql_handler.SetParameterValues(1, "@OldValue", SqlDbType.NVarChar, OldValue)
        sql_handler.SetParameterValues(2, "@NewValue", SqlDbType.NVarChar, NewValue)
        sql_handler.SetParameterValues(3, "@Type", SqlDbType.NVarChar, Type)
        sql_handler.SetParameterValues(4, "@Remarks", SqlDbType.NVarChar, Remarks)
        sql_handler.SetParameterValues(5, "@UserCode", SqlDbType.NVarChar, UserCode)
        If sql_handler.OpenConnection Then
            If sql_handler.ExecuteNonQuery(sql, CommandType.Text) Then
                result = True
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GenerateMTRCode(ByVal CustomerCode As String) As String
        Dim result As String = ""


        Return result
    End Function

    Public Function Save_PL_ProductionOrder(ByVal LotNumber As String, ByVal PONumber As String, ByVal WaferLot As String, ByVal LotCode As String, _
                                            ByVal StartDate As String, ByVal CommitDate As String, ByVal Destination As String, ByVal DateCode As String, _
                                            ByVal LotSuffix As String, ByVal CustomerCode As String, ByVal DeviceCode As String, ByVal DiePartCode As String, _
                                            ByVal ProductCode As String, ByVal Keygroup As String, ByVal LotQty As Integer, Optional Remarks As String = "") As Boolean
        Dim result As Boolean = False
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim strSQL As String = "INSERT INTO PL_ProductionOrder " & _
                               "(LotNumber, PONumber, WaferLot, LotQty, LotCode, ReceivedDate, " & _
                               "StartDate, CommitDate, Destination, DateCode, Remarks, Auxiliary1, Auxiliary2, " & _
                               "LotSuffix, TopMark, BottomMark, WireType, CustomerCode, DeviceCode, DiePartCode, " & _
                               "ProductCode, ProcessTypeCode, UserCode, KeyGroup) " & _
                               "VALUES(@LotNumber, @PONumber, @WaferLot, @LotQty, @LotCode, @ReceivedDate, " & _
                               "@StartDate, @CommitDate, @Destination, @DateCode, @Remarks, @Auxiliary1, @Auxiliary2, " & _
                               "@LotSuffix, @TopMark, @BottomMark, @WireType, @CustomerCode, @DeviceCode, @DiePartCode, " & _
                               "@ProductCode, @ProcessTypeCode, @UserCode, @KeyGroup)"


        sql_handler.CreateParameter(24)
        sql_handler.SetParameterValues(0, "@LotNumber", SqlDbType.VarChar, LotNumber)
        sql_handler.SetParameterValues(1, "@PONumber", SqlDbType.VarChar, PONumber)
        sql_handler.SetParameterValues(2, "@WaferLot", SqlDbType.VarChar, WaferLot)
        sql_handler.SetParameterValues(4, "@LotCode", SqlDbType.VarChar, LotCode)
        sql_handler.SetParameterValues(5, "@ReceivedDate", SqlDbType.VarChar, Now.Date)
        sql_handler.SetParameterValues(6, "@StartDate", SqlDbType.VarChar, StartDate)
        sql_handler.SetParameterValues(7, "@CommitDate", SqlDbType.VarChar, CommitDate)
        sql_handler.SetParameterValues(8, "@Destination", SqlDbType.VarChar, Destination)
        sql_handler.SetParameterValues(9, "@DateCode", SqlDbType.VarChar, DateCode)
        sql_handler.SetParameterValues(10, "@Remarks", SqlDbType.VarChar, Remarks)
        sql_handler.SetParameterValues(11, "@Auxiliary1", SqlDbType.VarChar, "")
        sql_handler.SetParameterValues(12, "@Auxiliary2", SqlDbType.VarChar, "")
        sql_handler.SetParameterValues(13, "@LotSuffix", SqlDbType.VarChar, LotSuffix)
        sql_handler.SetParameterValues(14, "@TopMark", SqlDbType.VarChar, "")
        sql_handler.SetParameterValues(15, "@BottomMark", SqlDbType.VarChar, "")
        sql_handler.SetParameterValues(16, "@WireType", SqlDbType.VarChar, "")
        sql_handler.SetParameterValues(17, "@CustomerCode", SqlDbType.VarChar, CustomerCode)
        sql_handler.SetParameterValues(18, "@DeviceCode", SqlDbType.VarChar, DeviceCode)
        sql_handler.SetParameterValues(19, "@DiePartCode", SqlDbType.VarChar, DiePartCode)
        sql_handler.SetParameterValues(20, "@ProductCode", SqlDbType.VarChar, ProductCode)
        sql_handler.SetParameterValues(21, "@ProcessTypeCode", SqlDbType.VarChar, 1)
        sql_handler.SetParameterValues(22, "@UserCode", SqlDbType.VarChar, 1)
        sql_handler.SetParameterValues(23, "@KeyGroup", SqlDbType.VarChar, Keygroup)
        sql_handler.SetParameterValues(3, "@LotQty", SqlDbType.VarChar, LotQty)

        If sql_handler.OpenConnection Then
            If sql_handler.ExecuteNonQuery(strSQL, CommandType.Text) Then
                result = True
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function


    Public Function UpdateDieCurrent_Qty(ByVal LotCode As String, ByVal Qty As Integer) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        Dim curntQty As Integer = GetCurrent_Qty(LotCode)
        Dim diff As Integer = Val(curntQty) - Val(Qty)
        Dim sql As String = "UPDATE TRN_Lot SET CurrentQty=@CurrentQty WHERE LotCode=@LotCode"
        sql_handler.CreateParameter(2)
        sql_handler.SetParameterValues(0, "@LotCode", SqlDbType.NVarChar, LotCode)
        sql_handler.SetParameterValues(1, "@CurrentQty", SqlDbType.NVarChar, diff)
        If sql_handler.OpenConnection Then
            If sql_handler.ExecuteNonQuery(sql, CommandType.Text) Then
                result = True
                If diff = 0 Then
                    UpdateLotStatus_Terminate(LotCode)
                End If
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function isYieldingStation(ByVal FlowCode As Integer, ByVal RecipeCode As Integer, ByVal stageCode As Integer) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT * FROM PS_Flow_YieldingStation WHERE FlowCode=@FlowCode AND RecipeCode=@RecipeCode AND stageCode=@stageCode"
        sql_handler.CreateParameter(3)
        sql_handler.SetParameterValues(0, "@FlowCode", SqlDbType.NVarChar, FlowCode)
        sql_handler.SetParameterValues(1, "@RecipeCode", SqlDbType.NVarChar, RecipeCode)
        sql_handler.SetParameterValues(2, "@stageCode", SqlDbType.NVarChar, stageCode)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If Not dr.Read Then
                    result = False
                End If
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function isYieldingCustomer(ByVal CustomerCode As Integer) As Boolean
        Dim result As Boolean = False
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT * FROM MS_Customer_YieldingTransaction WHERE CustomerCode=@CustomerCode AND isYielding=1"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@CustomerCode", SqlDbType.NVarChar, CustomerCode)

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = True
                End If
            Else
                result = False
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function isCustomerUseDeviceinSETUP(ByVal CustomerCode As Integer) As Boolean
        Dim result As Boolean = False
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT * FROM MS_Customer_YieldingTransaction WHERE CustomerCode=@CustomerCode AND isSetUpDevice=1"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@CustomerCode", SqlDbType.NVarChar, CustomerCode)

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = True
                End If
            Else
                result = False
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function checkifDualSameDie(ByVal Device As String) As Boolean
        Dim result As Boolean = False
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT Device FROM CST_ONSEMI_DualSameDie_List WHERE Device=@Device"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@Device", SqlDbType.NVarChar, Device)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = True
                End If
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function


    Public Function Save_AutoDateCode(ByVal POCode As Integer, ByVal DateCode As String, ByVal User As String) As Boolean
        Dim result As Boolean = True

        Dim sql_handler As New SQLHandler
        Dim ds As New DataSet
        Dim DateCodeMode As String = "DA"
        Dim Shift As String = ""

        If (Now > Now.Date & " 06:00AM" And Now < Now.Date & " 02:00PM") Then
            Shift = "A"
        ElseIf (Now > Now.Date & " 02:00PM" And Now < Now.Date & " 10:00PM") Then
            Shift = "B"
        Else
            Shift = "C"
        End If

        Dim strSQL As String = "INSERT INTO  PL_ProductionOrder_DateCode_Detail (POCode, DateCode, ShiftCode, TranDate, TranBy, VerifiedDate, VerifiedBy, DateCodeMode) " & _
                               "VALUES (@POCode, @DateCode, @ShiftCode, @TranDate, @TranBy, @VerifiedDate, @VerifiedBy, @DateCodeMode)"
        sql_handler.CreateParameter(8)
        sql_handler.SetParameterValues(0, "@POCode", SqlDbType.BigInt, POCode)
        sql_handler.SetParameterValues(1, "@DateCode", SqlDbType.NVarChar, DateCode)
        sql_handler.SetParameterValues(2, "@ShiftCode", SqlDbType.NVarChar, Shift)
        sql_handler.SetParameterValues(3, "@TranDate", SqlDbType.Date, Now)
        sql_handler.SetParameterValues(4, "@TranBy", SqlDbType.BigInt, User)
        sql_handler.SetParameterValues(5, "@VerifiedDate", SqlDbType.Date, Now)
        sql_handler.SetParameterValues(6, "@VerifiedBy", SqlDbType.BigInt, User)
        sql_handler.SetParameterValues(7, "@DateCodeMode", SqlDbType.NVarChar, DateCodeMode)

        If sql_handler.OpenConnection Then
            If Not sql_handler.ExecuteNonQuery(strSQL, CommandType.Text) Then
                result = False
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing

        Return result
    End Function

    Public Function UpdateShipmentStatus_Shipped(ByVal LotCode As Long) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        Dim sql As String = "UPDATE TRN_Lot SET CurrentStatusCode=7 WHERE LotCode=@LotCode"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@LotCode", SqlDbType.BigInt, LotCode)
        If sql_handler.OpenConnection Then
            If Not sql_handler.ExecuteNonQuery(sql, CommandType.Text) Then
                result = False
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function UpdateLotStatus_Terminate(ByVal LotCode As Long) As Boolean
        Dim result As Boolean = True
        Dim sql_handler As New SQLHandler
        Dim sql As String = "UPDATE TRN_Lot SET CurrentStatusCode=6 WHERE LotCode=@LotCode"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@LotCode", SqlDbType.BigInt, LotCode)
        If sql_handler.OpenConnection Then
            If Not sql_handler.ExecuteNonQuery(sql, CommandType.Text) Then
                result = False
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetLot(ByVal strLot As String) As String
        GetLot = strLot
        Dim pos As Integer = InStrRev(strLot, "-")
        If pos > 0 Then
            GetLot = Mid(strLot, 1, pos - 1)
        End If
        Return GetLot
    End Function

    Public Function getShipmentPO(ByVal Device As String) As String()
        Dim result(2) As String
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT POID,UnitPrice  FROM CST_ONSEMI_POLINE_Details WHERE ItemNumber=@ItemNumber"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@ItemNumber", SqlDbType.NVarChar, Device)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result(0) = dr.Item("POID")
                    result(1) = dr.Item("UnitPrice")
                End If
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        Return result
    End Function


    Public Function GetSetupQty(ByVal Package As String, ByVal LeadType As String) As Integer
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT SublotQty FROM  CST_ONSEMI_PackageQty_Setup " & _
                            "WHERE Package='" & Package & "' AND LeadCount ='" & LeadType & "'"
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("SublotQty")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetPOCode(ByVal LotNumber As String) As Integer
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT POCode FROM  PL_ProductionOrder WHERE LotNumber=@LotNumber"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@LotNumber", SqlDbType.NVarChar, LotNumber)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("POCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            result = 0
        End If
        sql_handler = Nothing
        Return result
    End Function


    Public Function isAliasExist_PL_ProductionOrders(ByVal LotNumber As String) As Boolean
        Dim result As Boolean = False
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT Lotnumber FROM  PL_ProductionOrder WHERE LotNumber=@LotNumber"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@LotNumber", SqlDbType.NVarChar, LotNumber)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = True
                End If
            End If
            sql_handler.CloseConnection()
        Else
            result = True
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetCurrent_Qty(ByVal LotCode As String) As Integer
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT CurrentQty FROM TRN_Lot WHERE LotCode=@LotCode"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@LotCode", SqlDbType.NVarChar, LotCode)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("CurrentQty")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            result = 0
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetLotDateCode(ByVal lotnumber As String) As String
        Dim result As String = ""
        Dim YearCode As String = Now.Date.Year
        Dim WeekChar As String = "W"
        Dim WorkWeek As String = lotnumber.Substring(3, 2)

        result = YearCode & WeekChar & WorkWeek
        Return result
    End Function

    Public Function GetProdcutCode(ByVal Device As String, ByVal BuildType As String, ByVal ProcessType As String) As String
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT TOP 1 B.ProductCode  " & _
                            "FROM  PS_Product_Attribute A INNER JOIN " & _
                            "PS_Product B on A.ProductCode=B.ProductCode AND B.Active=1 AND A.ParameterCode=5 " & _
                            "WHERE A.ParameterValue='" & Device & "' AND B.ProductID LIKE 'P%-" & BuildType & "%'"
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("ProductCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetMaterialCode(ByVal MaterialID As String) As String
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT MaterialCode  " & _
                            "FROM  PS_Material " & _
                            "WHERE MaterialID='" & MaterialID & "'"
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("MaterialCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function


    Public Function GetActive_PONumber(ByVal MaterialID As String) As String
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT MaterialCode  " & _
                            "FROM  PS_Material " & _
                            "WHERE MaterialID='" & MaterialID & "'"
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("MaterialCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetPCI_Length(ByVal ProductCode As String) As Integer
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT LEN(ParameterValue) AS PCI " & _
                            "FROM  PS_Product_Attribute " & _
                            "WHERE ProductCode='" & ProductCode & "' AND ParameterCode=201"
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("PCI")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetProductCode_V2(ByVal Device As String, ByVal DiePart As String, ByVal Customer As String) As String
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT  B.ProductCode  " & _
                            "FROM  PS_Product_Attribute A INNER JOIN " & _
                            "PS_Product B on A.ProductCode=B.ProductCode AND B.Active=1 AND A.ParameterCode=5 INNER JOIN " & _
                            "PS_Product_Attribute AS C on B.ProductCode=C.ProductCode AND C.ParameterCode=201 INNER JOIN " & _
                            "PS_Product_BOM AS D on D.ProductCode=B.ProductCode INNER JOIN " & _
                            "PS_Material as E on E.MaterialCode=D.MaterialCode AND E.MaterialDescription IN ('DIE','DIE1','DIE2') INNER JOIN " & _
                            "MS_Customer AS F on B.CustomerCode=F.CustomerCode " & _
                            "WHERE A.ParameterValue='" & Device & "' AND B.ProductID LIKE 'P%' AND LEN(C.ParameterValue)>12 AND E.MaterialID='" & DiePart & "' " & _
                            "AND F.CustomerID='" & Customer & "'"
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("ProductCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetProductCode_PDIP(ByVal Device As String, ByVal BuildType As String, ByVal ICDie As String, _
                                        ByVal FETDie As String, ByVal Customer As String) As String
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT  TOP 1 B.ProductCode,E.MaterialID AS Device,B.ProductID, E.MaterialID AS ICDie,H.MaterialID as FETDie " & _
                            "FROM  PS_Product_Attribute A INNER JOIN " & _
                            "PS_Product B on A.ProductCode=B.ProductCode AND B.Active=1 AND A.ParameterCode=5 INNER JOIN " & _
                            "PS_Product_Attribute AS C on B.ProductCode=C.ProductCode AND C.ParameterCode=201 INNER JOIN " & _
                            "PS_Product_BOM AS D on D.ProductCode=B.ProductCode " & _
                            "LEFT  JOIN (SELECT AA.ProductCode, BB.MaterialCode, BB.MaterialID, BB.MaterialDescription FROM PS_Product_BOM AA  " & _
                            "LEFT JOIN PS_Material BB On AA.MaterialCode = BB.MaterialCode) E On A.ProductCode = E.ProductCode And E.MaterialDescription IN ('IC', 'IC DIE','DIE','DIE1','DIE 1') " & _
                            "LEFT  JOIN (SELECT AA.ProductCode, BB.MaterialCode, BB.MaterialID, BB.MaterialDescription FROM PS_Product_BOM AA " & _
                            "LEFT JOIN PS_Material BB On AA.MaterialCode = BB.MaterialCode) H On A.ProductCode = H.ProductCode And H.MaterialDescription IN ('FET', 'FET DIE','DIE2','DIE 2') " & _
                            "LEFT JOIN MS_Customer AS F on B.CustomerCode=F.CustomerCode " & _
                            "WHERE A.ParameterValue=@Device AND B.ProductID LIKE 'P%' + @BuildType AND LEN(C.ParameterValue) > 10 " & _
                            "AND E.MaterialID=@ICDie AND H.MaterialID=@FETDie AND F.CustomerID=@CustomerID"

        sql_handler.CreateParameter(5)
        sql_handler.SetParameterValues(0, "@Device", SqlDbType.NVarChar, Device)
        sql_handler.SetParameterValues(1, "@BuildType", SqlDbType.NVarChar, BuildType)
        sql_handler.SetParameterValues(2, "@ICDie", SqlDbType.NVarChar, ICDie)
        sql_handler.SetParameterValues(3, "@FETDie", SqlDbType.NVarChar, FETDie)
        sql_handler.SetParameterValues(4, "@CustomerID", SqlDbType.NVarChar, Customer)

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("ProductCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function


    Public Function GetProduct_Code_RMA(ByVal Device As String, ByVal Customer As String) As String()
        Dim result(2) As String
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT B.ProductID,B.ProductCode " & _
                            "FROM PS_Product_Attribute A " & _
                            "INNER JOIN PS_Product B on B.ProductCode=A.ProductCode " & _
                            "LEFT JOIN PS_Product_Attribute C on A.ProductCode=C.ProductCode AND C.ParameterCode=201 " & _
                            "LEFT JOIN MS_Customer D on B.CustomerCode=D.CustomerCode " & _
                            "WHERE A.ParameterValue=@Device AND LEN(C.ParameterValue) > 10 AND B.ProductID LIKE 'T%' AND D.CustomerID=@CustomerID "

        sql_handler.CreateParameter(2)
        sql_handler.SetParameterValues(0, "@Device", SqlDbType.NVarChar, Device)
        sql_handler.SetParameterValues(1, "@CustomerID", SqlDbType.NVarChar, Customer)

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result(0) = dr.Item("ProductCode")
                    result(1) = dr.Item("ProductID")
                Else
                    result(0) = 0
                End If
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GenerateLotCode(ByVal startDate As String, ByVal LastLotCode As String) As String
        Dim sql_handler As New SQLHandler
        Dim dr As SqlClient.SqlDataReader = Nothing
        Dim lotcode As String = ""
        Dim YearCodeValue As String = ""
        Dim WeekCodeValue As String = ""

        If LastLotCode <> "" Then
            If LastLotCode <> "" Then
                If LastLotCode.Length = 2 Then
                    Dim first As String = Strings.Left(LastLotCode, 1)
                    Dim second As String = Strings.Right(LastLotCode, 1)
                    If second = "Z" Then
                        second = "A"
                        If first = "Z" Then
                            first = "A"
                        ElseIf first = "9" Then
                            first = "A"
                            second = "A"
                        Else
                            first = ChrW(Asc(first) + 1)
                        End If
                    Else
                        If second = "9" Then
                            If first = "Z" Then
                                first = 1
                                second = "A"
                            Else
                                first = ChrW(Asc(first) + 1)
                                second = 1
                            End If
                        Else
                            second = ChrW(Asc(second) + 1)
                        End If
                    End If
                    lotcode = first + second
                ElseIf lotcode.Length = 1 Then
                    lotcode = "AA"
                Else
                    lotcode = ""
                End If
            End If
            Return lotcode
        End If

        Dim sqlText As String = "SELECT TOP 1 LotCode  FROM PL_ProductionOrder " & _
                                "WHERE (LotNumber LIKE '" & GetSiteCode() & YearMapping() & WorkWeekCode(startDate) & "%') " & _
                                "ORDER BY POCode DESC "

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sqlText, dr, CommandType.Text) Then
                If dr.Read Then
                    lotcode = dr.Item("LotCode")
                    If lotcode = "" Then
                        lotcode = "A"
                    End If
                Else
                    lotcode = "A"
                End If

                dr.Close()
                dr = Nothing
                If lotcode <> "" Then
                    If lotcode.Length = 2 Then
                        Dim first As String = Strings.Left(lotcode, 1)
                        Dim second As String = Strings.Right(lotcode, 1)
                        If second = "Z" Then
                            If first = "Z" Then
                                first = "A"
                                second = 1
                            ElseIf first = "9" Then
                                first = "A"
                                second = "A"
                            Else
                                first = ChrW(Asc(first) + 1)
                            End If
                        Else
                            If second = "9" Then
                                If first = "Z" Then
                                    first = 1
                                    second = "A"
                                Else
                                    first = ChrW(Asc(first) + 1)
                                    second = 1
                                End If
                            Else
                                second = ChrW(Asc(second) + 1)
                            End If
                        End If
                        lotcode = first + second
                    ElseIf lotcode.Length = 1 Then
                        lotcode = "AA"
                    Else
                        lotcode = ""
                    End If
                End If
            Else
                lotcode = ""
            End If
            sql_handler.CloseConnection()
        Else
            lotcode = ""
        End If


        sql_handler = Nothing
        Return lotcode
    End Function

    Public Function getCustomerCode_PerLot(ByVal LotNumber As String) As Integer
        Dim result As Integer = 0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim strSQL As String = "SELECT CustomerCode FROM TRN_LotStart WHERE LotAlias=@LotNumber"

        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@LotNumber", SqlDbType.NVarChar, LotNumber)

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(strSQL, dr, CommandType.Text) Then
                If dr.Read Then
                    result = Val(dr.Item("CustomerCode"))
                End If
                dr.Close()
                dr = Nothing
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        Return result
    End Function
#End Region

#Region "Create Lot Numbering SQL"

    '' Generate PID for Dumping
    Public Function GeneratedPID(ByVal startDate As String, ByVal LoopCount As Integer) As String
        Dim PID As String = ""
        Dim SiteCode As String = GetSiteCode()
        Dim YearMappingCode As String = YearMapping()
        Dim WeekCode As String = WorkWeekCode(startDate)
        Dim PIDCount As Integer = PIDCounter(SiteCode & YearMappingCode & WeekCode)
        PIDCount += LoopCount + 1
        PID = SiteCode & YearMappingCode & WeekCode & PIDCount.ToString.PadLeft(4, "0") & OptionCode
        Return PID
    End Function

    ' Create LotNumber based on setup on SQL
    Public Function GetCustomerCode(ByVal CustomerID As String) As String
        Dim result As String = ""
        Dim sql As String = "SELECT CustomerCode FROM MS_Customer WHERE CustomerID='" & CustomerID & "'"
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("CustomerCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetSiteCode() As String
        Dim result As String = ""
        Dim sql As String = "SELECT SiteCode FROM CST_ONSEMI_SiteCode WHERE Active=1"
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("SiteCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function YearMapping() As String
        Dim result As String = ""
        Dim sql As String = "SELECT YearCode FROM CST_ONSEMI_YearMapping WHERE Year='" & Now.Date.Year & "'"
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("YearCode")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function WorkWeekCode(ByVal startdate As String) As String
        Dim result As String = ""
        Dim sql As String = "SELECT Workweek FROM  CST_ONSEMI_Calendar " & _
                             "WHERE FromDate <= '" & startdate & "' And ToDate >= '" & startdate & "'"
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("Workweek").ToString.PadLeft(2, "0")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function GetDestination(ByVal Operations As String, ByVal IsCaptive As Boolean) As String
        Dim result As String = ""
        Dim sql As String = "SELECT Destination FROM  CST_ONSEMI_DestinationMapping " & _
                            "WHERE Operations ='" & Operations & "' AND Captive='" & IsCaptive & "'"
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("Destination").ToString
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function PIDCounter(ByVal LotChar As String) As Integer
        Dim result As Integer = 0
        Dim sql As String = "SELECT COUNT(DISTINCT LotNumber) as PIDCtr FROM PL_ProductionOrder WHERE Lotnumber LIKE '" & LotChar & "%'"
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing

        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = dr.Item("PIDCtr")
                End If
            End If
            sql_handler.CloseConnection()
        Else
            Return sql_handler.GetErrorMessage
        End If
        sql_handler = Nothing
        Return result
    End Function
#End Region

#Region "Auto Hold Reject"
    Public Function AutoHOLD_BIN3_BIN4(ByVal Lotcode As String, ByVal RecipeSequence As Integer, _
                                           ByVal QtyIn As Integer, ByVal productCode As Integer, ByRef HoldReject As String, _
                                           ByRef setupBINLIMIT As String) As Integer
        Dim result As Double = 0.0
        Dim sql_handler As New SQLHandler
        Dim ds As New DataSet
        Dim rejectPercentage As Double = 0.0
        Dim rejectValue As Integer = 0
        Dim BIN_Limit As Double = 0.0
        Dim BIN3 As Boolean = False
        Dim BIN4 As Boolean = False
        Dim sql As String = "SELECT B.RejectCode, B.RejectID, B.RejectDescription, A.RejectQty " & _
                            "FROM TRN_Lot_Rejects A INNER JOIN PS_Reject B On A.RejectCode = B.RejectCode " & _
                            "WHERE A.LotCode = @LotCode And A.SequenceCode = @SequenceCode"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@LotCode", SqlDbType.Int, Lotcode)
        sql_handler.SetParameterValues(1, "@SequenceCode", SqlDbType.Int, RecipeSequence)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataSet(sql, ds, CommandType.Text) Then
                With ds.Tables(0)
                    For i As Integer = 0 To .Rows.Count - 1
                        If CInt(.Rows.Item("RejectCode").ToString) = 693 Then   'BIN3 
                            rejectValue = CInt(.Rows.Item("RejectQty").ToString)
                            BIN_Limit = getBINLimit(.Rows.Item("RejectID").ToString, productCode) 'GET BIN3 Limit
                            BIN3 = True
                            BIN4 = False
                        ElseIf CInt(.Rows.Item("RejectCode").ToString) = 694 Then   'BIN4 
                            rejectValue = CInt(.Rows.Item("RejectQty").ToString)
                            BIN_Limit = getBINLimit(.Rows.Item("RejectID").ToString, productCode) 'GET BIN4 Limit
                            BIN3 = False
                            BIN4 = True
                        End If
                    Next

                    If rejectValue <> 0 Then
                        rejectPercentage = Math.Round((rejectValue / QtyIn) * 100, 2)
                        If rejectPercentage > BIN_Limit Then
                            If BIN3 Then
                                HoldReject = "BIN3"
                            ElseIf BIN4 Then
                                HoldReject = "BIN4"
                            End If
                            setupBINLIMIT = BIN_Limit
                            result = rejectPercentage
                        End If
                    End If
                End With
            End If
            sql_handler.CloseConnection()
        End If
        Return result
    End Function

    Public Function getBINLimit(ByVal RejectID As String, ByVal ProductCode As Integer) As Double
        Dim result As Double = 0.0
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT ParameterValue FROM PS_Product_Attribute WHERE ParameterCode=@ParameterCode AND ProductCode=@ProductCode"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@ParameterCode", SqlDbType.Int, IIf(RejectID = "BIN3", 354, 356))
        sql_handler.SetParameterValues(1, "@ProductCode", SqlDbType.Int, ProductCode)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = CDbl(dr.Item("ParameterValue").ToString.Replace("%", ""))
                End If
            End If
            sql_handler.CloseConnection()
        End If
        sql_handler = Nothing
        Return result
    End Function

    Public Function isAutoHoldStation(ByVal RecipeCode As Integer) As Boolean
        Dim result As Boolean = False
        Dim sql_handler As New SQLHandler
        Dim dr As SqlDataReader = Nothing
        Dim sql As String = "SELECT RecipeCode FROM PS_Recipe_AutoHold WHERE RecipeCode=@RecipeCode"
        sql_handler.CreateParameter(1)
        sql_handler.SetParameterValues(0, "@RecipeCode", SqlDbType.Int, RecipeCode)
        If sql_handler.OpenConnection Then
            If sql_handler.FillDataReader(sql, dr, CommandType.Text) Then
                If dr.Read Then
                    result = True
                End If
            End If
            sql_handler.CloseConnection()
        Else
            result = False
        End If
        sql_handler = Nothing
        Return result
    End Function
#End Region

End Class
