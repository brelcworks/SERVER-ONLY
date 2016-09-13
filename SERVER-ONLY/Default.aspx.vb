Imports ClosedXML.Excel
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Web.Configuration

Public Class _Default
    Inherits System.Web.UI.Page
    Private CON_AM As New SqlConnection("Server=7d398a2f-1a2b-4338-bcc7-a66000a64b47.sqlserver.sequelizer.com;Database=db7d398a2f1a2b4338bcc7a66000a64b47;User ID=kjvstqndeaoallkm;Password=xsPrEXzwwVnd4TZxZ2Yag3qZbGjiipdL843dyHbK6AvazBnzikiGKxxCbWq7Nqoh;")
    Private filePath As String
    Private fileStream As FileStream
    Private streamWriter As StreamWriter
    Private CT As Integer = 0
    Private ct1 As Integer = 0
    Private ISCOM As Boolean = True
    Private AD As Boolean = False
    Private LP As Boolean = False
    Private LP1 As Boolean = False
    Private dlrpt As Boolean = False
    Public CON5 As New System.Data.OleDb.OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("\App_Data\ERR\ERR.accdb") & ";Persist Security Info=False;")
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            DLBCKLBL.Text = "TODAY BACKUP WILL DONE AFTER "
            DLRPTLBL.Text = "TODAY REPORT WILL SENT AFTER "
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim ss As Integer = Now.ToString("HH") * 60 * 60 + Now.ToString("mm") * 60 + Now.ToString("ss")
        Dim ss1 As Integer = Now.ToString("ss")
        CT = ss1
        ct1 = ss1

        Try
            If CT = 25 Then
                ISCOM = False
                hr_bck()
            End If
        Catch ex As Exception
            EXLERR(Now.ToString, ex.ToString)
        End Try
        Try
            Dim hr As String = Now.ToString("hh")
            If hr = "10" Then
                WebConfigurationManager.AppSettings.Set("dlrpset", "true")
            End If
            Dim vl As String = WebConfigurationManager.AppSettings("dlrpset")
            If hr = "05" Then
                If vl = "true" Then
                    DLYRPT()
                End If
            End If
        Catch ex As Exception
            EXLERR(Now.ToString, ex.ToString)
        End Try
        Try
            If CT = 20 Then
                LP = False
                rmtrckr()
            End If
        Catch ex As Exception
            EXLERR(Now.ToString, ex.ToString)
        End Try
        Try
            If CT = 45 Then
                If IsLastDay(Today()) = False Then
                    LP1 = False
                    MON_REP()
                End If
            End If
        Catch ex As Exception
            EXLERR(Now.ToString, ex.ToString)
        End Try
        Try
            Dim hr As String = Now.ToString("HH")
            CTSTA.Text = CT
            nt.Text = Now.ToString("dd-MMMM-yyyy hh:mm:ss tt fff") & " Total Seconds of Today is " & ss
        Catch ex As Exception
            EXLERR(Now.ToString, ex.ToString)
        End Try
        Try
            If ct1 = 59 Then
                If CON_AM.State <> ConnectionState.Open Then
                    CON_AM.Open()
                    ERR.Text = "SERVER CONNECTION IS " & CON_AM.State.ToString
                End If
            End If
        Catch ex As Exception
            EXLERR(Now.ToString, ex.ToString)
        End Try
    End Sub
    Protected Sub hr_bck()
        Do Until ISCOM = True
            Try
                If CON_AM.State <> ConnectionState.Open Then CON_AM.Open()
                Dim POP_AM_DA As New SqlDataAdapter("SELECT * FROM BILL1", CON_AM)
                Dim POP_AM_DT As New DataTable
                POP_AM_DA.Fill(POP_AM_DT)
                Dim workbook = New XLWorkbook()
                Dim worksheet = workbook.Worksheets.Add("BILL1")

                Dim colIndex As Integer = 0
                For col As Integer = 1 To POP_AM_DT.Columns.Count
                    worksheet.Cell(1, col).Value = POP_AM_DT.Columns(colIndex).ToString()
                    worksheet.Cell(1, col).Style.Font.Bold = True
                    colIndex += 1
                Next

                colIndex = 0
                Dim rowIndex As Integer = 0
                For row As Integer = 2 To POP_AM_DT.Rows.Count + 1
                    For col As Integer = 1 To POP_AM_DT.Columns.Count
                        worksheet.Cell(row, col).Value = POP_AM_DT.Rows(rowIndex)(colIndex).ToString()
                        colIndex += 1
                    Next
                    colIndex = 0
                    rowIndex += 1
                Next
                Dim PMR_AM_DA As New SqlDataAdapter("SELECT * FROM BILLs", CON_AM)
                Dim PMR_AM_DT As New DataTable
                PMR_AM_DA.Fill(PMR_AM_DT)
                workbook.Worksheets.Add("BILL")
                worksheet = workbook.Worksheet("BILL")
                Dim colIndex1 As Integer = 0
                For col As Integer = 1 To PMR_AM_DT.Columns.Count
                    worksheet.Cell(1, col).Value = PMR_AM_DT.Columns(colIndex1).ToString()
                    worksheet.Cell(1, col).Style.Font.Bold = True
                    colIndex1 += 1
                Next

                colIndex1 = 0
                Dim rowIndex1 As Integer = 0
                For row As Integer = 2 To PMR_AM_DT.Rows.Count + 1
                    For col As Integer = 1 To PMR_AM_DT.Columns.Count
                        worksheet.Cell(row, col).Value = PMR_AM_DT.Rows(rowIndex1)(colIndex1).ToString()
                        colIndex1 += 1
                    Next
                    colIndex1 = 0
                    rowIndex1 += 1
                Next

                Dim PMR_AM_DA1 As New SqlDataAdapter("SELECT * FROM MAINPOPUS", CON_AM)
                Dim PMR_AM_DT1 As New DataTable
                PMR_AM_DA1.Fill(PMR_AM_DT1)
                workbook.Worksheets.Add("POP")
                worksheet = workbook.Worksheet("POP")
                Dim colIndex2 As Integer = 0
                For col As Integer = 1 To PMR_AM_DT1.Columns.Count
                    worksheet.Cell(1, col).Value = PMR_AM_DT1.Columns(colIndex2).ToString()
                    worksheet.Cell(1, col).Style.Font.Bold = True
                    colIndex2 += 1
                Next

                colIndex2 = 0
                Dim rowIndex2 As Integer = 0
                For row As Integer = 2 To PMR_AM_DT1.Rows.Count + 1
                    For col As Integer = 1 To PMR_AM_DT1.Columns.Count
                        worksheet.Cell(row, col).Value = PMR_AM_DT1.Rows(rowIndex2)(colIndex2).ToString()
                        colIndex2 += 1
                    Next
                    colIndex2 = 0
                    rowIndex2 += 1
                Next

                Dim PMR_AM_DA2 As New SqlDataAdapter("SELECT * FROM pmrs", CON_AM)
                Dim PMR_AM_DT2 As New DataTable
                PMR_AM_DA2.Fill(PMR_AM_DT2)
                workbook.Worksheets.Add("PMR")
                worksheet = workbook.Worksheet("PMR")
                Dim colIndex3 As Integer = 0
                For col As Integer = 1 To PMR_AM_DT2.Columns.Count
                    worksheet.Cell(1, col).Value = PMR_AM_DT2.Columns(colIndex3).ToString()
                    worksheet.Cell(1, col).Style.Font.Bold = True
                    colIndex3 += 1
                Next

                colIndex3 = 0
                Dim rowIndex3 As Integer = 0
                For row As Integer = 2 To PMR_AM_DT2.Rows.Count + 1
                    For col As Integer = 1 To PMR_AM_DT2.Columns.Count
                        worksheet.Cell(row, col).Value = PMR_AM_DT2.Rows(rowIndex3)(colIndex3).ToString()
                        colIndex3 += 1
                    Next
                    colIndex3 = 0
                    rowIndex3 += 1
                Next

                Dim X As String = Server.MapPath("\App_Data\BCK\" & Format(Now, "dd-MMMM-yyyy hh-mm-ss-fff tt") & ".xlsx")
                workbook.SaveAs(X)

                Try
                    Dim url1 As String = "https://dav.box.com/dav/ASP"
                    Dim port1 As String = "443"
                    If port1 <> "" Then
                        Dim u As New Uri(url1)
                        Dim host As String = u.Host
                        url1 = url1.Replace(host, host & ":" & port1)
                    End If
                    Dim XY As String = Format(Now, "dd-MMMM-yyyy hh-mm-ss-fff tt")
                    url1 = url1.TrimEnd("/"c) & "/" & XY
                    Dim Request1 As System.Net.HttpWebRequest
                    Request1 = CType(System.Net.WebRequest.Create(url1),
                              System.Net.HttpWebRequest)
                    Request1.Credentials = New NetworkCredential("brelcworks@gmail.com", "Indian123")
                    Request1.Method = "MKCOL"
                    Dim Response1 As System.Net.HttpWebResponse
                    Response1 = CType(Request1.GetResponse(), System.Net.HttpWebResponse)
                    Response1.Close()
                    Dim fileLength As Long = FileIO.FileSystem.GetFileInfo(X).Length
                    Dim url As String = "https://dav.box.com/dav/ASP/" & XY
                    Dim port As String = "443"
                    If port <> "" Then
                        Dim u As New Uri(url)
                        Dim host As String = u.Host
                        url = url.Replace(host, host & ":" & port)
                    End If
                    url = url.TrimEnd("/"c) & "/" & Path.GetFileName(X)
                    Dim request As HttpWebRequest =
                    DirectCast(System.Net.HttpWebRequest.Create(url), HttpWebRequest)
                    request.Credentials = New NetworkCredential("brelcworks@gmail.com", "Indian123")
                    request.Method = WebRequestMethods.Http.Put
                    request.ContentLength = fileLength
                    request.SendChunked = True
                    request.Headers.Add("Translate: f")
                    request.AllowWriteStreamBuffering = True
                    Dim s As IO.Stream = Nothing
                    Try
                        s = request.GetRequestStream()
                    Catch ex As Exception
                        EXLERR(Now.ToString, ex.ToString)
                    End Try
                    Dim fs As New IO.FileStream(X, IO.FileMode.Open, IO.FileAccess.Read)
                    Dim byteTransferRate As Integer = 1024
                    Dim bytes(byteTransferRate - 1) As Byte
                    Dim bytesRead As Integer = 0
                    Dim totalBytesRead As Long = 0
                    Do
                        bytesRead = fs.Read(bytes, 0, bytes.Length)
                        If bytesRead > 0 Then
                            totalBytesRead += bytesRead
                            s.Write(bytes, 0, bytesRead)
                        End If
                    Loop While bytesRead > 0
                    s.Close()
                    s.Dispose()
                    s = Nothing
                    fs.Close()
                    fs.Dispose()
                    fs = Nothing
                    Dim response As HttpWebResponse = Nothing
                    Try
                        response = DirectCast(request.GetResponse(), HttpWebResponse)
                        Dim code As HttpStatusCode = response.StatusCode
                        response.Close()
                        response = Nothing
                        If totalBytesRead = fileLength AndAlso
                        code = HttpStatusCode.Created Then
                            Dim d1 As Date = Today
                        Else
                            MsgBox("The file did not upload successfully.")
                        End If
                    Catch ex As Exception
                        EXLERR(Now.ToString, ex.ToString)
                    End Try
                Catch ex As Exception
                    EXLERR(Now.ToString, ex.ToString)
                End Try
                Dim path1 As String = Server.MapPath("/App_Data/BCK/")

                ERR.Text = "OK"
                ISCOM = True
            Catch ex As Exception
                EXLERR(Now.ToString, ex.ToString)
                err_display(ex.ToString)
                ISCOM = False
            End Try
        Loop
    End Sub
    Private Sub DeleteDirectory(path As String)
        If Directory.Exists(path) Then
            'Delete all files from the Directory
            For Each filepath As String In Directory.GetFiles(path)
                System.IO.File.Delete(filepath)
            Next
        End If
    End Sub
    Public Sub OpenFile1()
        Dim strPath As String
        strPath = Server.MapPath("/App_Data/ERR/log.log")
        If System.IO.File.Exists(strPath) Then
            fileStream = New FileStream(strPath, FileMode.Append, FileAccess.Write)
        Else
            fileStream = New FileStream(strPath, FileMode.Create, FileAccess.Write)
        End If
        streamWriter = New StreamWriter(fileStream)
    End Sub

    Public Sub WriteLog(ByVal strComments As String)
        OpenFile1()
        streamWriter.WriteLine(vbCrLf & "--------------------" & vbCrLf & "--------------------" & vbCrLf & "Error :" & Format(Now(), "dd-MMMM-yyyy hh:mm:ss:fff tt") & " :- " & strComments)
        CloseFile()
    End Sub

    Public Sub CloseFile()
        streamWriter.Close()
        fileStream.Close()
    End Sub
    Private Sub EXLERR(ByVal ETIME As String, ByVal ERR As String)
        If CON5.State <> ConnectionState.Open Then CON5.Open()
        Dim ICMD As New OleDb.OleDbCommand
        ICMD.CommandType = CommandType.Text
        ICMD.CommandText = "INSERT INTO ERR (ETIME, ERR) VALUES (@ETIME, @ERR)"
        ICMD.Parameters.AddWithValue("@ETIME", ETIME)
        ICMD.Parameters.AddWithValue("@ERR", ERR)
        ICMD.Connection = CON5
        ICMD.ExecuteNonQuery()
    End Sub
    Private Sub DLYRPT()
        Do Until AD = True
            Try
                Dim PMR_AM_DA As New SqlDataAdapter("SELECT * FROM PMRs", CON_AM)
                Dim PMR_AM_DT As New DataTable
                PMR_AM_DA.Fill(PMR_AM_DT)
                Dim mon As String = Today.ToString("MMM")
                Dim D1, D2 As Date
                Select Case mon
                    Case "Jan"
                        D1 = "1/1/" & Today.ToString("yyyy")
                        D2 = "1/31/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Feb"
                        D1 = "2/1/" & Today.ToString("yyyy")
                        D2 = "2/28/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Mar"
                        D1 = "3/1/" & Today.ToString("yyyy")
                        D2 = "3/31/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Apr"
                        D1 = "4/1/" & Today.ToString("yyyy")
                        D2 = "4/30/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "May"
                        D1 = "5/1/" & Today.ToString("yyyy")
                        D2 = "5/31/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "June"
                        D1 = "6/1/" & Today.ToString("yyyy")
                        D2 = "6/30/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "July"
                        D1 = "7/1/" & Today.ToString("yyyy")
                        D2 = "7/31/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Aug"
                        D1 = "8/1/" & Today.ToString("yyyy")
                        D2 = "8/31/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Sep"
                        D1 = "9/1/" & Today.ToString("yyyy")
                        D2 = "9/30/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Oct"
                        D1 = "10/1/" & Today.ToString("yyyy")
                        D2 = "10/31/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Nov"
                        D1 = "11/1/" & Today.ToString("yyyy")
                        D2 = "11/30/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Dec"
                        D1 = "12/1/" & Today.ToString("yyyy")
                        D2 = "12/31/" & Today.ToString("yyyy") & " 23:59:59"
                End Select
                Dim TOD1 As Date = Today()
                Dim TS As TimeSpan = New TimeSpan(23, 59, 59)
                Dim TOD2 As Date = TOD1.Add(TS)
                Dim dv As New DataView(PMR_AM_DT)
                dv.RowFilter = "stype='OIL SERVICE' AND DOS >= #" & TOD1 & "# and DOS <= #" & TOD2 & "#"
                Dim PM_COUNT As Integer = dv.Count
                Dim workbook = New XLWorkbook()
                Dim xlWorkSheet = workbook.Worksheets.Add("PM_REPORT")
                xlWorkSheet.Cell(1, 1).Value = "Sl. No"
                xlWorkSheet.Cell(1, 2).Value = "Customer"
                xlWorkSheet.Cell(1, 3).Value = "AMC / Out of Scope"
                xlWorkSheet.Cell(1, 4).Value = "DT. of Complaint"
                xlWorkSheet.Cell(1, 5).Value = "Time of Compliant"
                xlWorkSheet.Cell(1, 6).Value = "Complaint / Service No."
                xlWorkSheet.Cell(1, 7).Value = "Site ID"
                xlWorkSheet.Cell(1, 8).Value = "Site Name"
                xlWorkSheet.Cell(1, 9).Value = "District"
                xlWorkSheet.Cell(1, 10).Value = "State"
                xlWorkSheet.Cell(1, 11).Value = "Complaint Category"
                xlWorkSheet.Cell(1, 12).Value = "ESN"
                xlWorkSheet.Cell(1, 13).Value = "ENG Model"
                xlWorkSheet.Cell(1, 14).Value = "KVA"
                xlWorkSheet.Cell(1, 15).Value = "DOI"
                xlWorkSheet.Cell(1, 16).Value = "Genset NO"
                xlWorkSheet.Cell(1, 17).Value = "ALT. MAKE"
                xlWorkSheet.Cell(1, 18).Value = "ALT. SR NO"
                xlWorkSheet.Cell(1, 19).Value = "BATTERY SR NO"
                xlWorkSheet.Cell(1, 20).Value = "HMR"
                xlWorkSheet.Cell(1, 21).Value = "Nature of Complaint"
                xlWorkSheet.Cell(1, 22).Value = "Severity"
                xlWorkSheet.Cell(1, 23).Value = "Reason for failure"
                xlWorkSheet.Cell(1, 24).Value = "Status"
                xlWorkSheet.Cell(1, 25).Value = "Complaint Closure Date"
                xlWorkSheet.Cell(1, 26).Value = "Compliant closure time"
                xlWorkSheet.Cell(1, 27).Value = "(U-W).Value-(O-W).Value"
                xlWorkSheet.Cell(1, 28).Value = "Action Taken"
                xlWorkSheet.Cell(1, 29).Value = "Material Changed"
                xlWorkSheet.Cell(1, 30).Value = "Service Dealer"
                xlWorkSheet.Cell(1, 31).Value = "Service Tech. Name"
                xlWorkSheet.Cell(1, 32).Value = "Supplied By"
                xlWorkSheet.Cell(1, 33).Value = "AMC Status"
                xlWorkSheet.Cell(1, 34).Value = "TTR"
                xlWorkSheet.Cell(1, 35).Value = "OSLA/WSLA"
                xlWorkSheet.Cell(1, 36).Value = "Time Slot"
                xlWorkSheet.Cell(1, 37).Value = "REASON FOR EXCEEDING SLA TIME"
                For i As Integer = 0 To dv.Count - 1
                    xlWorkSheet.Cell(i + 2, 1).Value = i.ToString + 1
                    xlWorkSheet.Cell(i + 2, 2).Value = dv(i)("CUST")
                    xlWorkSheet.Cell(i + 2, 3).Value = dv(i)("stype")
                    Dim DT1 As DateTime
                    If Not IsDBNull(dv(i)("DOS")) Then
                        DT1 = dv(i)("DOS")
                        xlWorkSheet.Cell(i + 2, 4).Value = DT1.ToString("dd-MMM-yyyy")
                        xlWorkSheet.Cell(i + 2, 5).Value = DT1.ToString("hh:mm tt")
                    End If
                    xlWorkSheet.Cell(i + 2, 6).Value = dv(i)("recid")
                    xlWorkSheet.Cell(i + 2, 7).Value = dv(i)("sid")
                    xlWorkSheet.Cell(i + 2, 8).Value = dv(i)("sname")
                    xlWorkSheet.Cell(i + 2, 9).Value = dv(i)("dist")
                    xlWorkSheet.Cell(i + 2, 10).Value = dv(i)("state")
                    xlWorkSheet.Cell(i + 2, 11).Value = dv(i)("ccate")
                    xlWorkSheet.Cell(i + 2, 12).Value = dv(i)("engine_no")
                    xlWorkSheet.Cell(i + 2, 13).Value = dv(i)("model")
                    xlWorkSheet.Cell(i + 2, 14).Value = dv(i)("KVA")
                    xlWorkSheet.Cell(i + 2, 15).Value = dv(i)("DOI")
                    xlWorkSheet.Cell(i + 2, 16).Value = dv(i)("DGNO")
                    xlWorkSheet.Cell(i + 2, 17).Value = dv(i)("AMAKE")
                    xlWorkSheet.Cell(i + 2, 18).Value = dv(i)("ALSN")
                    xlWorkSheet.Cell(i + 2, 19).Value = dv(i)("BSN")
                    xlWorkSheet.Cell(i + 2, 20).Value = dv(i)("HMR")
                    xlWorkSheet.Cell(i + 2, 21).Value = dv(i)("CNAT")
                    xlWorkSheet.Cell(i + 2, 22).Value = dv(i)("SERV")
                    xlWorkSheet.Cell(i + 2, 23).Value = dv(i)("RFAIL")
                    xlWorkSheet.Cell(i + 2, 24).Value = dv(i)("STA")
                    Dim DT2 As DateTime
                    If Not IsDBNull(dv(i)("CDATI")) Then
                        DT2 = dv(i)("CDATI")
                        xlWorkSheet.Cell(i + 2, 25).Value = DT2.ToString("dd-MMM-yyyy")
                        xlWorkSheet.Cell(i + 2, 26).Value = DT2.ToString("hh:mm tt")
                    End If
                    xlWorkSheet.Cell(i + 2, 27).Value = dv(i)("warr")
                    xlWorkSheet.Cell(i + 2, 28).Value = dv(i)("action")
                    xlWorkSheet.Cell(i + 2, 29).Value = dv(i)("meterial")
                    xlWorkSheet.Cell(i + 2, 30).Value = "A1587"
                    xlWorkSheet.Cell(i + 2, 31).Value = dv(i)("technician")
                    xlWorkSheet.Cell(i + 2, 32).Value = dv(i)("oea")
                    xlWorkSheet.Cell(i + 2, 33).Value = dv(i)("amc")
                    xlWorkSheet.Cell(i + 2, 34).Value = dv(i)("ttr")
                    xlWorkSheet.Cell(i + 2, 35).Value = dv(i)("sla")
                    xlWorkSheet.Cell(i + 2, 36).Value = dv(i)("tslot")
                    xlWorkSheet.Cell(i + 2, 37).Value = dv(i)("resla")
                Next
                xlWorkSheet.Range("D2", "D" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy")
                xlWorkSheet.Range("E2", "E" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("hh:mm AM/PM")
                xlWorkSheet.Range("O2", "O" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy hh:mm:ss")
                xlWorkSheet.Range("Y2", "Y" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy")
                xlWorkSheet.Range("z2", "z" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("hh:mm AM/PM")
                xlWorkSheet.Range("L2", "L" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("#")
                xlWorkSheet.Columns().AdjustToContents()
                xlWorkSheet.SheetView.FreezeRows(1)
                Dim range As ClosedXML.Excel.IXLRange = xlWorkSheet.RangeUsed()
                Dim RCNT As String = "AK" & range.RowCount()
                xlWorkSheet.Range("a1", RCNT).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", "AK1").Style.Fill.BackgroundColor = XLColor.Turquoise

                workbook.Worksheets.Add("CM_REPORT")
                xlWorkSheet = workbook.Worksheet("CM_REPORT")
                dv.RowFilter = "stype='CUSTOMER COMPLAINT' AND DOS >= #" & TOD1 & "# and DOS <= #" & TOD2 & "#"
                Dim CM_COUNT As Integer = dv.Count
                xlWorkSheet.Cell(1, 1).Value = "Sl. No"
                xlWorkSheet.Cell(1, 2).Value = "Customer"
                xlWorkSheet.Cell(1, 3).Value = "AMC / Out of Scope"
                xlWorkSheet.Cell(1, 4).Value = "DT. of Complaint"
                xlWorkSheet.Cell(1, 5).Value = "Time of Compliant"
                xlWorkSheet.Cell(1, 6).Value = "Complaint / Service No."
                xlWorkSheet.Cell(1, 7).Value = "Site ID"
                xlWorkSheet.Cell(1, 8).Value = "Site Name"
                xlWorkSheet.Cell(1, 9).Value = "District"
                xlWorkSheet.Cell(1, 10).Value = "State"
                xlWorkSheet.Cell(1, 11).Value = "Complaint Category"
                xlWorkSheet.Cell(1, 12).Value = "ESN"
                xlWorkSheet.Cell(1, 13).Value = "ENG Model"
                xlWorkSheet.Cell(1, 14).Value = "KVA"
                xlWorkSheet.Cell(1, 15).Value = "DOI"
                xlWorkSheet.Cell(1, 16).Value = "Genset NO"
                xlWorkSheet.Cell(1, 17).Value = "ALT. MAKE"
                xlWorkSheet.Cell(1, 18).Value = "ALT. SR NO"
                xlWorkSheet.Cell(1, 19).Value = "BATTERY SR NO"
                xlWorkSheet.Cell(1, 20).Value = "HMR"
                xlWorkSheet.Cell(1, 21).Value = "Nature of Complaint"
                xlWorkSheet.Cell(1, 22).Value = "Severity"
                xlWorkSheet.Cell(1, 23).Value = "Reason for failure"
                xlWorkSheet.Cell(1, 24).Value = "Status"
                xlWorkSheet.Cell(1, 25).Value = "Complaint Closure Date"
                xlWorkSheet.Cell(1, 26).Value = "Compliant closure time"
                xlWorkSheet.Cell(1, 27).Value = "(U-W).Value-(O-W).Value"
                xlWorkSheet.Cell(1, 28).Value = "Action Taken"
                xlWorkSheet.Cell(1, 29).Value = "Material Changed"
                xlWorkSheet.Cell(1, 30).Value = "Service Dealer"
                xlWorkSheet.Cell(1, 31).Value = "Service Tech. Name"
                xlWorkSheet.Cell(1, 32).Value = "Supplied By"
                xlWorkSheet.Cell(1, 33).Value = "AMC Status"
                xlWorkSheet.Cell(1, 34).Value = "TTR"
                xlWorkSheet.Cell(1, 35).Value = "OSLA/WSLA"
                xlWorkSheet.Cell(1, 36).Value = "Time Slot"
                xlWorkSheet.Cell(1, 37).Value = "REASON FOR EXCEEDING SLA TIME"
                For i As Integer = 0 To dv.Count - 1
                    xlWorkSheet.Cell(i + 2, 1).Value = i.ToString + 1
                    xlWorkSheet.Cell(i + 2, 2).Value = dv(i)("CUST")
                    xlWorkSheet.Cell(i + 2, 3).Value = dv(i)("stype")
                    Dim DT1 As DateTime
                    If Not IsDBNull(dv(i)("DOS")) Then
                        DT1 = dv(i)("DOS")
                        xlWorkSheet.Cell(i + 2, 4).Value = DT1.ToString("dd-MMM-yyyy")
                        xlWorkSheet.Cell(i + 2, 5).Value = DT1.ToString("hh:mm tt")
                    End If
                    xlWorkSheet.Cell(i + 2, 6).Value = dv(i)("recid")
                    xlWorkSheet.Cell(i + 2, 7).Value = dv(i)("sid")
                    xlWorkSheet.Cell(i + 2, 8).Value = dv(i)("sname")
                    xlWorkSheet.Cell(i + 2, 9).Value = dv(i)("dist")
                    xlWorkSheet.Cell(i + 2, 10).Value = dv(i)("state")
                    xlWorkSheet.Cell(i + 2, 11).Value = dv(i)("ccate")
                    xlWorkSheet.Cell(i + 2, 12).Value = dv(i)("engine_no")
                    xlWorkSheet.Cell(i + 2, 13).Value = dv(i)("model")
                    xlWorkSheet.Cell(i + 2, 14).Value = dv(i)("KVA")
                    xlWorkSheet.Cell(i + 2, 15).Value = dv(i)("DOI")
                    xlWorkSheet.Cell(i + 2, 16).Value = dv(i)("DGNO")
                    xlWorkSheet.Cell(i + 2, 17).Value = dv(i)("AMAKE")
                    xlWorkSheet.Cell(i + 2, 18).Value = dv(i)("ALSN")
                    xlWorkSheet.Cell(i + 2, 19).Value = dv(i)("BSN")
                    xlWorkSheet.Cell(i + 2, 20).Value = dv(i)("HMR")
                    xlWorkSheet.Cell(i + 2, 21).Value = dv(i)("CNAT")
                    xlWorkSheet.Cell(i + 2, 22).Value = dv(i)("SERV")
                    xlWorkSheet.Cell(i + 2, 23).Value = dv(i)("RFAIL")
                    xlWorkSheet.Cell(i + 2, 24).Value = dv(i)("STA")
                    Dim DT2 As DateTime
                    If Not IsDBNull(dv(i)("CDATI")) Then
                        DT2 = dv(i)("CDATI")
                        xlWorkSheet.Cell(i + 2, 25).Value = DT2.ToString("dd-MMM-yyyy")
                        xlWorkSheet.Cell(i + 2, 26).Value = DT2.ToString("hh:mm tt")
                    End If
                    xlWorkSheet.Cell(i + 2, 27).Value = dv(i)("warr")
                    xlWorkSheet.Cell(i + 2, 28).Value = dv(i)("action")
                    xlWorkSheet.Cell(i + 2, 29).Value = dv(i)("meterial")
                    xlWorkSheet.Cell(i + 2, 30).Value = "A1587"
                    xlWorkSheet.Cell(i + 2, 31).Value = dv(i)("technician")
                    xlWorkSheet.Cell(i + 2, 32).Value = dv(i)("oea")
                    xlWorkSheet.Cell(i + 2, 33).Value = dv(i)("amc")
                    xlWorkSheet.Cell(i + 2, 34).Value = dv(i)("ttr")
                    xlWorkSheet.Cell(i + 2, 35).Value = dv(i)("sla")
                    xlWorkSheet.Cell(i + 2, 36).Value = dv(i)("tslot")
                    xlWorkSheet.Cell(i + 2, 37).Value = dv(i)("resla")
                Next
                xlWorkSheet.Range("D2", "D" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy")
                xlWorkSheet.Range("E2", "E" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("hh:mm AM/PM")
                xlWorkSheet.Range("O2", "O" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy hh:mm:ss")
                xlWorkSheet.Range("Y2", "Y" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy")
                xlWorkSheet.Range("z2", "z" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("hh:mm AM/PM")
                xlWorkSheet.Range("L2", "L" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("#")
                xlWorkSheet.Columns().AdjustToContents()
                xlWorkSheet.SheetView.FreezeRows(1)
                range = xlWorkSheet.RangeUsed()
                RCNT = "AK" & range.RowCount()
                xlWorkSheet.Range("a1", RCNT).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", "AK1").Style.Fill.BackgroundColor = XLColor.Turquoise

                workbook.Worksheets.Add("VISIT_REPORT")
                xlWorkSheet = workbook.Worksheet("VISIT_REPORT")
                dv.RowFilter = "stype='PM VISIT' AND DOS >= #" & TOD1 & "# and DOS <= #" & TOD2 & "#"
                Dim VS_COUNT As Integer = dv.Count
                xlWorkSheet.Cell(1, 1).Value = "Sl. No"
                xlWorkSheet.Cell(1, 2).Value = "Customer"
                xlWorkSheet.Cell(1, 3).Value = "AMC / Out of Scope"
                xlWorkSheet.Cell(1, 4).Value = "DT. of Complaint"
                xlWorkSheet.Cell(1, 5).Value = "Time of Compliant"
                xlWorkSheet.Cell(1, 6).Value = "Complaint / Service No."
                xlWorkSheet.Cell(1, 7).Value = "Site ID"
                xlWorkSheet.Cell(1, 8).Value = "Site Name"
                xlWorkSheet.Cell(1, 9).Value = "District"
                xlWorkSheet.Cell(1, 10).Value = "State"
                xlWorkSheet.Cell(1, 11).Value = "Complaint Category"
                xlWorkSheet.Cell(1, 12).Value = "ESN"
                xlWorkSheet.Cell(1, 13).Value = "ENG Model"
                xlWorkSheet.Cell(1, 14).Value = "KVA"
                xlWorkSheet.Cell(1, 15).Value = "DOI"
                xlWorkSheet.Cell(1, 16).Value = "Genset NO"
                xlWorkSheet.Cell(1, 17).Value = "ALT. MAKE"
                xlWorkSheet.Cell(1, 18).Value = "ALT. SR NO"
                xlWorkSheet.Cell(1, 19).Value = "BATTERY SR NO"
                xlWorkSheet.Cell(1, 20).Value = "HMR"
                xlWorkSheet.Cell(1, 21).Value = "Nature of Complaint"
                xlWorkSheet.Cell(1, 22).Value = "Severity"
                xlWorkSheet.Cell(1, 23).Value = "Reason for failure"
                xlWorkSheet.Cell(1, 24).Value = "Status"
                xlWorkSheet.Cell(1, 25).Value = "Complaint Closure Date"
                xlWorkSheet.Cell(1, 26).Value = "Compliant closure time"
                xlWorkSheet.Cell(1, 27).Value = "(U-W).Value-(O-W).Value"
                xlWorkSheet.Cell(1, 28).Value = "Action Taken"
                xlWorkSheet.Cell(1, 29).Value = "Material Changed"
                xlWorkSheet.Cell(1, 30).Value = "Service Dealer"
                xlWorkSheet.Cell(1, 31).Value = "Service Tech. Name"
                xlWorkSheet.Cell(1, 32).Value = "Supplied By"
                xlWorkSheet.Cell(1, 33).Value = "AMC Status"
                xlWorkSheet.Cell(1, 34).Value = "TTR"
                xlWorkSheet.Cell(1, 35).Value = "OSLA/WSLA"
                xlWorkSheet.Cell(1, 36).Value = "Time Slot"
                xlWorkSheet.Cell(1, 37).Value = "REASON FOR EXCEEDING SLA TIME"
                For i As Integer = 0 To dv.Count - 1
                    xlWorkSheet.Cell(i + 2, 1).Value = i.ToString + 1
                    xlWorkSheet.Cell(i + 2, 2).Value = dv(i)("CUST")
                    xlWorkSheet.Cell(i + 2, 3).Value = dv(i)("stype")
                    Dim DT1 As DateTime
                    If Not IsDBNull(dv(i)("DOS")) Then
                        DT1 = dv(i)("DOS")
                        xlWorkSheet.Cell(i + 2, 4).Value = DT1.ToString("dd-MMM-yyyy")
                        xlWorkSheet.Cell(i + 2, 5).Value = DT1.ToString("hh:mm tt")
                    End If
                    xlWorkSheet.Cell(i + 2, 6).Value = dv(i)("recid")
                    xlWorkSheet.Cell(i + 2, 7).Value = dv(i)("sid")
                    xlWorkSheet.Cell(i + 2, 8).Value = dv(i)("sname")
                    xlWorkSheet.Cell(i + 2, 9).Value = dv(i)("dist")
                    xlWorkSheet.Cell(i + 2, 10).Value = dv(i)("state")
                    xlWorkSheet.Cell(i + 2, 11).Value = dv(i)("ccate")
                    xlWorkSheet.Cell(i + 2, 12).Value = dv(i)("engine_no")
                    xlWorkSheet.Cell(i + 2, 13).Value = dv(i)("model")
                    xlWorkSheet.Cell(i + 2, 14).Value = dv(i)("KVA")
                    xlWorkSheet.Cell(i + 2, 15).Value = dv(i)("DOI")
                    xlWorkSheet.Cell(i + 2, 16).Value = dv(i)("DGNO")
                    xlWorkSheet.Cell(i + 2, 17).Value = dv(i)("AMAKE")
                    xlWorkSheet.Cell(i + 2, 18).Value = dv(i)("ALSN")
                    xlWorkSheet.Cell(i + 2, 19).Value = dv(i)("BSN")
                    xlWorkSheet.Cell(i + 2, 20).Value = dv(i)("HMR")
                    xlWorkSheet.Cell(i + 2, 21).Value = dv(i)("CNAT")
                    xlWorkSheet.Cell(i + 2, 22).Value = dv(i)("SERV")
                    xlWorkSheet.Cell(i + 2, 23).Value = dv(i)("RFAIL")
                    xlWorkSheet.Cell(i + 2, 24).Value = dv(i)("STA")
                    Dim DT2 As DateTime
                    If Not IsDBNull(dv(i)("CDATI")) Then
                        DT2 = dv(i)("CDATI")
                        xlWorkSheet.Cell(i + 2, 25).Value = DT2.ToString("dd-MMM-yyyy")
                        xlWorkSheet.Cell(i + 2, 26).Value = DT2.ToString("hh:mm tt")
                    End If
                    xlWorkSheet.Cell(i + 2, 27).Value = dv(i)("warr")
                    xlWorkSheet.Cell(i + 2, 28).Value = dv(i)("action")
                    xlWorkSheet.Cell(i + 2, 29).Value = dv(i)("meterial")
                    xlWorkSheet.Cell(i + 2, 30).Value = "A1587"
                    xlWorkSheet.Cell(i + 2, 31).Value = dv(i)("technician")
                    xlWorkSheet.Cell(i + 2, 32).Value = dv(i)("oea")
                    xlWorkSheet.Cell(i + 2, 33).Value = dv(i)("amc")
                    xlWorkSheet.Cell(i + 2, 34).Value = dv(i)("ttr")
                    xlWorkSheet.Cell(i + 2, 35).Value = dv(i)("sla")
                    xlWorkSheet.Cell(i + 2, 36).Value = dv(i)("tslot")
                    xlWorkSheet.Cell(i + 2, 37).Value = dv(i)("resla")
                Next
                xlWorkSheet.Range("D2", "D" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy")
                xlWorkSheet.Range("E2", "E" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("hh:mm AM/PM")
                xlWorkSheet.Range("O2", "O" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy hh:mm:ss")
                xlWorkSheet.Range("Y2", "Y" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy")
                xlWorkSheet.Range("z2", "z" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("hh:mm AM/PM")
                xlWorkSheet.Range("L2", "L" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("#")
                xlWorkSheet.Columns().AdjustToContents()
                xlWorkSheet.SheetView.FreezeRows(1)
                range = xlWorkSheet.RangeUsed()
                RCNT = "AK" & range.RowCount()
                xlWorkSheet.Range("a1", RCNT).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", "AK1").Style.Fill.BackgroundColor = XLColor.Turquoise

                Dim excelfile1 As String = Server.MapPath("\App_Data\DATA\" & "DAILY REPORT" & Format(Today, "dd-MMM-yyyy") & ".xlsx")
                workbook.SaveAs(excelfile1)

                Dim mail As New MailMessage
                mail.Subject = "DAILY REPORT ON " & Format(Today, "dd-MMM-yyyy")
                mail.To.Add("pathllk3@gmail.com")
                mail.CC.Add("brelcworks@YAHOO.COM")
                mail.From = New MailAddress("brelcworks@gmail.com")
                Dim MBODY As String = "Dear Sir," & vbCrLf & vbTab & "Today Report is as Given Below !" & vbCrLf & vbCrLf _
                    & "TOTAL PM COUNT:" & vbTab & vbTab & PM_COUNT & vbCrLf _
                    & "TOTAL CM COUNT:" & vbTab & vbTab & CM_COUNT & vbCrLf _
                    & "TOTAL VISIT COUNT:" & vbTab & VS_COUNT & vbCrLf _
                    & "DETAILS ARE IN THE ATTACHMENT" & vbCrLf _
                    & "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ " & vbCrLf _
                    & "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ " & vbCrLf _
                    & vbCrLf & "Thanks & Regards" & vbCrLf & "Anjan Paul" & vbCrLf & "For, B & R Electrical Works"
                mail.Body = MBODY
                Dim attach As New Attachment(excelfile1)
                mail.Attachments.Add(attach)
                Dim smtp As New SmtpClient("smtp.gmail.com")
                smtp.EnableSsl = True
                smtp.Credentials = New System.Net.NetworkCredential("brelcworks", "ratanbose")
                smtp.Port = "587"
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess Or DeliveryNotificationOptions.OnFailure
                smtp.Send(mail)
                EXLERR(Now.ToString, "DAILY REPORT SENT")
                DLRPTLBL.Text = "DAILY REPORT SENT"
                WebConfigurationManager.AppSettings.Set("dlrpset", "false")
                AD = True
            Catch ex As Exception
                AD = False
                EXLERR(Now.ToString, ex.ToString)
                err_display(ex.ToString)
            End Try
        Loop
    End Sub
    Private Sub rmtrckr()
        Do Until LP = True
            Try
                Dim PMR_AM_DA As New SqlDataAdapter("SELECT * FROM PMRs", CON_AM)
                Dim PMR_AM_DT As New DataTable
                PMR_AM_DA.Fill(PMR_AM_DT)
                Dim mon As String = Today.ToString("MMM")
                Dim D1, D2 As Date
                Select Case mon
                    Case "Jan"
                        D1 = "12/26/" & (Today.ToString("yyyy") - 1)
                        D2 = "1/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Feb"
                        D1 = "1/26/" & Today.ToString("yyyy")
                        D2 = "2/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Mar"
                        D1 = "2/26/" & Today.ToString("yyyy")
                        D2 = "3/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Apr"
                        D1 = "3/26/" & Today.ToString("yyyy")
                        D2 = "4/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "May"
                        D1 = "4/26/" & Today.ToString("yyyy")
                        D2 = "5/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "June"
                        D1 = "5/26/" & Today.ToString("yyyy")
                        D2 = "6/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "July"
                        D1 = "6/26/" & Today.ToString("yyyy")
                        D2 = "7/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Aug"
                        D1 = "7/26/" & Today.ToString("yyyy")
                        D2 = "8/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Sep"
                        D1 = "8/26/" & Today.ToString("yyyy")
                        D2 = "9/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Oct"
                        D1 = "9/26/" & Today.ToString("yyyy")
                        D2 = "10/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Nov"
                        D1 = "10/26/" & Today.ToString("yyyy")
                        D2 = "11/25/" & Today.ToString("yyyy") & " 23:59:59"
                    Case "Dec"
                        D1 = "11/26/" & Today.ToString("yyyy")
                        D2 = "12/25/" & Today.ToString("yyyy") & " 23:59:59"
                End Select
                Dim dv As New DataView(PMR_AM_DT)
                dv.RowFilter = "DOS >= #" & D1 & "# and DOS <= #" & D2 & "#"
                Dim PM_COUNT As Integer = dv.Count
                Dim workbook = New XLWorkbook()
                Dim xlWorkSheet = workbook.Worksheets.Add("RM_TRACKER")
                xlWorkSheet.Cell(1, 1).Value = "Sl. No"
                xlWorkSheet.Cell(1, 2).Value = "Customer"
                xlWorkSheet.Cell(1, 3).Value = "AMC / Out of Scope"
                xlWorkSheet.Cell(1, 4).Value = "DT. of Complaint"
                xlWorkSheet.Cell(1, 5).Value = "Time of Compliant"
                xlWorkSheet.Cell(1, 6).Value = "Complaint / Service No."
                xlWorkSheet.Cell(1, 7).Value = "Site ID"
                xlWorkSheet.Cell(1, 8).Value = "Site Name"
                xlWorkSheet.Cell(1, 9).Value = "District"
                xlWorkSheet.Cell(1, 10).Value = "State"
                xlWorkSheet.Cell(1, 11).Value = "Complaint Category"
                xlWorkSheet.Cell(1, 12).Value = "ESN"
                xlWorkSheet.Cell(1, 13).Value = "ENG Model"
                xlWorkSheet.Cell(1, 14).Value = "KVA"
                xlWorkSheet.Cell(1, 15).Value = "DOI"
                xlWorkSheet.Cell(1, 16).Value = "Genset NO"
                xlWorkSheet.Cell(1, 17).Value = "ALT. MAKE"
                xlWorkSheet.Cell(1, 18).Value = "ALT. SR NO"
                xlWorkSheet.Cell(1, 19).Value = "BATTERY SR NO"
                xlWorkSheet.Cell(1, 20).Value = "HMR"
                xlWorkSheet.Cell(1, 21).Value = "Nature of Complaint"
                xlWorkSheet.Cell(1, 22).Value = "Severity"
                xlWorkSheet.Cell(1, 23).Value = "Reason for failure"
                xlWorkSheet.Cell(1, 24).Value = "Status"
                xlWorkSheet.Cell(1, 25).Value = "Complaint Closure Date"
                xlWorkSheet.Cell(1, 26).Value = "Compliant closure time"
                xlWorkSheet.Cell(1, 27).Value = "(U-W).Value-(O-W).Value"
                xlWorkSheet.Cell(1, 28).Value = "Action Taken"
                xlWorkSheet.Cell(1, 29).Value = "Material Changed"
                xlWorkSheet.Cell(1, 30).Value = "Service Dealer"
                xlWorkSheet.Cell(1, 31).Value = "Service Tech. Name"
                xlWorkSheet.Cell(1, 32).Value = "Supplied By"
                xlWorkSheet.Cell(1, 33).Value = "AMC Status"
                xlWorkSheet.Cell(1, 34).Value = "TTR"
                xlWorkSheet.Cell(1, 35).Value = "OSLA/WSLA"
                xlWorkSheet.Cell(1, 36).Value = "Time Slot"
                xlWorkSheet.Cell(1, 37).Value = "REASON FOR EXCEEDING SLA TIME"
                For i As Integer = 0 To dv.Count - 1
                    xlWorkSheet.Cell(i + 2, 1).Value = i.ToString + 1
                    xlWorkSheet.Cell(i + 2, 2).Value = dv(i)("CUST")
                    xlWorkSheet.Cell(i + 2, 3).Value = dv(i)("stype")
                    Dim DT1 As DateTime
                    If Not IsDBNull(dv(i)("DOS")) Then
                        DT1 = dv(i)("DOS")
                        xlWorkSheet.Cell(i + 2, 4).Value = DT1.ToString("dd-MMM-yyyy")
                        xlWorkSheet.Cell(i + 2, 5).Value = DT1.ToString("hh:mm tt")
                    End If
                    xlWorkSheet.Cell(i + 2, 6).Value = dv(i)("recid")
                    xlWorkSheet.Cell(i + 2, 7).Value = dv(i)("sid")
                    xlWorkSheet.Cell(i + 2, 8).Value = dv(i)("sname")
                    xlWorkSheet.Cell(i + 2, 9).Value = dv(i)("dist")
                    xlWorkSheet.Cell(i + 2, 10).Value = dv(i)("state")
                    xlWorkSheet.Cell(i + 2, 11).Value = dv(i)("ccate")
                    xlWorkSheet.Cell(i + 2, 12).Value = dv(i)("engine_no")
                    xlWorkSheet.Cell(i + 2, 13).Value = dv(i)("model")
                    xlWorkSheet.Cell(i + 2, 14).Value = dv(i)("KVA")
                    xlWorkSheet.Cell(i + 2, 15).Value = dv(i)("DOI")
                    xlWorkSheet.Cell(i + 2, 16).Value = dv(i)("DGNO")
                    xlWorkSheet.Cell(i + 2, 17).Value = dv(i)("AMAKE")
                    xlWorkSheet.Cell(i + 2, 18).Value = dv(i)("ALSN")
                    xlWorkSheet.Cell(i + 2, 19).Value = dv(i)("BSN")
                    xlWorkSheet.Cell(i + 2, 20).Value = dv(i)("HMR")
                    xlWorkSheet.Cell(i + 2, 21).Value = dv(i)("CNAT")
                    xlWorkSheet.Cell(i + 2, 22).Value = dv(i)("SERV")
                    xlWorkSheet.Cell(i + 2, 23).Value = dv(i)("RFAIL")
                    xlWorkSheet.Cell(i + 2, 24).Value = dv(i)("STA")
                    Dim DT2 As DateTime
                    If Not IsDBNull(dv(i)("CDATI")) Then
                        DT2 = dv(i)("CDATI")
                        xlWorkSheet.Cell(i + 2, 25).Value = DT2.ToString("dd-MMM-yyyy")
                        xlWorkSheet.Cell(i + 2, 26).Value = DT2.ToString("hh:mm tt")
                    End If
                    xlWorkSheet.Cell(i + 2, 27).Value = dv(i)("warr")
                    xlWorkSheet.Cell(i + 2, 28).Value = dv(i)("action")
                    xlWorkSheet.Cell(i + 2, 29).Value = dv(i)("meterial")
                    xlWorkSheet.Cell(i + 2, 30).Value = "A1587"
                    xlWorkSheet.Cell(i + 2, 31).Value = dv(i)("technician")
                    xlWorkSheet.Cell(i + 2, 32).Value = dv(i)("oea")
                    xlWorkSheet.Cell(i + 2, 33).Value = dv(i)("amc")
                    xlWorkSheet.Cell(i + 2, 34).Value = dv(i)("ttr")
                    xlWorkSheet.Cell(i + 2, 35).Value = dv(i)("sla")
                    xlWorkSheet.Cell(i + 2, 36).Value = dv(i)("tslot")
                    xlWorkSheet.Cell(i + 2, 37).Value = dv(i)("resla")
                Next
                xlWorkSheet.Range("D2", "D" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy")
                xlWorkSheet.Range("E2", "E" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("hh:mm AM/PM")
                xlWorkSheet.Range("O2", "O" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy hh:mm:ss")
                xlWorkSheet.Range("Y2", "Y" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("dd-MMM-yyyy")
                xlWorkSheet.Range("z2", "z" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("hh:mm AM/PM")
                xlWorkSheet.Range("L2", "L" & xlWorkSheet.RangeUsed().RowCount()).Style.NumberFormat.SetFormat("#")
                xlWorkSheet.Columns().AdjustToContents()
                xlWorkSheet.SheetView.FreezeRows(1)
                Dim range As ClosedXML.Excel.IXLRange = xlWorkSheet.RangeUsed()
                Dim RCNT As String = "AK" & range.RowCount()
                xlWorkSheet.Range("a1", RCNT).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", RCNT).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("a1", "AK1").Style.Fill.BackgroundColor = XLColor.Turquoise
                Dim excelfile1 As String = Server.MapPath("\App_Data\DATA\" & "RM TRACKER" & Format(Today, "dd-MMM-yyyy") & ".xlsx")
                workbook.SaveAs(excelfile1)
                Dim mail As New MailMessage
                mail.Subject = "RM TRACKER FOR THE MONTH OF " & Now.ToString("MMM") & " - " & Now.ToString("yyyy")
                mail.To.Add("pathllk3@gmail.com")
                mail.CC.Add("brelcworks@YAHOO.COM")
                mail.From = New MailAddress("brelcworks@gmail.com")
                mail.Body = "Dear Sir," & vbCrLf & vbTab & "Please Find The RM Tracker Sheet in The Attachment." & vbCrLf & vbCrLf & vbCrLf & "Thanks & Regards" & vbCrLf & "Anjan Paul" & vbCrLf & "For, B & R Electrical Works"
                Dim attach As New Attachment(excelfile1)
                mail.Attachments.Add(attach)
                Dim smtp As New SmtpClient("smtp.gmail.com")
                smtp.EnableSsl = True
                smtp.Credentials = New System.Net.NetworkCredential("brelcworks", "ratanbose")
                smtp.Port = "587"
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess Or DeliveryNotificationOptions.OnFailure

                EXLERR(Now.ToString, "RM TRACKER SENT")
                ERR.Text = "RM TRACKER SENT"
                LP = True
            Catch ex As Exception
                LP = False
                EXLERR(Now.ToString, ex.ToString)
                err_display(ex.ToString)
            End Try
        Loop
    End Sub
    Private Sub MON_REP()
        Do Until LP1 = True
            Try
                If CON_AM.State <> ConnectionState.Open Then CON_AM.Open()
                Dim workbook = New XLWorkbook()
                Dim xlWorkSheet = workbook.Worksheets.Add("DASHBOARD")
                Dim PMR_AM_DA As New SqlDataAdapter("SELECT RAT_PH FROM MAINPOPUS GROUP BY RAT_PH", CON_AM)
                Dim PMR_AM_DT As New DataTable
                PMR_AM_DA.Fill(PMR_AM_DT)
                Dim PMR_AM_DA1 As New SqlDataAdapter("SELECT CNAME FROM MAINPOPUS GROUP BY CNAME", CON_AM)
                Dim PMR_AM_DT1 As New DataTable
                PMR_AM_DA1.Fill(PMR_AM_DT1)
                Dim PMR_AM_DA2 As New SqlDataAdapter("SELECT * FROM MAINPOPUS", CON_AM)
                Dim PMR_AM_DT2 As New DataTable
                PMR_AM_DA2.Fill(PMR_AM_DT2)
                workbook.Worksheets.Add("POPULATION")
                xlWorkSheet = workbook.Worksheet("POPULATION")
                For I As Integer = 0 To PMR_AM_DT.Rows.Count - 1
                    xlWorkSheet.Cell(5, I + 3).Value = PMR_AM_DT(I)("RAT_PH")
                    For Y As Integer = 0 To PMR_AM_DT1.Rows.Count - 1
                        Dim dv As New DataView(PMR_AM_DT2)
                        dv.RowFilter = "CNAME='" & PMR_AM_DT1(Y)("CNAME") & "'" & " AND RAT_PH='" & PMR_AM_DT(I)("RAT_PH") & "'"
                        xlWorkSheet.Cell(Y + 6, 2).Value = PMR_AM_DT1(Y)("CNAME")
                        xlWorkSheet.Cell(Y + 6, I + 3).Value = dv.Count
                    Next
                Next
                Dim Z As Integer = xlWorkSheet.ColumnsUsed.Count + 1
                Dim V As Integer = xlWorkSheet.RowsUsed.Count + 4
                For i As Integer = 6 To V
                    Dim fcol = ChrW(2 + 65)
                    Dim lcol = ChrW(Z - 1 + 65)
                    xlWorkSheet.Cell(i, Z + 1).FormulaA1 = "=sum(" & fcol & i & ":" & lcol & i & ")"
                Next
                For i As Integer = 3 To Z
                    Dim col = ChrW(i - 1 + 65)
                    xlWorkSheet.Cell(V + 1, i).FormulaA1 = "=sum(" & col & "6" & ":" & col & V & ")"
                Next
                Dim col1 = ChrW(Z + 65)
                xlWorkSheet.Cell(V + 1, Z + 1).FormulaA1 = "=sum(" & col1 & "6" & ":" & col1 & V & ")"
                xlWorkSheet.Cell(5, Z + 1).Value = "TOTAL"
                xlWorkSheet.Cell(V + 1, 2).Value = "TOTAL"

                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, 5, xlWorkSheet.ColumnsUsed.Count + 1).Style.Fill.BackgroundColor = XLColor.Turquoise
                Dim LSTCOL As Integer = xlWorkSheet.ColumnsUsed.Count + 1
                Dim PMR_AM_DA3 As New SqlDataAdapter("SELECT AMC FROM MAINPOPUS GROUP BY AMC", CON_AM)
                Dim PMR_AM_DT3 As New DataTable
                PMR_AM_DA3.Fill(PMR_AM_DT3)
                For I As Integer = 0 To PMR_AM_DT3.Rows.Count - 1
                    xlWorkSheet.Cell(5, I + (LSTCOL + 3)).Value = PMR_AM_DT3(I)("AMC")
                    For Y As Integer = 0 To PMR_AM_DT1.Rows.Count - 1
                        Dim dv As New DataView(PMR_AM_DT2)
                        dv.RowFilter = "CNAME='" & PMR_AM_DT1(Y)("CNAME") & "'" & " AND AMC='" & PMR_AM_DT3(I)("AMC") & "'"
                        xlWorkSheet.Cell(Y + 6, LSTCOL + 2).Value = PMR_AM_DT1(Y)("CNAME")
                        xlWorkSheet.Cell(Y + 6, I + (LSTCOL + 3)).Value = dv.Count
                    Next
                Next
                Dim Z1 As Integer = xlWorkSheet.ColumnsUsed.Count + 1
                Dim V1 As Integer = xlWorkSheet.RowsUsed.Count + 4
                For i As Integer = 6 To V1
                    Dim fcol = ChrW((LSTCOL + 2) + 65)
                    Dim lcol = ChrW(Z1 + 65)
                    xlWorkSheet.Cell(i, Z1 + 2).FormulaA1 = "=sum(" & fcol & i & ":" & lcol & i & ")"
                Next
                For i As Integer = (LSTCOL + 3) To Z1 + 1
                    Dim col = ChrW(i - 1 + 65)
                    xlWorkSheet.Cell(V1, i).FormulaA1 = "=sum(" & col & "6" & ":" & col & V1 - 1 & ")"
                Next
                xlWorkSheet.Cell(V1, LSTCOL + 2).Value = "TOTAL"
                xlWorkSheet.Cell(5, Z1 + 2).Value = "TOTAL"
                xlWorkSheet.Columns().AdjustToContents()
                xlWorkSheet.Range(5, LSTCOL + 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 2).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, LSTCOL + 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 2).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, LSTCOL + 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 2).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, LSTCOL + 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 2).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, LSTCOL + 2, 5, xlWorkSheet.ColumnsUsed.Count + 2).Style.Fill.BackgroundColor = XLColor.Turquoise

                xlWorkSheet.Cell(2, 2).Value = "POPULATION DETAILS FOR THE MONTH OF " & Now.ToString("MMM") & " - " & Now.ToString("yyyy")
                xlWorkSheet.Range("b2", ChrW(85) & "2").Merge()
                xlWorkSheet.Row(2).Height = 26.25
                xlWorkSheet.Cell(4, 2).Value = "RATING WISE DETAILS"
                xlWorkSheet.Cell(4, LSTCOL + 2).Value = "AMC WISE DETAILS"
                xlWorkSheet.Range("B2").Style.Font.FontName = "Book Antiqua"
                xlWorkSheet.Range("B2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                xlWorkSheet.Range("B2").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
                xlWorkSheet.Range("B2").Style.Font.FontSize = 20
                xlWorkSheet.Range("B2").Style.Font.Bold = True
                xlWorkSheet.Range("A5", "Z5").Style.Font.Bold = True

                workbook.Worksheets.Add("CPT")
                xlWorkSheet = workbook.Worksheet("CPT")
                Dim PMR_AM_DA4 As New SqlDataAdapter("SELECT AMC FROM PMRS GROUP BY AMC", CON_AM)
                Dim PMR_AM_DT4 As New DataTable
                PMR_AM_DA4.Fill(PMR_AM_DT4)
                Dim PMR_AM_DA5 As New SqlDataAdapter("SELECT CUST FROM PMRS GROUP BY CUST", CON_AM)
                Dim PMR_AM_DT5 As New DataTable
                PMR_AM_DA5.Fill(PMR_AM_DT5)
                Dim PMR_AM_DA6 As New SqlDataAdapter("SELECT * FROM PMRS", CON_AM)
                Dim PMR_AM_DT6 As New DataTable
                PMR_AM_DA6.Fill(PMR_AM_DT6)
                Dim D1 As Date = Now.ToString("MM") & "/01/" & Now.ToString("yyyy")
                Dim D2 As Date = Now.ToString("MM") & "/30/" & Now.ToString("yyyy")
                For I As Integer = 0 To PMR_AM_DT4.Rows.Count - 1
                    xlWorkSheet.Cell(5, I + 4).Value = PMR_AM_DT4(I)("AMC")
                    For Y As Integer = 0 To PMR_AM_DT5.Rows.Count - 1
                        Dim dv As New DataView(PMR_AM_DT6)
                        dv.RowFilter = "CUST='" & PMR_AM_DT5(Y)("CUST") & "'" & " AND AMC='" & PMR_AM_DT4(I)("AMC") & "' AND DOS >= #" & D1 & "# and DOS <= #" & D2 & "# AND STYPE='CUSTOMER COMPLAINT'"
                        Dim dv1 As New DataView(PMR_AM_DT2)
                        dv1.RowFilter = "CNAME='" & PMR_AM_DT5(Y)("CUST") & "'" & " AND AMC='AMC'"
                        xlWorkSheet.Cell(Y + 6, 2).Value = PMR_AM_DT5(Y)("CUST")
                        xlWorkSheet.Cell(Y + 6, 3).Value = dv1.Count
                        xlWorkSheet.Cell(Y + 6, I + 4).Value = dv.Count
                    Next
                Next
                Dim Z2 As Integer = xlWorkSheet.ColumnsUsed.Count + 1
                Dim V2 As Integer = xlWorkSheet.RowsUsed.Count + 4
                For i As Integer = 6 To V2 + 1
                    Dim fcol = ChrW(3 + 65)
                    Dim lcol = ChrW(Z2 - 1 + 65)
                    xlWorkSheet.Cell(i, Z2 + 1).FormulaA1 = ChrW(3 + 65) & i & "/" & ChrW(2 + 65) & i & "*1000"
                Next
                For i As Integer = 3 To Z2
                    Dim col = ChrW(i - 1 + 65)
                    xlWorkSheet.Cell(V2 + 1, i).FormulaA1 = "=sum(" & col & "6" & ":" & col & V2 & ")"
                Next
                xlWorkSheet.Cell(5, Z2 + 1).Value = "TOTAL"
                xlWorkSheet.Cell(V2 + 1, 2).Value = "TOTAL"
                xlWorkSheet.Cell(5, 3).Value = "TOTAL AMC POPULATION"
                xlWorkSheet.Cell(5, 2).Value = "CUSTOMER"
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, 5, xlWorkSheet.ColumnsUsed.Count + 1).Style.Fill.BackgroundColor = XLColor.Turquoise
                xlWorkSheet.Columns().AdjustToContents()
                xlWorkSheet.Row(2).Height = 0
                Dim CPT As String = xlWorkSheet.Cell(V2, Z2).Value


                workbook.Worksheets.Add("SERVICE")
                xlWorkSheet = workbook.Worksheet("SERVICE")
                Dim PM_DA As New SqlDataAdapter("SELECT STYPE FROM PMRS GROUP BY STYPE", CON_AM)
                Dim PM_DT As New DataTable
                PM_DA.Fill(PM_DT)
                Dim PM_DA1 As New SqlDataAdapter("SELECT CUST FROM PMRS GROUP BY CUST", CON_AM)
                Dim PM_DT1 As New DataTable
                PM_DA1.Fill(PM_DT1)
                Dim PM_DA2 As New SqlDataAdapter("SELECT * FROM PMRS", CON_AM)
                Dim PM_DT2 As New DataTable
                PM_DA2.Fill(PM_DT2)
                For I As Integer = 0 To PM_DT.Rows.Count - 1
                    xlWorkSheet.Cell(5, I + 3).Value = PM_DT(I)("STYPE")
                    For Y As Integer = 0 To PM_DT1.Rows.Count - 1
                        Dim dv As New DataView(PM_DT2)
                        dv.RowFilter = "CUST='" & PM_DT1(Y)("CUST") & "'" & " AND STYPE='" & PM_DT(I)("STYPE") & "' AND DOS >= #" & D1 & "# and DOS <= #" & D2 & "#"
                        xlWorkSheet.Cell(Y + 6, 2).Value = PM_DT1(Y)("CUST")
                        xlWorkSheet.Cell(Y + 6, I + 3).Value = dv.Count
                    Next
                Next
                Dim Z3 As Integer = xlWorkSheet.ColumnsUsed.Count + 1
                Dim V3 As Integer = xlWorkSheet.RowsUsed.Count + 4
                For i As Integer = 6 To V3
                    Dim fcol = ChrW(2 + 65)
                    Dim lcol = ChrW(Z3 - 1 + 65)
                    xlWorkSheet.Cell(i, Z3 + 1).FormulaA1 = "=sum(" & fcol & i & ":" & lcol & i & ")"
                Next
                For i As Integer = 3 To Z3
                    Dim col = ChrW(i - 1 + 65)
                    xlWorkSheet.Cell(V3 + 1, i).FormulaA1 = "=sum(" & col & "6" & ":" & col & V3 & ")"
                Next
                Dim col2 = ChrW(Z3 + 65)
                xlWorkSheet.Cell(V3 + 1, Z3 + 1).FormulaA1 = "=sum(" & col2 & "6" & ":" & col2 & V3 & ")"
                xlWorkSheet.Cell(5, Z3 + 1).Value = "TOTAL"
                xlWorkSheet.Cell(V3 + 1, 2).Value = "TOTAL"
                xlWorkSheet.Cell(5, 2).Value = "CUSTOMER"
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, 5, xlWorkSheet.ColumnsUsed.Count + 1).Style.Fill.BackgroundColor = XLColor.Turquoise
                xlWorkSheet.Columns().AdjustToContents()

                workbook.Worksheets.Add("HMR")
                xlWorkSheet = workbook.Worksheet("HMR")
                Dim HM_DA As New SqlDataAdapter("SELECT HMAGE FROM MAINPOPUS GROUP BY HMAGE", CON_AM)
                Dim HM_DT As New DataTable
                HM_DA.Fill(HM_DT)
                Dim HM_DA2 As New SqlDataAdapter("SELECT * FROM MAINPOPUS", CON_AM)
                Dim HM_DT2 As New DataTable
                HM_DA2.Fill(HM_DT2)
                For I As Integer = 0 To HM_DT.Rows.Count - 1
                    xlWorkSheet.Cell(5, I + 3).Value = HM_DT(I)("HMAGE")
                    For Y As Integer = 0 To PMR_AM_DT1.Rows.Count - 1
                        Dim dv As New DataView(HM_DT2)
                        dv.RowFilter = "CNAME='" & PMR_AM_DT1(Y)("CNAME") & "'" & " AND HMAGE='" & HM_DT(I)("HMAGE") & "'"
                        xlWorkSheet.Cell(Y + 6, 2).Value = PMR_AM_DT1(Y)("CNAME")
                        xlWorkSheet.Cell(Y + 6, I + 3).Value = dv.Count
                    Next
                Next
                Dim Z6 As Integer = xlWorkSheet.ColumnsUsed.Count + 1
                Dim V6 As Integer = xlWorkSheet.RowsUsed.Count + 4
                For i As Integer = 6 To V6
                    Dim fcol = ChrW(2 + 65)
                    Dim lcol = ChrW(Z6 - 1 + 65)
                    xlWorkSheet.Cell(i, Z6 + 1).FormulaA1 = "=sum(" & fcol & i & ":" & lcol & i & ")"
                Next
                For i As Integer = 3 To Z6
                    Dim col = ChrW(i - 1 + 65)
                    xlWorkSheet.Cell(V6 + 1, i).FormulaA1 = "=sum(" & col & "6" & ":" & col & V6 & ")"
                Next
                Dim col8 = ChrW(Z6 + 65)
                xlWorkSheet.Cell(V6 + 1, Z6 + 1).FormulaA1 = "=sum(" & col8 & "6" & ":" & col8 & V6 & ")"
                xlWorkSheet.Cell(5, Z6 + 1).Value = "TOTAL"
                xlWorkSheet.Cell(V6 + 1, 2).Value = "TOTAL"
                xlWorkSheet.Cell(5, 2).Value = "CUSTOMER"
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, xlWorkSheet.RowsUsed.Count + 4, xlWorkSheet.ColumnsUsed.Count + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range(5, 2, 5, xlWorkSheet.ColumnsUsed.Count + 1).Style.Fill.BackgroundColor = XLColor.Turquoise
                xlWorkSheet.Columns().AdjustToContents()

                workbook.Worksheets.Add("SLA")
                xlWorkSheet = workbook.Worksheet("SLA")
                xlWorkSheet.Cell(2, 2).Value = "SLA DETAILS"
                xlWorkSheet.Range("B2", "E2").Merge()
                xlWorkSheet.Cell(3, 2).Value = "DESCRIPTION"
                xlWorkSheet.Cell(3, 3).Value = "NO OF COMPLAINT"
                xlWorkSheet.Cell(3, 4).Value = "WITHIN SLA"
                xlWorkSheet.Cell(3, 5).Value = "OUT OF SLA"
                xlWorkSheet.Cell(4, 2).Value = "MINOR"
                xlWorkSheet.Cell(5, 2).Value = "MAJOR"
                Dim PMR_dv1 As New DataView(PM_DT2)
                PMR_dv1.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MINOR'"
                Dim MINOR_COM1 As Integer = PMR_dv1.Count
                xlWorkSheet.Cell(4, 3).Value = MINOR_COM1
                PMR_dv1.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MAJOR'"
                Dim MAJOR_COM1 As Integer = PMR_dv1.Count
                xlWorkSheet.Cell(5, 3).Value = MAJOR_COM1
                PMR_dv1.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MINOR' AND SLA='WITHIN SLA'"
                xlWorkSheet.Cell(4, 4).Value = PMR_dv1.Count
                PMR_dv1.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MAJOR' AND SLA='WITHIN SLA'"
                xlWorkSheet.Cell(5, 4).Value = PMR_dv1.Count
                PMR_dv1.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MINOR' AND SLA='OUT OF SLA'"
                xlWorkSheet.Cell(4, 5).Value = PMR_dv1.Count
                PMR_dv1.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MAJOR' AND SLA='OUT OF SLA'"
                xlWorkSheet.Cell(5, 5).Value = PMR_dv1.Count
                xlWorkSheet.Range("B3", "E5").Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("B3", "E5").Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("B3", "E5").Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("B3", "E5").Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("B3", "E3").Style.Fill.BackgroundColor = XLColor.Turquoise
                xlWorkSheet.Range("B2").Style.Fill.BackgroundColor = XLColor.IndianYellow
                xlWorkSheet.Columns().AdjustToContents()

                xlWorkSheet = workbook.Worksheet("DASHBOARD")
                Dim POP_dv As New DataView(PMR_AM_DT2)
                POP_dv.RowFilter = "AMC='AMC'"
                xlWorkSheet.Cell(7, 2).Value = "CPT DETAILS"
                xlWorkSheet.Range("b7", "d7").Merge()
                xlWorkSheet.Cell(8, 2).Value = "AMC POPULATION"
                xlWorkSheet.Cell(8, 3).Value = "TARGET"
                xlWorkSheet.Cell(8, 4).Value = "ACTUAL"
                POP_dv.RowFilter = "AMC='AMC'"
                Dim AMC_CON As Integer = POP_dv.Count
                xlWorkSheet.Cell(9, 2).Value = AMC_CON
                xlWorkSheet.Cell(9, 3).Value = "1"
                Dim PMR_dv As New DataView(PM_DT2)
                PMR_dv.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "#"
                xlWorkSheet.Cell(9, 4).Value = PMR_dv.Count
                xlWorkSheet.Range("b8", "d9").Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b8", "d9").Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b8", "d9").Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b8", "d9").Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b8", "d8").Style.Fill.BackgroundColor = XLColor.Turquoise
                xlWorkSheet.Range("b7").Style.Fill.BackgroundColor = XLColor.Purple

                xlWorkSheet.Cell(2, 2).Value = "SLA DETAILS"
                xlWorkSheet.Range("b2", "e2").Merge()
                xlWorkSheet.Cell(3, 2).Value = "DESCRIPTION"
                xlWorkSheet.Cell(3, 3).Value = "NO OF COMPLAINT"
                xlWorkSheet.Cell(3, 4).Value = "WITHIN SLA"
                xlWorkSheet.Cell(3, 5).Value = "OUT OF SLA"
                xlWorkSheet.Cell(4, 2).Value = "MINOR"
                xlWorkSheet.Cell(5, 2).Value = "MAJOR"
                PMR_dv.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MINOR'"
                Dim MINOR_COM As Integer = PMR_dv.Count
                xlWorkSheet.Cell(4, 3).Value = MINOR_COM
                PMR_dv.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MAJOR'"
                Dim MAJOR_COM As Integer = PMR_dv.Count
                xlWorkSheet.Cell(5, 3).Value = MAJOR_COM
                PMR_dv.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MINOR' AND SLA='WITHIN SLA'"
                xlWorkSheet.Cell(4, 4).Value = PMR_dv.Count
                PMR_dv.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MAJOR' AND SLA='WITHIN SLA'"
                xlWorkSheet.Cell(5, 4).Value = PMR_dv.Count
                PMR_dv.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MINOR' AND SLA='OUT OF SLA'"
                xlWorkSheet.Cell(4, 5).Value = PMR_dv.Count
                PMR_dv.RowFilter = "STYPE='CUSTOMER COMPLAINT' AND CDATI >= #" & D1 & "# and CDATI <= #" & D2 & "# AND SERV='MAJOR' AND SLA='OUT OF SLA'"
                xlWorkSheet.Cell(5, 5).Value = PMR_dv.Count
                xlWorkSheet.Range("b3", "e5").Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b3", "e5").Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b3", "e5").Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b3", "e5").Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b3", "e3").Style.Fill.BackgroundColor = XLColor.Turquoise
                xlWorkSheet.Range("b2").Style.Fill.BackgroundColor = XLColor.Purple
                xlWorkSheet.Columns().AdjustToContents()

                Dim CST_CON As Integer = 0
                For Y As Integer = 0 To PM_DT1.Rows.Count - 1
                    Dim dv As New DataView(PM_DT2)
                    dv.RowFilter = "CUST='" & PM_DT1(Y)("CUST") & "'" & " AND STYPE='OIL SERVICE' AND DOS >= #" & D1 & "# and DOS <= #" & D2 & "#"
                    xlWorkSheet.Cell(Y + 19, 2).Value = PM_DT1(Y)("CUST")
                    xlWorkSheet.Cell(Y + 19, 3).Value = dv.Count
                    xlWorkSheet.Cell(Y + 19, 4).Value = dv.Count
                    CST_CON = CST_CON + 1
                Next
                xlWorkSheet.Cell(17, 2).Value = "PM DETAILS"
                xlWorkSheet.Range("b17", "d17").Merge()
                xlWorkSheet.Cell(18, 2).Value = "CUSTOMER"
                xlWorkSheet.Cell(18, 3).Value = "PM DONE"
                xlWorkSheet.Cell(18, 4).Value = "PM PLAN"
                Dim PM_L_ROW As Integer = xlWorkSheet.RowsUsed.Count
                xlWorkSheet.Range("b17", "d" & CST_CON + 18).Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b17", "d" & CST_CON + 18).Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b17", "d" & CST_CON + 18).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b17", "d" & CST_CON + 18).Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b18", "d18").Style.Fill.BackgroundColor = XLColor.Turquoise
                xlWorkSheet.Range("b17").Style.Fill.BackgroundColor = XLColor.Purple
                xlWorkSheet.Columns().AdjustToContents()

                xlWorkSheet.Cell(11, 2).Value = "ORDER VS TARGET ACHIVEMENT (LAST 3 MONTHS)"
                xlWorkSheet.Range("b11", "d11").Merge()
                xlWorkSheet.Cell(12, 2).Value = "MONTH"
                xlWorkSheet.Cell(12, 3).Value = "ORDER AMOUNT"
                xlWorkSheet.Cell(12, 4).Value = "TARGET AMOUNT"
                Dim MON As Integer = Now.ToString("MM")
                Dim PU_TOT As Integer = 0
                For I As Integer = 0 To 2
                    Dim DT1 As Date = MON - I & "/1/" & Now.ToString("yyyy")
                    Dim DT2 As Date = MON - I & "/" & Date.DaysInMonth(Today.Year, MON - I) & "/" & Now.ToString("yyyy")
                    Dim PUR_DA As New SqlDataAdapter("SELECT * FROM PURCHSE1 WHERE BDATE BETWEEN @StartDate AND @EndDate", CON_AM)
                    PUR_DA.SelectCommand.Parameters.AddWithValue("@StartDate", DT1)
                    PUR_DA.SelectCommand.Parameters.AddWithValue("@EndDate", DT2)
                    Dim PUR_DT As New DataTable
                    PUR_DA.Fill(PUR_DT)
                    Dim TOT As Integer = 0
                    If PUR_DT.Rows.Count > 0 Then
                        If Not IsDBNull(PUR_DT(I)("BAMT")) Then TOT = PUR_DT.Compute("Sum(BAMT)", "")
                    End If

                    xlWorkSheet.Cell(I + 13, 2).Value = MON - I & "/1/" & Now.ToString("yyyy")
                    xlWorkSheet.Cell(I + 13, 2).Style.NumberFormat.SetFormat("MMM-yyyy")
                    xlWorkSheet.Cell(I + 13, 3).Value = TOT
                    xlWorkSheet.Cell(I + 13, 4).Value = "35000"
                Next
                xlWorkSheet.Range("b11", "d15").Style.Border.TopBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b11", "d15").Style.Border.RightBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b11", "d15").Style.Border.LeftBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b11", "d15").Style.Border.BottomBorder = XLBorderStyleValues.Thin
                xlWorkSheet.Range("b12", "d12").Style.Fill.BackgroundColor = XLColor.Turquoise
                xlWorkSheet.Range("b11").Style.Fill.BackgroundColor = XLColor.Purple
                xlWorkSheet.Columns().AdjustToContents()
                xlWorkSheet.Range("a1", xlWorkSheet.LastColumnUsed.ColumnLetter & xlWorkSheet.LastRowUsed.RowNumber + 1).Style.Border.OutsideBorder = XLBorderStyleValues.SlantDashDot
                xlWorkSheet.Range("a1", xlWorkSheet.LastColumnUsed.ColumnLetter & xlWorkSheet.LastRowUsed.RowNumber + 1).Style.Border.OutsideBorderColor = XLColor.TractorRed
                Dim excelfile1 As String = Server.MapPath("\App_Data\DATA\" & "MONTHLY REPORT FOR " & Format(Today, "MMM-yyyy") & ".xlsx")
                workbook.SaveAs(excelfile1)

                Dim rng = xlWorkSheet.RangeUsed
                Dim mail As New MailMessage
                mail.Subject = "MONTLY REPORT FOR " & Format(Today, "MMM-yyyy")
                mail.To.Add("pathllk3@gmail.com")
                mail.CC.Add("brelcworks@YAHOO.COM")
                mail.From = New MailAddress("brelcworks@gmail.com")
                mail.Body = "Dear Sir," & vbCrLf & vbTab & "Please Find The Monthly Report Sheet in The Attachment." & vbCrLf & vbCrLf & vbCrLf & "Thanks & Regards" & vbCrLf & "Anjan Paul" & vbCrLf & "For, B & R Electrical Works"
                Dim attach As New Attachment(excelfile1)
                mail.Attachments.Add(attach)
                Dim smtp As New SmtpClient("smtp.gmail.com")
                smtp.EnableSsl = True
                smtp.Credentials = New System.Net.NetworkCredential("brelcworks", "ratanbose")
                smtp.Port = "587"
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess Or DeliveryNotificationOptions.OnFailure

                ERR.Text = "MONTHLY REPORT SENT"
                LP1 = True
            Catch ex As Exception
                LP1 = False
                EXLERR(Now.ToString, ex.ToString)
                err_display(ex.ToString)
            End Try
        Loop
    End Sub
    Function IsLastDay(ByVal myDate As Date) As Boolean
        Return myDate.Day = Date.DaysInMonth(myDate.Year, myDate.Month)
    End Function
    Protected Sub err_display(ByVal msg As String)
        ERR.Text = msg
    End Sub


End Class