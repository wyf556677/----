Imports System.Data.Odbc
Imports System.Security.Cryptography
Imports System.Text

Public Class Form1
    Public str As String = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=0"
    Public ESC As Boolean
    Dim myCommand As New OdbcCommand
    Dim mySqlConc As New OdbcConnection
    Dim myReader As OdbcDataReader
    '  Dim dbstr As String = "Driver={MariaDB ODBC 2.0 Driver};Server=120.55.166.174;Database=pvmanager_new;User=root; Password=519618;"
    Dim dbstr As String = "Driver={MariaDB ODBC 2.0 Driver};Server=139.129.229.15;Database=pvmanager_new;User=root; Password=519618;"
    Private Property xlUp As Object

    Private Function GetProv(ByVal city As String, ByRef prov() As String)
        For i As Integer = 0 To UBound(prov)
            If InStr(prov(i), city) > 0 Then
                Dim tmp() As String = Split(prov(i).Substring(0, prov(i).Length - 1), "','")
                Return tmp(8).Replace("province':'", "")
            End If
        Next
        Return ""
    End Function
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox2.Text <> "58707780" Then
            MsgBox("验证码错误！")
            Return
        End If

        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始下载资源列表......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If RadioButton1.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=0"
            End If
            If RadioButton2.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=1"
            End If
            If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=2"
            End If
            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)

            str = "http://pv.tihe-china.com/manager/php/city.php?cmd=ListCity&limit=-1"
            data = wc.DownloadData(str)
            Dim citys As String = System.Text.Encoding.UTF8.GetString(data)
            Dim xx() As String
            xx = Split(citys, "[{")
            citys = xx(1).Replace("}]}", "")
            xx = Split(citys, "},{")


            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim prov0, prov1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim panel As Integer = 0
            Dim paneltotal As Integer = 0
            Dim panel2 As Integer = 0
            Dim paneltotal2 As Integer = 0
            Dim theater As Integer = 0
            Dim theatertotal As Integer = 0
            Dim seat As Integer = 0
            Dim renci As Integer = 0
            Dim seattotal As Integer = 0
            Dim rencitotal As Integer = 0
            Dim piaofang As Integer = 0
            Dim piaofangtotal As Integer = 0
            Dim sumrow As Integer = 0
            Dim provrow As Integer = 5
            Dim row As Integer = 4
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()


            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "销售资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 60
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:L").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:L1").Merge()

                .Range("B2:L2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:L2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:L2").Font.Bold = True
                .Range("B2:L2").Font.Size = 18
                .Range("B2:L2").Merge()

                .Range("B4:L4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:L4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:L4").Font.Bold = True
                .Range("A4:L2000").Font.Size = 10
                .Range("B4:L4").Font.Color = Color.White
                .Range("B4:L4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:L4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:L4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "影厅数"
                .Cells(4, 10) = "座位数"
                .Cells(4, 11) = "人次"
                .Cells(4, 12) = "票房"

                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CheckBox1.Checked Then
                        If run_time(zzz(28), NumericUpDown1.Value) = False Then
                            Continue For
                        End If
                    End If
                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = theater
                            .Cells(sumrow, 10) = seat
                            .Cells(sumrow, 11) = renci
                            .Cells(sumrow, 12) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(8)
                    .Cells(row, 10) = zzz(9)
                    .Cells(row, 11) = zzz(10)
                    .Cells(row, 12) = zzz(11)
                    If zzz(20) = 1 Then
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = theater
                    .Cells(sumrow, 10) = seat
                    .Cells(sumrow, 11) = renci
                    .Cells(sumrow, 12) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:L3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:L3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = theatertotal
                .Cells(3, 10) = seattotal
                .Cells(3, 11) = rencitotal
                .Cells(3, 12) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:L" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("I5:I" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载四连屏资源列表......")
            '四连屏---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            paneltotal = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("销售资源列表"))

            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "四连屏资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 60
                .Columns("H:H").ColumnWidth = 10
                .Columns("I:I").ColumnWidth = 10
                .Columns("J:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "连屏数"
                .Cells(4, 10) = "影厅数"
                .Cells(4, 11) = "座位数"
                .Cells(4, 12) = "人次"
                .Cells(4, 13) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If zzz(23) = "0" Then
                        Continue For
                    End If
                    If CheckBox1.Checked Then
                        If run_time(zzz(28), NumericUpDown1.Value) = False Then
                            Continue For
                        End If
                    End If
                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel2
                            .Cells(sumrow, 10) = theater
                            .Cells(sumrow, 11) = seat
                            .Cells(sumrow, 12) = renci
                            .Cells(sumrow, 13) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        paneltotal2 = paneltotal2 + panel2
                        panel2 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(23)) Then
                        panel = panel + CInt(zzz(23))
                    End If

                    'If IsNumeric(zzz(22)) Then
                    '    panel2 = panel2 + CInt(zzz(22)) / 4
                    'End If
                    If IsNumeric(zzz(23)) Then
                        panel2 = panel2 + CInt(zzz(23)) \ 4
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(23)
                    ' .Cells(row, 9) = (CInt(zzz(22)) / 4).ToString + "组"
                    .Cells(row, 9) = (CInt(zzz(23)) \ 4).ToString + "组"

                    .Cells(row, 10) = zzz(8)
                    .Cells(row, 11) = zzz(9)
                    .Cells(row, 12) = zzz(10)
                    .Cells(row, 13) = zzz(11)
                    If zzz(20) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel2
                    .Cells(sumrow, 10) = theater
                    .Cells(sumrow, 11) = seat
                    .Cells(sumrow, 12) = renci
                    .Cells(sumrow, 13) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal2 = paneltotal2 + panel2
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal2
                .Cells(3, 10) = theatertotal
                .Cells(3, 11) = seattotal
                .Cells(3, 12) = rencitotal
                .Cells(3, 13) = piaofangtotal

                .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（四连屏）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"

                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            ListBox1.Items.Add("开始下载两连屏资源列表......")

            '两连屏---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            paneltotal = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("四连屏资源列表"))

            OrgSheet = excel.Worksheets(3)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "两连屏资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 60
                .Columns("H:H").ColumnWidth = 10
                .Columns("I:I").ColumnWidth = 10
                .Columns("J:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "连屏数"
                .Cells(4, 10) = "影厅数"
                .Cells(4, 11) = "座位数"
                .Cells(4, 12) = "人次"
                .Cells(4, 13) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If zzz(24) = "0" Then
                        Continue For
                    End If
                    If CheckBox1.Checked Then
                        If run_time(zzz(28), NumericUpDown1.Value) = False Then
                            Continue For
                        End If
                    End If
                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel2
                            .Cells(sumrow, 10) = theater
                            .Cells(sumrow, 11) = seat
                            .Cells(sumrow, 12) = renci
                            .Cells(sumrow, 13) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        paneltotal2 = paneltotal2 + panel2
                        panel2 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(24)) Then
                        panel = panel + CInt(zzz(24))
                    End If
                    'If IsNumeric(zzz(23)) Then
                    '    panel2 = panel2 + CInt(zzz(23)) / 2
                    'End If
                    If IsNumeric(zzz(24)) Then
                        panel2 = panel2 + CInt(zzz(24)) \ 2
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(24)
                    ' .Cells(row, 9) = (CInt(zzz(23)) / 2).ToString + "组"
                    .Cells(row, 9) = (CInt(zzz(24)) \ 2).ToString + "组"
                    .Cells(row, 10) = zzz(8)
                    .Cells(row, 11) = zzz(9)
                    .Cells(row, 12) = zzz(10)
                    .Cells(row, 13) = zzz(11)
                    If zzz(20) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel2
                    .Cells(sumrow, 10) = theater
                    .Cells(sumrow, 11) = seat
                    .Cells(sumrow, 12) = renci
                    .Cells(sumrow, 13) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal2 = paneltotal2 + panel2
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal2
                .Cells(3, 10) = theatertotal
                .Cells(3, 11) = seattotal
                .Cells(3, 12) = rencitotal
                .Cells(3, 13) = piaofangtotal

                .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（两连屏）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"

                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载有声资源列表......")

            '有声资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            paneltotal = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("两连屏资源列表"))

            OrgSheet = excel.Worksheets(4)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "有声资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 60
                .Columns("H:H").ColumnWidth = 10
                .Columns("I:I").ColumnWidth = 10
                .Columns("J:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "有声屏数"
                .Cells(4, 10) = "影厅数"
                .Cells(4, 11) = "座位数"
                .Cells(4, 12) = "人次"
                .Cells(4, 13) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If zzz(25) = "0" Then
                        Continue For
                    End If
                    If CheckBox1.Checked Then
                        If run_time(zzz(28), NumericUpDown1.Value) = False Then
                            Continue For
                        End If
                    End If
                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel2
                            .Cells(sumrow, 10) = theater
                            .Cells(sumrow, 11) = seat
                            .Cells(sumrow, 12) = renci
                            .Cells(sumrow, 13) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        paneltotal2 = paneltotal2 + panel2
                        panel2 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                    End If
                    If IsNumeric(zzz(25)) Then
                        panel2 = panel2 + CInt(zzz(25))
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(25)
                    .Cells(row, 10) = zzz(8)
                    .Cells(row, 11) = zzz(9)
                    .Cells(row, 12) = zzz(10)
                    .Cells(row, 13) = zzz(11)
                    If zzz(20) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel2
                    .Cells(sumrow, 10) = theater
                    .Cells(sumrow, 11) = seat
                    .Cells(sumrow, 12) = renci
                    .Cells(sumrow, 13) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal2 = paneltotal2 + panel2
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal2
                .Cells(3, 10) = theatertotal
                .Cells(3, 11) = seattotal
                .Cells(3, 12) = rencitotal
                .Cells(3, 13) = piaofangtotal

                .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（有声屏）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"

                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载立屏资源列表......")

            '有立屏资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            paneltotal = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("有声资源列表"))

            OrgSheet = excel.Worksheets(5)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "立屏资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 60
                .Columns("H:H").ColumnWidth = 10
                .Columns("I:I").ColumnWidth = 10
                .Columns("J:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "立屏数"
                .Cells(4, 10) = "影厅数"
                .Cells(4, 11) = "座位数"
                .Cells(4, 12) = "人次"
                .Cells(4, 13) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If zzz(26) = "0" Then
                        Continue For
                    End If
                    If CheckBox1.Checked Then
                        If run_time(zzz(28), NumericUpDown1.Value) = False Then
                            Continue For
                        End If
                    End If
                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel2
                            .Cells(sumrow, 10) = theater
                            .Cells(sumrow, 11) = seat
                            .Cells(sumrow, 12) = renci
                            .Cells(sumrow, 13) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        paneltotal2 = paneltotal2 + panel2
                        panel2 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                    End If
                    If IsNumeric(zzz(26)) Then
                        panel2 = panel2 + CInt(zzz(26))
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(26)
                    .Cells(row, 10) = zzz(8)
                    .Cells(row, 11) = zzz(9)
                    .Cells(row, 12) = zzz(10)
                    .Cells(row, 13) = zzz(11)
                    If zzz(20) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel2
                    .Cells(sumrow, 10) = theater
                    .Cells(sumrow, 11) = seat
                    .Cells(sumrow, 12) = renci
                    .Cells(sumrow, 13) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal2 = paneltotal2 + panel2
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal2
                .Cells(3, 10) = theatertotal
                .Cells(3, 11) = seattotal
                .Cells(3, 12) = rencitotal
                .Cells(3, 13) = piaofangtotal

                .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（立屏）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"

                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载可装饰屏资源列表......")

            '可装饰屏资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            paneltotal = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("立屏资源列表"))

            OrgSheet = excel.Worksheets(6)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "可装饰资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 60
                .Columns("H:H").ColumnWidth = 10
                .Columns("I:I").ColumnWidth = 10
                .Columns("J:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "可装饰屏数"
                .Cells(4, 10) = "影厅数"
                .Cells(4, 11) = "座位数"
                .Cells(4, 12) = "人次"
                .Cells(4, 13) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If zzz(27) = "0" Then
                        Continue For
                    End If
                    If CheckBox1.Checked Then
                        If run_time(zzz(28), NumericUpDown1.Value) = False Then
                            Continue For
                        End If
                    End If
                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel2
                            .Cells(sumrow, 10) = theater
                            .Cells(sumrow, 11) = seat
                            .Cells(sumrow, 12) = renci
                            .Cells(sumrow, 13) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        paneltotal2 = paneltotal2 + panel2
                        panel2 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                    End If
                    If IsNumeric(zzz(27)) Then
                        panel2 = panel2 + CInt(zzz(27))
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(26)
                    .Cells(row, 10) = zzz(8)
                    .Cells(row, 11) = zzz(9)
                    .Cells(row, 12) = zzz(10)
                    .Cells(row, 13) = zzz(11)
                    If zzz(20) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel2
                    .Cells(sumrow, 10) = theater
                    .Cells(sumrow, 11) = seat
                    .Cells(sumrow, 12) = renci
                    .Cells(sumrow, 13) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal2 = paneltotal2 + panel2
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal2
                .Cells(3, 10) = theatertotal
                .Cells(3, 11) = seattotal
                .Cells(3, 12) = rencitotal
                .Cells(3, 13) = piaofangtotal

                .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（可装饰屏）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"

                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            If RadioButton4.Checked = True Then
                ListBox1.Items.Add("开始下载非正常营业资源列表......")
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=2&tradingId=2"

                data = wc.DownloadData(str)
                content = System.Text.Encoding.UTF8.GetString(data)

                city0 = ""
                city1 = ""
                prov0 = ""
                prov1 = ""
                region = 0
                regiontotal = 0
                citycount = 0
                citytotal = 0
                panel = 0
                paneltotal = 0
                theater = 0
                theatertotal = 0
                seat = 0
                seattotal = 0
                renci = 0
                rencitotal = 0
                piaofang = 0
                piaofangtotal = 0
                sumrow = 0
                provrow = 5
                row = 4
                sumcity = 0

                excel.Worksheets.Add(After:=excel.Worksheets("可装饰资源列表"))

                OrgSheet = excel.Worksheets(7)
                excel.ActiveWindow.DisplayGridlines = False
                With OrgSheet
                    .Name = "非正常营业影院"
                    .Columns("A:A").ColumnWidth = 5
                    .Columns("B:D").ColumnWidth = 8.38
                    .Columns("E:E").ColumnWidth = 28.88
                    .Columns("F:F").ColumnWidth = 47.5
                    .Columns("G:G").ColumnWidth = 60
                    .Columns("H:H").ColumnWidth = 7.5
                    .Columns("I:L").ColumnWidth = 7

                    .Rows("1:1").RowHeight = 14.25
                    .Rows("2:2").RowHeight = 86.25
                    .Rows("3:3").RowHeight = 16.5
                    .Rows("4:4").RowHeight = 14.25
                    .Rows("5:2000").RowHeight = 16.5

                    .Range("A1:L1").Merge()

                    .Range("B2:L2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B2:L2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B2:L2").Font.Bold = True
                    .Range("B2:L2").Font.Size = 18
                    .Range("B2:L2").Merge()

                    .Range("B4:L4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B4:L4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B4:L4").Font.Bold = True
                    .Range("A4:L2000").Font.Size = 10
                    .Range("B4:L4").Font.Color = Color.White
                    .Range("B4:L4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                    .Range("B4:L4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                    .Range("B4:L4").Interior.Color = 812276



                    '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                    .Cells(4, 2) = "省份"
                    .Cells(4, 3) = "城市"
                    .Cells(4, 4) = "序号"
                    .Cells(4, 5) = "影院名称"
                    .Cells(4, 6) = "地址"
                    .Cells(4, 7) = "附近写字楼、附近商场、百货"
                    .Cells(4, 8) = "屏数"
                    .Cells(4, 9) = "影厅数"
                    .Cells(4, 10) = "座位数"
                    .Cells(4, 11) = "人次"
                    .Cells(4, 12) = "票房"

                    yy = Split(content, "[{")
                    content = yy(1).Replace("}]}", "")

                    yy = Split(content, "},{")
                    ProgressBar1.Maximum = UBound(yy)
                    ProgressBar1.Minimum = 0
                    ProgressBar1.Visible = True
                    For i As Integer = 0 To UBound(yy)
                        ProgressBar1.Value = i
                        zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                        For j As Integer = 0 To UBound(zd)
                            value = Split(zd(j), "':'")
                            zzz(j) = value(1)
                        Next
                        If CheckBox1.Checked Then
                            If run_time(zzz(28), NumericUpDown1.Value) = False Then
                                Continue For
                            End If
                        End If
                        city1 = zzz(19)
                        prov1 = GetProv(city1, xx)

                        If prov0 <> prov1 Then
                            If prov0 <> "" Then
                                .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                                .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                                .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                                .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                                .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                                .Cells(provrow.ToString, 2) = prov0
                                provrow = row + 1
                            End If
                        End If

                        If city0 <> city1 Then
                            If city0 <> "" Then
                                .Cells(sumrow, 3) = "1"
                                .Cells(sumrow, 4) = "小计："
                                .Cells(sumrow, 5) = region
                                .Cells(sumrow, 8) = panel
                                .Cells(sumrow, 9) = theater
                                .Cells(sumrow, 10) = seat
                                .Cells(sumrow, 11) = renci
                                .Cells(sumrow, 12) = piaofang
                                .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                                .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                                .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                                .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                                .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                                .Cells((sumrow + 1).ToString, 3) = city0
                            End If
                            sumcity = sumcity + 1
                            row = row + 1
                            sumrow = row
                            regiontotal = regiontotal + region
                            region = 0
                            paneltotal = paneltotal + panel
                            panel = 0
                            theatertotal = theatertotal + theater
                            theater = 0
                            seattotal = seattotal + seat
                            seat = 0
                            rencitotal = rencitotal + renci
                            renci = 0
                            piaofangtotal = piaofangtotal + piaofang
                            piaofang = 0
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.ThemeColor = 7
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.TintAndShade = 0.599993896298105
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternTintAndShade = 0
                        End If


                        row = row + 1
                        region = region + 1
                        If IsNumeric(zzz(2)) Then
                            panel = panel + CInt(zzz(2))
                        End If
                        If IsNumeric(zzz(8)) Then
                            theater = theater + CInt(zzz(8))
                        End If
                        If IsNumeric(zzz(9)) Then
                            seat = seat + CInt(zzz(9))
                        End If
                        If IsNumeric(zzz(10)) Then
                            renci = renci + CInt(zzz(10))
                        End If
                        If IsNumeric(zzz(11)) Then
                            piaofang = piaofang + CInt(zzz(11))
                        End If
                        '   .Cells(row, 2) = prov1
                        '   .Cells(row, 3) = zzz(18)
                        .Cells(row, 4) = region
                        .Cells(row, 5) = zzz(3)
                        .Cells(row, 6) = zzz(5)
                        .Cells(row, 7) = zzz(21)
                        .Cells(row, 8) = zzz(2)
                        .Cells(row, 9) = zzz(8)
                        .Cells(row, 10) = zzz(9)
                        .Cells(row, 11) = zzz(10)
                        .Cells(row, 12) = zzz(11)
                        If zzz(19) = 1 Then
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.ThemeColor = 8
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.TintAndShade = 0.599993896298105
                            .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternTintAndShade = 0
                        End If
                        city0 = city1
                        prov0 = prov1
                    Next



                    If prov0 <> "" Then
                        .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                        .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                        .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                        .Cells(provrow.ToString, 2) = prov0
                    End If



                    If city0 <> "" Then
                        .Cells(sumrow, 3) = "1"
                        .Cells(sumrow, 4) = "小计："
                        .Cells(sumrow, 5) = region
                        .Cells(sumrow, 8) = panel
                        .Cells(sumrow, 9) = theater
                        .Cells(sumrow, 10) = seat
                        .Cells(sumrow, 11) = renci
                        .Cells(sumrow, 12) = piaofang
                        .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                        .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                        .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                        .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                        .Cells((sumrow + 1).ToString, 3) = city0
                    End If

                    regiontotal = regiontotal + region
                    paneltotal = paneltotal + panel
                    theatertotal = theatertotal + theater
                    seattotal = seattotal + seat
                    rencitotal = rencitotal + renci
                    piaofangtotal = piaofangtotal + piaofang
                    .Range("A3:L3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("A3:L3").Font.Bold = True
                    .Cells(3, 3) = sumcity
                    .Cells(3, 5) = regiontotal
                    .Cells(3, 8) = paneltotal
                    .Cells(3, 9) = theatertotal
                    .Cells(3, 10) = seattotal
                    .Cells(3, 11) = rencitotal
                    .Cells(3, 12) = piaofangtotal

                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（非正常营业影院）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"

                    .Range("A4:L" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                    .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                    .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                    .Range("I5:I" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                    .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                    With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                        .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    End With
                    With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                        .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    End With
                    With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                        .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    End With
                    With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                        .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    End With
                    With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                        .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    End With
                    With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                        .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                    End With

                End With
                excel.Range("A5").Select()
                excel.ActiveWindow.FreezePanes = True
                ProgressBar1.Visible = False
                ProgressBar1.Value = 0

            End If


            ListBox1.Items.Add("正在保存资源列表.......")
            excel.Worksheets(1).select()
            excel.DisplayAlerts = False
            If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "全国影院资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "全国影院资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\全国影院资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\全国影院资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            End If

            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("资源列表下载完毕！")
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If TextBox2.Text <> "58707780" Then
            MsgBox("验证码错误！")
            Return
        End If

        On Error Resume Next
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ESC = False
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            '          str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYB&limit=-1&working_mode=0&tradingId=0"
            If RadioButton1.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYB&limit=-1&working_mode=0&tradingId=0"
            End If
            If RadioButton2.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYB&limit=-1&working_mode=0&tradingId=1"
            End If
            If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYB&limit=-1&working_mode=0&tradingId=2"
            End If
            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)
            Dim yy() As String
            Dim zd() As String
            Dim value() As String
            Dim zzz(28) As String
            Dim region(3000) As String
            Dim panel(3000) As String
            Dim tag(3000) As Boolean
            Dim city As String
            Dim index As Integer


            ListBox1.Items.Clear()

            yy = Split(content, "[{")
            content = yy(1).Replace("}]}", "")

            yy = Split(content, "},{")
            For i As Integer = 0 To UBound(yy)
                zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                For j As Integer = 0 To UBound(zd)
                    value = Split(zd(j), "':'")
                    zzz(j) = value(1)
                Next
                region(i) = zzz(3)
                panel(i) = zzz(2)
            Next
            For i As Integer = 0 To UBound(yy)
                tag(i) = False
            Next
            ProgressBar1.Maximum = UBound(yy)
            ProgressBar1.Minimum = 0
            ProgressBar1.Visible = True
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer = 6
            Dim flag As Boolean
            Dim yyname As String
            excel.Workbooks.Open(OpenFileDialog1.FileName)
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                While (True)
                    Application.DoEvents()
                    If ESC Then
                        Exit While
                    End If
                    If myrow - 6 <= ProgressBar1.Maximum Then
                        ProgressBar1.Value = myrow - 6
                    End If
                    yyname = .Cells(myrow, 5).value
                    city = .Cells(myrow, 3).value
                    If Not IsNumeric(yyname) Then
                        If yyname = "" Then
                            Exit While
                        End If

                        flag = False
                        For i As Integer = 0 To UBound(yy)
                            If region(i) = yyname Then
                                tag(i) = True
                                If panel(i) = .Cells(myrow, 8).value.ToString Then
                                    flag = True
                                    Exit For
                                End If
                                ListBox1.Items.Add("第" + myrow.ToString + "行  " + yyname + "  " + .Cells(myrow, 8).value.ToString + "--->" + panel(i))
                                .Range("D" + myrow.ToString + ":L" + myrow.ToString).Interior.ThemeColor = 10
                                flag = True
                                Exit For
                            End If
                        Next
                        If flag = False Then
                            ListBox1.Items.Add("第" + myrow.ToString + "行  " + yyname + "  " + .Cells(myrow, 8).value.ToString + "--->没有找到！")
                            .Range("D" + myrow.ToString + ":L" + myrow.ToString).Interior.ThemeColor = 9
                        End If
                    End If
                    myrow = myrow + 1
                End While
                index = 1
                For i As Integer = 0 To UBound(yy)
                    If tag(i) = False Then
                        .Cells(myrow, 3) = index.ToString
                        .Cells(myrow, 4) = "新增影院："
                        .Cells(myrow, 5) = region(i)
                        .Cells(myrow, 6) = panel(i) + "块屏"
                        .Range("B" + myrow.ToString + ":L" + myrow.ToString).Interior.ThemeColor = 6
                        myrow = myrow + 1
                        index += 1
                    End If
                Next
            End With

            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            Dim s() As String
            Dim path As String
            Dim filename As String
            s = Split(OpenFileDialog1.FileName, "\")
            filename = s(UBound(s))
            ListBox1.Items.Add(" ")
            path = OpenFileDialog1.FileName.Replace(".xlsx", "比较结果.xlsx")
            excel.DisplayAlerts = False
            If ESC = False Then
                ListBox1.Items.Add("比较结果保存在： " + path)
                excel.Workbooks(1).SaveAs(path)
            End If
            If ESC Then
                excel.Workbooks(1).Close(SaveChanges:=False)
            Else
                excel.Workbooks(1).Close()
            End If

            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            If ESC = False Then
                ListBox1.Items.Add("资源列表比较完毕！")
            End If
        End If
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            ESC = True
        End If
    End Sub

    Public Sub New()

        MyBase.New()
        MyBase.KeyPreview = True
        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox2.Text <> "58707780" Then
            MsgBox("验证码错误！")
            Return
        End If

        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始下载资源列表......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            str = "http://pv.tihe-china.com/manager/php/terminal.php?cmd=ListTerminalForPvshowZYA"

            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)

            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim prov0, prov1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim panel As Integer = 0
            Dim paneltotal As Integer = 0
            Dim theater As Integer = 0
            Dim theatertotal As Integer = 0
            Dim seat As Integer = 0
            Dim seattotal As Integer = 0
            Dim piaofang As Integer = 0
            Dim piaofangtotal As Integer = 0
            Dim sumrow As Integer = 0
            Dim provrow As Integer = 5
            Dim row As Integer = 4
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()


            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "屏幕终端列表"
                .Columns("A:A").ColumnWidth = 8
                .Columns("B:B").ColumnWidth = 30
                .Columns("B:B").WrapText = True
                .Columns("C:C").ColumnWidth = 10
                .Columns("D:D").ColumnWidth = 20
                .Columns("E:E").ColumnWidth = 40
                .Columns("E:E").WrapText = True
                .Columns("F:F").ColumnWidth = 20
                .Columns("G:G").ColumnWidth = 30
                .Columns("H:H").ColumnWidth = 30
                .Columns("I:I").ColumnWidth = 10
                .Columns("J:J").ColumnWidth = 20
                .Columns("K:L").ColumnWidth = 20
                .Columns("L:L").ColumnWidth = 10
                .Columns("M:M").ColumnWidth = 40

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:10000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M10000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276


                .Cells(4, 2) = "影院名称"
                .Cells(4, 3) = "终端ID"
                .Cells(4, 4) = "MAC"
                .Cells(4, 5) = "海报机编号"
                .Cells(4, 6) = "开机时间"
                .Cells(4, 7) = "点位描述"
                .Cells(4, 8) = "连屏信息"
                .Cells(4, 9) = "终端状态"
                .Cells(4, 10) = "终端尺寸"
                .Cells(4, 11) = "资产编号"
                .Cells(4, 12) = "终端型号"
                .Cells(4, 13) = "备注"

                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next

                    prov1 = zzz(0)


                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                            row += 1
                            .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                            .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                            .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                            .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                            .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0

                        End If
                    End If


                    row = row + 1
                    region = region + 1

                    .Cells(row, 3) = zzz(1)
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(12).Replace("\n", "")
                    .Cells(row, 7) = zzz(5)
                    .Cells(row, 8) = zzz(6)
                    .Cells(row, 9) = zzz(4)
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True

                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院终端资源列表" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + region.ToString + "台）"
                End If

                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("I5:I" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0


            ListBox1.Items.Add("正在保存资源列表.......")
            excel.DisplayAlerts = False
            If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "全国影院终端资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "全国影院终端资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\全国影院终端资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\全国影院终端资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            End If

            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("终端资源列表下载完毕！")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text <> "58707780" Then
            MsgBox("验证码错误！")
            Return
        End If

        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始下载资源列表......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            str = "http://pv.tihe-china.com/manager/php/terminal.php?cmd=ListTerminalForPvshowZYB"

            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)

            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim prov0, prov1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim panel As Integer = 0
            Dim paneltotal As Integer = 0
            Dim theater As Integer = 0
            Dim theatertotal As Integer = 0
            Dim seat As Integer = 0
            Dim seattotal As Integer = 0
            Dim piaofang As Integer = 0
            Dim piaofangtotal As Integer = 0
            Dim sumrow As Integer = 0
            Dim provrow As Integer = 5
            Dim row As Integer = 4
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()


            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "屏幕终端列表"
                .Columns("A:A").ColumnWidth = 8
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 30
                .Columns("D:D").ColumnWidth = 10
                .Columns("E:E").ColumnWidth = 15
                .Columns("F:F").ColumnWidth = 15
                .Columns("G:G").ColumnWidth = 15
                .Columns("H:H").ColumnWidth = 15
                .Columns("I:I").ColumnWidth = 20
                .Columns("J:J").ColumnWidth = 10
                .Columns("K:L").ColumnWidth = 10
                .Columns("L:L").ColumnWidth = 10
                .Columns("M:M").ColumnWidth = 10
                .Columns("N:N").ColumnWidth = 10


                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:10000").RowHeight = 16.5

                .Range("A1:N1").Merge()

                .Range("B2:N2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:N2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:N2").Font.Bold = True
                .Range("B2:N2").Font.Size = 18
                .Range("B2:N2").Merge()

                .Range("B4:N4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:N4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:N4").Font.Bold = True
                .Range("A4:N10000").Font.Size = 10
                .Range("B4:N4").Font.Color = Color.White
                .Range("B4:N4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:N4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:N4").Interior.Color = 812276


                .Cells(4, 2) = "terminalId"
                .Cells(4, 3) = "PCMark"
                .Cells(4, 4) = "regionId"
                .Cells(4, 5) = "terminalType"
                .Cells(4, 6) = "terminalsize"
                .Cells(4, 7) = "terminalno"
                .Cells(4, 8) = "pos"
                .Cells(4, 9) = "positiondesc"
                .Cells(4, 10) = "player_type"
                .Cells(4, 11) = "panelgroup"
                .Cells(4, 12) = "panelid"
                .Cells(4, 13) = "servergroup"
                .Cells(4, 14) = "f4"

                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next

                    row = row + 1
                    region += 1
                    .Cells(row, 2) = zzz(0)
                    .Cells(row, 3) = zzz(1)
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(4)
                    .Cells(row, 7) = zzz(5)
                    .Cells(row, 8) = zzz(6)
                    .Cells(row, 9) = zzz(7)
                    .Cells(row, 10) = zzz(8)
                    .Cells(row, 11) = zzz(9)
                    .Cells(row, 12) = zzz(10)
                    .Cells(row, 13) = zzz(11)
                    .Cells(row, 14) = zzz(12)
                    .Cells(row, 15) = zzz(13)
                    .Cells(row, 16) = zzz(14)

                Next





                .Range("A3:N3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:N3").Font.Bold = True

                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院终端资源列表" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + region.ToString + "台）"
                End If

                .Range("A4:N" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("I5:I" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:N" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:N" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:N" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:N" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:N" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:N" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:N" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0


            ListBox1.Items.Add("正在保存资源列表.......")
            excel.DisplayAlerts = False
            If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "全国影院终端资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "全国影院终端资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\全国影院终端资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\全国影院终端资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            End If

            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("终端资源列表下载完毕！")
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If TextBox2.Text <> "58707780" Then
            MsgBox("验证码错误！")
            Return
        End If

        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始下载资源列表......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If RadioButton1.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=0"
            End If
            If RadioButton2.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=1"
            End If
            If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=2"
            End If
            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)

            str = "http://pv.tihe-china.com/manager/php/city.php?cmd=ListCity&limit=-1"
            data = wc.DownloadData(str)
            Dim citys As String = System.Text.Encoding.UTF8.GetString(data)
            Dim xx() As String
            xx = Split(citys, "[{")
            citys = xx(1).Replace("}]}", "")
            xx = Split(citys, "},{")


            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim prov0, prov1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim panel As Integer = 0
            Dim panel1 As Integer = 0
            Dim paneltotal As Integer = 0
            Dim paneltotal1 As Integer = 0
            Dim panel2 As Integer = 0
            Dim paneltotal2 As Integer = 0
            Dim theater As Integer = 0
            Dim theatertotal As Integer = 0
            Dim seat As Integer = 0
            Dim seattotal As Integer = 0
            Dim renci As Integer = 0
            Dim rencitotal As Integer = 0
            Dim piaofang As Integer = 0
            Dim piaofangtotal As Integer = 0
            Dim sumrow As Integer = 0
            Dim provrow As Integer = 5
            Dim row As Integer = 4
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()


            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "销售资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 47.63
                .Columns("J:L").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:L1").Merge()

                .Range("B2:L2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:L2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:L2").Font.Bold = True
                .Range("B2:L2").Font.Size = 18
                .Range("B2:L2").Merge()

                .Range("B4:L4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:L4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:L4").Font.Bold = True
                .Range("A4:L2000").Font.Size = 10
                .Range("B4:L4").Font.Color = Color.White
                .Range("B4:L4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:L4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:L4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "点位描述"
                .Cells(4, 10) = "影厅数"
                .Cells(4, 11) = "座位数"
                .Cells(4, 12) = "人次"
                .Cells(4, 13) = "票房"

                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next

                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 10) = theater
                            .Cells(sumrow, 11) = seat
                            .Cells(sumrow, 12) = renci
                            .Cells(sumrow, 13) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(7)
                    .Cells(row, 10) = zzz(8)
                    .Cells(row, 11) = zzz(9)
                    .Cells(row, 12) = zzz(10)
                    .Cells(row, 13) = zzz(11)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 10) = theater
                    .Cells(sumrow, 11) = seat
                    .Cells(sumrow, 12) = renci
                    .Cells(sumrow, 13) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:L3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:L3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 10) = theatertotal
                .Cells(3, 11) = seattotal
                .Cells(3, 12) = rencitotal
                .Cells(3, 13) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:L" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("I5:I" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载2屏以上资源列表......")
            '2屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("销售资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "2屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 2 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 2
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 2
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0


            ListBox1.Items.Add("开始下载3屏以上资源列表......")
            '3屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("2屏以上资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "3屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 3 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 3
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 3
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0


            ListBox1.Items.Add("开始下载4屏以上资源列表......")
            '4屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("3屏以上资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "4屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 4 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 4
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 4
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载5屏以上资源列表......")
            '5屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("4屏以上资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "5屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 5 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 5
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 5
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载6屏以上资源列表......")
            '6屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("5屏以上资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "6屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 6 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 6
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 6
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载7屏以上资源列表......")
            '7屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("6屏以上资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "7屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 7 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 7
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 7
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载8屏以上资源列表......")
            '8屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("7屏以上资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "8屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 8 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 8
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 8
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载9屏以上资源列表......")
            '9屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("8屏以上资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "9屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 9 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 9
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 9
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载10屏以上资源列表......")
            '10屏以上资源---------------------------------------------------

            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            panel = 0
            panel1 = 0
            paneltotal = 0
            paneltotal1 = 0
            panel2 = 0
            paneltotal2 = 0
            theater = 0
            theatertotal = 0
            seat = 0
            seattotal = 0
            renci = 0
            rencitotal = 0
            piaofang = 0
            piaofangtotal = 0
            sumrow = 0
            provrow = 5
            row = 4
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("9屏以上资源列表"))
            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False

            With OrgSheet
                .Name = "10屏以上资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 33.25
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:I").ColumnWidth = 7.5
                .Columns("J:J").ColumnWidth = 47.63
                .Columns("K:M").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:M1").Merge()

                .Range("B2:M2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:M2").Font.Bold = True
                .Range("B2:M2").Font.Size = 18
                .Range("B2:M2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "组数"
                .Cells(4, 10) = "点位描述"
                .Cells(4, 11) = "影厅数"
                .Cells(4, 12) = "座位数"
                .Cells(4, 13) = "人次"
                .Cells(4, 14) = "票房"


                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CInt(zzz(2)) < 10 Then
                        Continue For
                    End If
                    city1 = zzz(18)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = panel1
                            .Cells(sumrow, 11) = theater
                            .Cells(sumrow, 12) = seat
                            .Cells(sumrow, 13) = renci
                            .Cells(sumrow, 14) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        paneltotal1 = paneltotal1 + panel1
                        panel = 0
                        panel1 = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                        panel1 = panel1 + CInt(zzz(2)) \ 10
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        piaofang = piaofang + CInt(zzz(10))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(21)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(2) \ 10
                    .Cells(row, 10) = zzz(7)
                    .Cells(row, 11) = zzz(8)
                    .Cells(row, 12) = zzz(9)
                    .Cells(row, 13) = zzz(10)
                    If zzz(19) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = panel1
                    .Cells(sumrow, 11) = theater
                    .Cells(sumrow, 12) = seat
                    .Cells(sumrow, 13) = renci
                    .Cells(sumrow, 14) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                paneltotal1 = paneltotal1 + panel1
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:M3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:M3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = paneltotal1
                .Cells(3, 11) = theatertotal
                .Cells(3, 12) = seattotal
                .Cells(3, 13) = rencitotal
                .Cells(3, 14) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）（2屏以上资源）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("J5:J" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("正在保存资源列表.......")
            excel.Worksheets(1).select()
            excel.DisplayAlerts = False
            If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "全国影院按屏数列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "全国影院按屏数列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\全国影院按屏数列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\全国影院按屏数列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            End If

            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("资源列表下载完毕！")
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If TextBox2.Text <> "58707780" Then
            MsgBox("验证码错误！")
            Return
        End If

        On Error Resume Next
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ESC = False
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            '          str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYB&limit=-1&working_mode=0&tradingId=0"
            If RadioButton1.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYB&limit=-1&working_mode=0&tradingId=0"
            End If
            If RadioButton2.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYB&limit=-1&working_mode=0&tradingId=1"
            End If
            If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYB&limit=-1&working_mode=0&tradingId=2"
            End If
            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)
            Dim yy() As String
            Dim zd() As String
            Dim value() As String
            Dim zzz(28) As String
            Dim region(3000) As String
            Dim regionName(3000) As String
            Dim cityName(3000) As String
            Dim tag(3000) As Boolean
            Dim city As String
            Dim index As Integer
            Dim mycol() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}


            ListBox1.Items.Clear()

            yy = Split(content, "[{")
            content = yy(1).Replace("}]}", "")

            yy = Split(content, "},{")
            For i As Integer = 0 To UBound(yy)
                zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                For j As Integer = 0 To UBound(zd)
                    value = Split(zd(j), "':'")
                    zzz(j) = value(1)
                Next
                regionName(i) = zzz(3)
                region(i) = zzz(4)
                cityName(i) = zzz(5)
            Next

            ProgressBar1.Maximum = UBound(yy)
            ProgressBar1.Minimum = 0
            ProgressBar1.Visible = True
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer
            Dim flag As Boolean
            Dim yyname As String
            Dim zzname As String
            Dim pfall As Decimal
            Dim pftihe As Decimal
            Dim rcall As Decimal
            Dim rctihe As Decimal
            Dim rccol As Integer
            Dim pfcol As Integer
            Dim mccol As Integer
            Dim zzcol As Integer
            Dim titlerow As Integer = 0

            excel.Workbooks.Open(OpenFileDialog1.FileName)
            For i As Integer = 1 To excel.Worksheets.Count
                ProgressBar1.Value = 0
                For k As Integer = 0 To UBound(yy)
                    tag(k) = False
                Next
                OrgSheet = excel.Worksheets(i)
                With (OrgSheet)
                    For k As Integer = 4 To 10
                        If .Cells(k, 1).value = "排名" Then
                            titlerow = k
                            Exit For
                        End If
                    Next
                    ListBox1.Items.Insert(0, "开始处理 " + .Name)
                    mccol = 0
                    For ii As Integer = 1 To 26
                        If .Cells(titlerow, ii).value = "影院名称" Then
                            mccol = ii
                            Exit For
                        End If
                    Next
                    pfcol = 0
                    For ii As Integer = 1 To 26
                        If .Cells(titlerow, ii).value = "票房(万)" Then
                            pfcol = ii
                            Exit For
                        End If
                    Next
                    If pfcol = 0 Then
                        MsgBox("未找到票房列！")
                        ESC = True
                    End If
                    rccol = 0
                    For ii As Integer = 1 To 26
                        If .Cells(titlerow, ii).value = "人次(万)" Then
                            rccol = ii
                            Exit For
                        End If
                    Next
                    If rccol = 0 Then
                        MsgBox("未找到人次列！")
                        ESC = True
                    End If
                    zzcol = 0
                    For ii As Integer = 1 To 26
                        If .Cells(titlerow, ii).value = "专资编码" Then
                            zzcol = ii
                            Exit For
                        End If
                    Next
                    If zzcol = 0 Then
                        MsgBox("未找到专资编码列！")
                        ESC = True
                    End If
                    myrow = 0
                    For ii As Integer = 6 To 16
                        If .Cells(ii, 1).value = "1" Then
                            myrow = ii
                            Exit For
                        End If
                    Next
                    If myrow = 0 Then
                        MsgBox("未找到数据行！")
                        ESC = True
                    End If
                    .Columns(mycol(mccol) + ":" + mycol(mccol)).Insert()
                    .Cells(titlerow, mccol + 1) = "影院名称（泰和）"
                    rccol += 1
                    pfcol += 1
                    zzcol += 1
                    pfall = 0.0
                    pftihe = 0.0
                    rcall = 0.0
                    rctihe = 0.0
                    While (True)
                        Application.DoEvents()
                        If ESC Then
                            Exit While
                        End If

                        zzname = Trim(.Cells(myrow, zzcol).value)
                        yyname = Trim(.Cells(myrow, mccol).value)

                        flag = False
                        pfall += .Cells(myrow, pfcol).value
                        rcall += .Cells(myrow, rccol).value
                        If yyname <> "" Then
                            For l As Integer = 0 To UBound(yy)
                                If Trim(TextBox1.Text) = "" Then
                                    If InStr(region(l), zzname) > 0 Then
                                        If ProgressBar1.Value < ProgressBar1.Maximum Then
                                            ProgressBar1.Value += 1
                                        End If
                                        .Cells(myrow, mccol + 1) = regionName(l)
                                        pftihe += .Cells(myrow, pfcol).value
                                        rctihe += .Cells(myrow, rccol).value
                                        ListBox1.Items.Insert(0, myrow.ToString + "   " + yyname + "-----已找到！")
                                        tag(l) = True
                                        flag = True
                                        Exit For
                                    End If
                                Else
                                    If InStr(region(l), zzname) > 0 And InStr(Trim(TextBox1.Text), cityName(l)) > 0 Then
                                        If ProgressBar1.Value < ProgressBar1.Maximum Then
                                            ProgressBar1.Value += 1
                                        End If
                                        .Cells(myrow, mccol + 1) = regionName(l)
                                        pftihe += .Cells(myrow, pfcol).value
                                        rctihe += .Cells(myrow, rccol).value
                                        ListBox1.Items.Insert(0, myrow.ToString + "   " + yyname + "-----已找到！")
                                        tag(l) = True
                                        flag = True
                                        Exit For
                                    End If
                                End If

                            Next
                        Else
                            Exit While
                        End If
                        If flag = False Then
                            .Rows(myrow.ToString + ":" + myrow.ToString).Delete()
                        Else
                            myrow += 1
                        End If
                    End While
                    Dim tmp As String
                    tmp = "导出数据如下：   总票房（万）" + Math.Round(pfall, 2).ToString + "  泰和票房（万）" + Math.Round(pftihe, 2).ToString + "   占比：" + Math.Round((pftihe / pfall) * 100, 2).ToString + "%;    总人次（万）" + Math.Round(rcall, 2).ToString + "  泰和人次（万）" + Math.Round(rctihe, 2).ToString + "   占比：" + Math.Round((rctihe / rcall) * 100, 2).ToString + "%"
                    .Cells(4, 1) = tmp
                    index = 1
                    myrow += titlerow
                    If ESC = False Then
                        For j As Integer = 0 To UBound(yy)
                            If tag(j) = False Then
                                .Cells(myrow, mccol) = "艺恩未统计的影院："
                                .Cells(myrow, mccol + 1) = regionName(j)
                                myrow += 1
                                ListBox1.Items.Insert(0, index.ToString + "   艺恩名：" + region(j) + "影院名：" + regionName(j) + "-----未找到!")
                                index += 1
                            End If
                        Next
                    End If
                End With
            Next

            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            Dim s() As String
            Dim path As String
            Dim filename As String
            s = Split(OpenFileDialog1.FileName, "\")
            filename = s(UBound(s))
            ListBox1.Items.Add(" ")
            path = OpenFileDialog1.FileName.Replace(".xls", "处理结果.xls")
            excel.DisplayAlerts = False


            ListBox1.Items.Insert(0, "处理结果保存在： " + path)
            excel.Workbooks(1).SaveAs(path)
            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            If ESC = False Then
                ListBox1.Items.Insert(0, "艺恩数据处理完毕！")
            End If
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Cursor.Current = Cursors.WaitCursor
        Dim myCommand As New OdbcCommand
        Dim mySqlConc As New OdbcConnection
        Dim enName As String
        Dim regionName As String

        Dim str As String = "Driver={MariaDB ODBC 2.0 Driver};Server=pv.tihe-china.com;Database=pvmanager_new;User=root; Password=519618;"

        mySqlConc.ConnectionString = str
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc
        ListBox1.Items.Clear()
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer = 3
            excel.Workbooks.Open(OpenFileDialog1.FileName)

            OrgSheet = excel.Worksheets(1)
            With (OrgSheet)
                While (True)
                    Application.DoEvents()
                    If ESC Then
                        Exit While
                    End If
                    enName = Trim(.Cells(myrow, 5).value)
                    regionName = Trim(.Cells(myrow, 6).value)
                    If enName = "" Or regionName = "" Then
                        Exit While
                    End If
                    ' ListBox1.Items.Add("开始处理 " + regionName)
                    myCommand.CommandText = String.Format("Update region set EnName = '{0}' where RegionName = '{1}'", enName, regionName)
                    If myCommand.ExecuteNonQuery() = 0 Then
                        ListBox1.Items.Add("导入失败 " + regionName)
                    End If
                    myrow += 1
                End While
            End With
            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
        End If

        mySqlConc.Close()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Cursor.Current = Cursors.WaitCursor
        Dim myCommand As New OdbcCommand
        Dim mySqlConc As New OdbcConnection
        Dim myReader As OdbcDataReader
        Dim regionid(5000) As Integer
        Dim imgurl(5000) As String
        Dim total As Integer
        Dim ttt As String


        Dim str As String = "Driver={MariaDB ODBC 2.0 Driver};Server=pv.tihe-china.com;Database=pvmanager_new;User=root; Password=519618;"

        mySqlConc.ConnectionString = str
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc
        myCommand.CommandText = String.Format("select RegionId,MapUrl from region")
        myReader = myCommand.ExecuteReader()
        total = 0
        While myReader.Read()
            If myReader.IsDBNull(0) Then
                regionid(total) = 0
            Else
                regionid(total) = myReader.GetInt64(0)
            End If

            If myReader.IsDBNull(1) Then
                imgurl(total) = ""
            Else
                imgurl(total) = Trim(myReader.GetString(1))
            End If
            total += 1
        End While
        myReader.Close()
        For i As Integer = 0 To total - 1
            ttt = imgurl(i).Replace("106.2.189.36", "139.129.229.15")
            myCommand.CommandText = String.Format("Update region set MapUrl = '{0}' where RegionId = '{1}'", ttt, regionid(i))
            myCommand.ExecuteNonQuery()
        Next

        mySqlConc.Close()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Cursor.Current = Cursors.WaitCursor
        Dim myCommand As New OdbcCommand
        Dim mySqlConc As New OdbcConnection
        Dim EnName As String
        Dim tmp As String
        Dim count As Integer

        Dim str As String = "Driver={MariaDB ODBC 2.0 Driver};Server=pv.tihe-china.com;Database=pvmanager_new;User=root; Password=519618;"

        mySqlConc.ConnectionString = str
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc
        ListBox1.Items.Clear()
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer = 3
            excel.Workbooks.Open(OpenFileDialog1.FileName)

            OrgSheet = excel.Worksheets(1)
            With (OrgSheet)
                While (True)
                    Application.DoEvents()
                    If ESC Then
                        Exit While
                    End If
                    tmp = Trim(.Cells(myrow, 4).value)
                    EnName = Trim(.Cells(myrow, 3).value)

                    If EnName = "end" Then
                        Exit While
                    End If
                    myCommand.CommandText = String.Format("select count(regionId) from region where RegionName = '{0}'", EnName)
                    count = myCommand.ExecuteScalar
                    If count > 0 Then
                        '  ListBox1.Items.Insert(0, "找到" + EnName)
                        If IsNumeric(tmp) Then
                            myCommand.CommandText = String.Format("Update region set fr_stat = {0} where RegionName = '{1}'", CInt(tmp), EnName)
                            myCommand.ExecuteNonQuery()
                        End If
                    Else
                        ListBox1.Items.Insert(0, "***********未找到" + EnName)
                    End If
                    myrow += 1
                End While
            End With
            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
        End If
        ListBox1.Items.Insert(0, "导入完毕！")
        mySqlConc.Close()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub ListBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedValueChanged
        If ListBox1.SelectedItems.Count = 0 Then
            Return
        End If
        Dim str As String = ""
        For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
            str = str + ListBox1.SelectedItems(i) + vbCrLf
        Next
        Clipboard.SetText(str)
    End Sub
    Private Function run_time(ByVal str As String, ByVal time As Integer) As Boolean
        If Trim(str) = "" Then
            Return False
        End If
        Dim tt As String = str.Replace("：", ":")
        Dim xx() As String = Split(tt, "-")
        If UBound(xx) <> 1 Then
            Return False
        End If
        Dim yy() As String = Split(xx(1), ":")
        If UBound(yy) <> 1 Then
            Return False
        End If
        Dim eh, em, sh, sm As Integer
        If Not IsNumeric(yy(0)) Or Not IsNumeric(yy(1)) Then
            Return False
        End If

        eh = CInt(yy(0))
        em = CInt(yy(1))
        If eh < 5 Then
            eh = eh + 24
        End If
        eh = eh * 60 + em
        Dim zz() As String = Split(xx(0), ":")
        If UBound(zz) <> 1 Then
            Return False
        End If
        If Not IsNumeric(zz(0)) Or Not IsNumeric(zz(1)) Then
            Return False
        End If
        sh = CInt(zz(0))
        sm = CInt(zz(1))
        sh = sh * 60 + sm
        If eh - sh >= time * 60 Then
            Return True
        End If
        Return False
    End Function

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Cursor.Current = Cursors.WaitCursor
        Dim myCommand As New OdbcCommand
        Dim mySqlConc As New OdbcConnection
        Dim remark As String
        Dim ID As String
        Dim flag As Integer = 0
        Dim tname As String
        Dim str As String = "Driver={MariaDB ODBC 2.0 Driver};Server=pv.tihe-china.com;Database=pvmanager_new;User=root; Password=519618;"

        mySqlConc.ConnectionString = str
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc
        ListBox1.Items.Clear()
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer = 5
            excel.Workbooks.Open(OpenFileDialog1.FileName)
            flag = 0
            OrgSheet = excel.Worksheets(1)
            With (OrgSheet)
                While (True)
                    Application.DoEvents()
                    If ESC Then
                        Exit While
                    End If
                    If flag >= 2 Then
                        Exit While
                    End If
                    ID = Trim(.Cells(myrow, 3).value)
                    tname = Trim(.Cells(myrow, 5).value)
                    remark = Trim(.Cells(myrow, 13).value)
                    If ID = "" Then
                        flag += 1
                        myrow += 1
                        Continue While
                    End If
                    flag = 0
                    myCommand.CommandText = String.Format("Update terminal set Remark = '{0}' where terminalID = {1}", remark, CInt(ID))
                    If myCommand.ExecuteNonQuery() = 0 Then
                        ListBox1.Items.Insert(0, tname + "导入失败 ")
                    Else
                        ListBox1.Items.Insert(0, tname + "导入成功 ")
                    End If
                    myrow += 1
                End While
            End With
            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
        End If

        mySqlConc.Close()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Cursor.Current = Cursors.WaitCursor
        Dim myCommand As New OdbcCommand
        Dim mySqlConc As New OdbcConnection
        Dim regionname As String
        Dim flag As Integer = 0
        Dim ttime As String
        Dim str As String = "Driver={MariaDB ODBC 2.0 Driver};Server=pv.tihe-china.com;Database=pvmanager_new;User=root; Password=519618;"

        mySqlConc.ConnectionString = str
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc
        ListBox1.Items.Clear()
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer = 5
            excel.Workbooks.Open(OpenFileDialog1.FileName)
            flag = 0
            OrgSheet = excel.Worksheets(1)
            With (OrgSheet)
                While (True)
                    Application.DoEvents()
                    If ESC Then
                        Exit While
                    End If

                    regionname = Trim(.Cells(myrow, 5).value)
                    ttime = Trim(.Cells(myrow, 7).value)
                    '   remark = Trim(.Cells(myrow, 13).value)
                    If regionname.Length < 6 Then
                        myrow += 1
                        Continue While
                    End If
                    If regionname = "" Then
                        Exit While
                    End If
                    myCommand.CommandText = String.Format("Update region set run_time = '{0}' where RegionName like '{1}%'", ttime, regionname)
                    If myCommand.ExecuteNonQuery() = 0 Then
                        ListBox1.Items.Insert(0, regionname + "****************导入失败 ")
                    Else
                        ListBox1.Items.Insert(0, regionname + "导入成功 ")
                    End If
                    myrow += 1
                End While
            End With
            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
        End If

        mySqlConc.Close()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If TextBox2.Text <> "58707780" Then
            MsgBox("验证码错误！")
            Return
        End If
        Cursor.Current = Cursors.WaitCursor
        Dim ConnectString As String = "Password=www518;Persist Security Info=True;User ID=sa;Initial Catalog=iposter;Data Source=192.168.1.9"
        Dim myCommand1 As New SqlClient.SqlCommand
        Dim mySqlConc1 As New SqlClient.SqlConnection
        Dim myReader1 As SqlClient.SqlDataReader

        mySqlConc1.ConnectionString = ConnectString
        Try
            mySqlConc1.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand1.Connection = mySqlConc1

        Dim myCommand As New SqlClient.SqlCommand
        Dim mySqlConc As New SqlClient.SqlConnection

        Dim total As Decimal
        Dim total2 As Decimal
        Dim quyu As String
        Dim ggsc As Decimal
        Dim bcpc As Integer
        Dim skrq As Date
        Dim xkrq As Date
        Dim ps As Integer
        Dim xx() As String
        Dim skl As Decimal
        Dim tmp As Decimal
        Dim ggmc As String

        Dim mydate As Integer
        Dim recs As Integer
        Dim str1 As String

        mydate = DateDiff("d", DateTimePicker1.Value, DateTimePicker2.Value) + 1
        total = NumericUpDown2.Value * mydate * 12 * 60 * 60
        total2 = 0
        ListBox1.Items.Clear()

        mySqlConc.ConnectionString = ConnectString
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc
        If ComboBox2.Text = "全部" Or ComboBox2.Text = "商业广告" Then
            If ComboBox1.Text = "全部" Then
                myCommand1.CommandText = String.Format("select count(ID) from 广告发布视图 where 媒体类别='商广' and 发布时间 <= '{0}' and 截止时间 >= '{1}'", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString)
            Else
                If ComboBox1.Text = "购买" Then
                    myCommand1.CommandText = String.Format("select count(ID) from 广告发布视图 where 媒体类别='商广' and  发布时间 <= '{0}' and 截止时间 >= '{1}' and 发布类别 like '%{2}%'", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString, "购买")
                Else
                    myCommand1.CommandText = String.Format("select count(ID) from 广告发布视图 where 媒体类别='商广' and 发布时间 <= '{0}' and 截止时间 >= '{1}' and 发布类别 = '{2}'", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString, ComboBox1.Text)
                End If
            End If

            recs = myCommand1.ExecuteScalar
            ProgressBar1.Maximum = recs
            ProgressBar1.Minimum = 0
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            If ComboBox1.Text = "全部" Then
                myCommand1.CommandText = String.Format("select 发布区域码,发布时长,发布频次,发布时间,截止时间,发布版本 from 广告发布视图 where 媒体类别='商广' and  发布时间 <= '{0}' and 截止时间 >= '{1}' order by 发布时间", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString)
            Else
                If ComboBox1.Text = "购买" Then
                    myCommand1.CommandText = String.Format("select 发布区域码,发布时长,发布频次,发布时间,截止时间,发布版本 from 广告发布视图 where 媒体类别='商广' and  发布时间 <= '{0}' and 截止时间 >= '{1}' and 发布类别 like '%{2}%' order by 发布时间", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString, "购买")
                Else
                    myCommand1.CommandText = String.Format("select 发布区域码,发布时长,发布频次,发布时间,截止时间,发布版本 from 广告发布视图 where 媒体类别='商广' and  发布时间 <= '{0}' and 截止时间 >= '{1}' and 发布类别 = '{2}' order by 发布时间", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString, ComboBox1.Text)
                End If
            End If
            myReader1 = myCommand1.ExecuteReader()
            While myReader1.Read()
                ps = 0
                If myReader1.IsDBNull(0) Then
                    quyu = ""
                Else
                    quyu = Trim(myReader1.GetString(0))
                End If
                If quyu = "" Then
                    ps = 0
                Else
                    If quyu = "全国" Then
                        ps = NumericUpDown2.Value
                    Else
                        xx = Split(quyu, ",")
                        For i As Integer = 0 To UBound(xx)
                            If IsNumeric(xx(i)) Then
                                myCommand.CommandText = String.Format("select postCount from 影院信息表 where regionId='{0}'", xx(i))
                                ps += CInt(myCommand.ExecuteScalar)
                            End If
                        Next
                    End If
                End If
                If myReader1.IsDBNull(1) Then
                    ggsc = 0.0
                Else
                    ggsc = myReader1.GetDecimal(1)
                End If
                If myReader1.IsDBNull(2) Then
                    bcpc = 0
                Else
                    bcpc = CInt(Trim(myReader1.GetInt32(2)))
                End If
                If myReader1.IsDBNull(3) Then
                    skrq = Now()
                Else
                    skrq = myReader1.GetDateTime(3)
                End If
                If DateDiff("d", skrq, DateTimePicker1.Value) > 0 Then
                    skrq = DateTimePicker1.Value
                End If
                If myReader1.IsDBNull(4) Then
                    xkrq = Now()
                Else
                    xkrq = myReader1.GetDateTime(4)
                End If
                If DateDiff("d", DateTimePicker2.Value, xkrq) > 0 Then
                    xkrq = DateTimePicker2.Value
                End If
                If myReader1.IsDBNull(5) Then
                    ggmc = ""
                Else
                    ggmc = myReader1.GetString(5)
                End If
                tmp = ggsc * bcpc * ps * (DateDiff("d", skrq, xkrq) + 1)
                str1 = skrq.ToShortDateString + "  " + xkrq.ToShortDateString + "   " + tmp.ToString + "       " + ggmc
                ListBox1.Items.Insert(0, str1)
                total2 += tmp
                Application.DoEvents()
                If ProgressBar1.Value < ProgressBar1.Maximum Then
                    ProgressBar1.Value += 1
                End If
            End While
            myReader1.Close()
        End If

        If ComboBox2.Text = "全部" Or ComboBox2.Text = "电影海报" Then
            If ComboBox1.Text = "全部" Then
                myCommand1.CommandText = String.Format("select count(ID) from 广告发布视图 where 媒体类别<>'商广' and 发布时间 <= '{0}' and 截止时间 >= '{1}'", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString)
            Else
                If ComboBox1.Text = "购买" Then
                    myCommand1.CommandText = String.Format("select count(ID) from 广告发布视图 where 媒体类别<>'商广' and  发布时间 <= '{0}' and 截止时间 >= '{1}' and 发布类别 like '%{2}%'", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString, "购买")
                Else
                    myCommand1.CommandText = String.Format("select count(ID) from 广告发布视图 where 媒体类别<>'商广' and 发布时间 <= '{0}' and 截止时间 >= '{1}' and 发布类别 = '{2}'", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString, ComboBox1.Text)
                End If
            End If

            recs = myCommand1.ExecuteScalar
            ProgressBar1.Maximum = recs
            ProgressBar1.Minimum = 0
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            If ComboBox1.Text = "全部" Then
                myCommand1.CommandText = String.Format("select 发布区域码,发布时长,发布频次,发布时间,截止时间,发布版本 from 广告发布视图 where 媒体类别<>'商广' and 发布时间 <= '{0}' and 截止时间 >= '{1}' order by 发布时间", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString)
            Else
                If ComboBox1.Text = "购买" Then
                    myCommand1.CommandText = String.Format("select 发布区域码,发布时长,发布频次,发布时间,截止时间,发布版本 from 广告发布视图 where 媒体类别<>'商广' and 发布时间 <= '{0}' and 截止时间 >= '{1}' and 发布类别 like '%{2}%' order by 发布时间", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString, "购买")
                Else
                    myCommand1.CommandText = String.Format("select 发布区域码,发布时长,发布频次,发布时间,截止时间,发布版本 from 广告发布视图 where 媒体类别<>'商广' and 发布时间 <= '{0}' and 截止时间 >= '{1}' and 发布类别 = '{2}' order by 发布时间", DateTimePicker2.Value.ToShortDateString, DateTimePicker1.Value.ToShortDateString, ComboBox1.Text)
                End If
            End If
            myReader1 = myCommand1.ExecuteReader()
            While myReader1.Read()
                ps = 0
                If myReader1.IsDBNull(0) Then
                    quyu = ""
                Else
                    quyu = Trim(myReader1.GetString(0))
                End If
                If quyu = "" Then
                    ps = 0
                Else
                    If quyu = "全国" Then
                        ps = NumericUpDown2.Value
                    Else
                        xx = Split(quyu, ",")
                        For i As Integer = 0 To UBound(xx)
                            If IsNumeric(xx(i)) Then
                                myCommand.CommandText = String.Format("select postCount from 影院信息表 where regionId='{0}'", xx(i))
                                ps += CInt(myCommand.ExecuteScalar)
                            End If
                        Next
                    End If
                End If
                If myReader1.IsDBNull(1) Then
                    ggsc = 0.0
                Else
                    ggsc = myReader1.GetDecimal(1)
                End If
                If myReader1.IsDBNull(2) Then
                    bcpc = 0
                Else
                    bcpc = CInt(Trim(myReader1.GetInt32(2)))
                End If
                If myReader1.IsDBNull(3) Then
                    skrq = Now()
                Else
                    skrq = myReader1.GetDateTime(3)
                End If
                If DateDiff("d", skrq, DateTimePicker1.Value) > 0 Then
                    skrq = DateTimePicker1.Value
                End If
                If myReader1.IsDBNull(4) Then
                    xkrq = Now()
                Else
                    xkrq = myReader1.GetDateTime(4)
                End If
                If DateDiff("d", DateTimePicker2.Value, xkrq) > 0 Then
                    xkrq = DateTimePicker2.Value
                End If
                If myReader1.IsDBNull(5) Then
                    ggmc = ""
                Else
                    ggmc = myReader1.GetString(5)
                End If
                tmp = ggsc * bcpc * ps * (DateDiff("d", skrq, xkrq) + 1)
                str1 = skrq.ToShortDateString + "  " + xkrq.ToShortDateString + "   " + tmp.ToString + "       " + ggmc
                ListBox1.Items.Insert(0, str1)
                total2 += tmp
                Application.DoEvents()
                If ProgressBar1.Value < ProgressBar1.Maximum Then
                    ProgressBar1.Value += 1
                End If
            End While
            myReader1.Close()
        End If

        skl = total2 / total * 100
        Label6.Text = "上刊率：" + Mid(skl.ToString, 1, 5) + "%"
        mySqlConc.Close()
        mySqlConc1.Close()
        ProgressBar1.Visible = False
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Cursor.Current = Cursors.WaitCursor
        Dim myCommand As New OdbcCommand
        Dim mySqlConc As New OdbcConnection
        Dim enName As String
        Dim regionName As String

        Dim str As String = "Driver={MariaDB ODBC 2.0 Driver};Server=pv.tihe-china.com;Database=pvmanager_new;User=root; Password=519618;"

        mySqlConc.ConnectionString = str
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc
        ListBox1.Items.Clear()
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer = 7
            excel.Workbooks.Open(OpenFileDialog1.FileName)

            OrgSheet = excel.Worksheets(1)
            With (OrgSheet)
                While (True)
                    Application.DoEvents()
                    If ESC Then
                        Exit While
                    End If
                    enName = Trim(.Cells(myrow, 3).value)
                    regionName = Trim(.Cells(myrow, 2).value)
                    If enName = "" Or regionName = "" Then
                        Exit While
                    End If
                    ' ListBox1.Items.Add("开始处理 " + regionName)
                    myCommand.CommandText = String.Format("Update region set EnName = '{0}' where EnName = '{1}'", "(" + enName + ")" + regionName, regionName)
                    If myCommand.ExecuteNonQuery() = 0 Then
                        ListBox1.Items.Add("导入失败 " + regionName)
                    Else
                        ListBox1.Items.Add("导入 " + regionName)
                    End If
                    myrow += 1
                End While
            End With
            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
        End If

        mySqlConc.Close()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        On Error Resume Next
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With
        ESC = False
        Dim zcbh As String
        Dim region(10000) As String
        Dim tag(10000) As Boolean
        Dim total As String
        Dim index As Integer
        Dim mycol() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

        Dim excel As New Microsoft.Office.Interop.Excel.Application()
        Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim myrow As Integer
        Dim flag As Boolean

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        ListBox1.Items.Clear()

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            excel.Workbooks.Open(OpenFileDialog1.FileName)

            OrgSheet = excel.Worksheets(1)
            myrow = 2
            index = 0
            With (OrgSheet)
                While (True)
                    If ESC Then
                        Exit While
                    End If
                    region(index) = .Cells(myrow, 3).value
                    tag(index) = False
                    If Trim(region(index)) = "" Then
                        Exit While
                    End If
                    index += 1
                    myrow += 1
                End While
            End With
            excel.Workbooks(1).Close()
        End If
        total = index
        ProgressBar1.Maximum = total
        ProgressBar1.Minimum = 0
        ProgressBar1.Visible = True
        ProgressBar1.Value = 0
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            excel.Workbooks.Open(OpenFileDialog1.FileName)

            OrgSheet = excel.Worksheets(1)
            myrow = 4
            With (OrgSheet)
                While (True)
                    If ESC Then
                        Exit While
                    End If
                    zcbh = .Cells(myrow, 4).value
                    If zcbh = "0" Then
                        Exit While
                    End If
                    flag = False
                    For i As Integer = 0 To total - 1
                        If zcbh = region(i) Then
                            tag(i) = True
                            flag = True
                            Exit For
                        End If
                    Next
                    If flag = False Then
                        ListBox1.Items.Add(myrow.ToString + "    " + zcbh + "**********************未找到！")
                        .Range("A" + myrow.ToString + ":P" + myrow.ToString).Interior.ThemeColor = 10
                    End If
                    myrow += 1
                    If ProgressBar1.Value < ProgressBar1.Maximum Then
                        ProgressBar1.Value += 1
                    End If
                End While
            End With
        End If

        ListBox1.Items.Add("固定资产未被查找：")
        For i As Integer = 0 To total - 1
            If tag(i) = False Then
                ListBox1.Items.Add((i + 2).ToString + "      " + region(i))
            End If
        Next

        ProgressBar1.Visible = False
        ProgressBar1.Value = 0
        Dim s() As String
        Dim path As String
        Dim filename As String
        s = Split(OpenFileDialog1.FileName, "\")
        filename = s(UBound(s))
        ListBox1.Items.Add(" ")
        path = OpenFileDialog1.FileName.Replace(".xls", "处理结果.xls")
        excel.DisplayAlerts = False


        ListBox1.Items.Insert(0, "处理结果保存在： " + path)
        excel.Workbooks(1).SaveAs(path)
        excel.Workbooks(1).Close()
        excel.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        excel = Nothing
        GC.Collect()
        Windows.Forms.Cursor.Current = Cursors.Default
        If ESC = False Then
            ListBox1.Items.Insert(0, "固定资产比较完毕！")
        End If

    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Cursor.Current = Cursors.WaitCursor
        Dim regionName As String
        Dim ConnectString As String = "Password=www518;Persist Security Info=True;User ID=sa;Initial Catalog=iposter;Data Source=192.168.1.9"
        Dim myCommand As New SqlClient.SqlCommand
        Dim mySqlConc As New SqlClient.SqlConnection

        mySqlConc.ConnectionString = ConnectString
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc


        ListBox1.Items.Clear()
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer = 4
            Dim index As Integer = 1
            excel.Workbooks.Open(OpenFileDialog1.FileName)
            Dim xx() As String
            OrgSheet = excel.Worksheets(1)
            With (OrgSheet)
                While (True)
                    Application.DoEvents()
                    If ESC Then
                        Exit While
                    End If

                    regionName = Trim(.Cells(myrow, 2).value)
                    If regionName = "" Then
                        myrow += 1
                        Continue While
                    End If
                    If myrow > 522 Then
                        Exit While
                    End If
                    xx = Split(regionName, " ")
                    If index > 500 Then
                        myCommand.CommandText = String.Format("insert into 选择信息 (二级行业) values('{0}') ", xx(0))
                        If myCommand.ExecuteNonQuery() = 0 Then
                            ListBox1.Items.Add("导入失败 " + xx(0))
                        Else
                            ListBox1.Items.Add("导入 " + xx(0))
                        End If
                    Else
                        myCommand.CommandText = String.Format("Update 选择信息 set 二级行业 = '{0}' where ID = {1}", xx(0), index)
                        If myCommand.ExecuteNonQuery() = 0 Then
                            ListBox1.Items.Add("导入失败 " + xx(0))
                        Else
                            ListBox1.Items.Add("导入 " + xx(0))
                        End If
                    End If

                    index += 1
                    myrow += 1
                End While
            End With
            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
        End If

        mySqlConc.Close()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Cursor.Current = Cursors.WaitCursor
        Dim ConnectString As String = "Password=www518;Persist Security Info=True;User ID=sa;Initial Catalog=iposter;Data Source=192.168.1.9"
        Dim myCommand1 As New SqlClient.SqlCommand
        Dim mySqlConc1 As New SqlClient.SqlConnection
        Dim myReader1 As SqlClient.SqlDataReader

        mySqlConc1.ConnectionString = ConnectString
        Try
            mySqlConc1.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand1.Connection = mySqlConc1

        Dim myCommand As New SqlClient.SqlCommand
        Dim mySqlConc As New SqlClient.SqlConnection


        Dim skrq As Date
        Dim xkrq As Date

        Dim yjid As Integer
        Dim hthao2 As String
        Dim leixing, hthao, khmc, hbmc, tfqy, quyu1, htzt, dkqk, beizhu, scyq, bbzt, lxr, sqrm, tel As String
        Dim hbsc As Decimal
        Dim hbpc As Integer
        Dim htks, htjs, sqsj As Date
        hthao2 = ""

        mySqlConc.ConnectionString = ConnectString
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc

        myCommand1.CommandText = "select * from 电影海报视图A order by 合同号"
        myReader1 = myCommand1.ExecuteReader()

        While myReader1.Read()
            leixing = Trim(myReader1.GetString(1))
            hthao = Trim(myReader1.GetString(2))
            khmc = Trim(myReader1.GetString(3))
            hbmc = Trim(myReader1.GetString(4))
            tfqy = Trim(myReader1.GetString(5))
            If myReader1.IsDBNull(6) Then
                quyu1 = ""
            Else
                quyu1 = Trim(myReader1.GetString(6))
            End If
            hbsc = myReader1.GetDecimal(7)
            hbpc = CInt(myReader1.GetString(8))

            htks = myReader1.GetDateTime(9)
            htjs = myReader1.GetDateTime(10)
            If myReader1.IsDBNull(11) Then
                htzt = ""
            Else
                htzt = Trim(myReader1.GetString(11))
            End If
            If myReader1.IsDBNull(12) Then
                dkqk = ""
            Else
                dkqk = Trim(myReader1.GetString(12))
            End If
            If myReader1.IsDBNull(13) Then
                beizhu = ""
            Else
                beizhu = Trim(myReader1.GetString(13))
            End If
            sqrm = Trim(myReader1.GetString(17))
            sqsj = myReader1.GetDateTime(18)
            If myReader1.IsDBNull(19) Then
                scyq = ""
            Else
                scyq = Trim(myReader1.GetString(19))
            End If
            If myReader1.IsDBNull(20) Then
                bbzt = ""
            Else
                bbzt = Trim(myReader1.GetString(20))
            End If
            skrq = myReader1.GetDateTime(21)
            xkrq = myReader1.GetDateTime(22)
            If myReader1.IsDBNull(23) Then
                lxr = ""
            Else
                lxr = Trim(myReader1.GetString(23))
            End If
            If myReader1.IsDBNull(24) Then
                tel = ""
            Else
                tel = Trim(myReader1.GetString(24))
            End If
            If hthao <> hthao2 Then
                myCommand.CommandText = String.Format("insert into 业绩描述表 (合同号,申请时间,所属区域,媒体类别,发布类别,资源情况,销售姓名,所属部门,客户类型,签约客户名称,客户联系方式,投放客户名称,一级行业,合同起始时间,合同结束时间,监播要求,合同状态,版本状态,到款比例,备注,特殊签批,部门领导,区域总经理,CEO签批,合规部,媒介执行,媒介总经理,行政签批,财务签批,到款日期) values( '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}',{18},'{19}',{20},{21},{22},{23},{24},{25},{26},{27},{28},'{29}')", _
                                                      hthao, sqsj, "北京", "影广", leixing, "全资源", sqrm, "影视事业部", "直客", khmc, lxr + tel, khmc, "影片宣发", htks, htjs, "", htzt, bbzt, 0, beizhu, 1, 1, 1, 1, 1, 1, 1, 1, 1, "1900-01-01")
                myCommand.ExecuteNonQuery()

                myCommand.CommandText = "SELECT @@IDENTITY AS 'Identity'"
                yjid = myCommand.ExecuteScalar
            End If

            myCommand.CommandText = String.Format("insert into 发布时间表 (业绩表ID,发布时间,截止时间,发布时长,发布频次,发布版本,发布区域名,发布区域码) values( {0},'{1}','{2}',{3},{4},'{5}','{6}','{7}')", _
                                                 yjid, skrq, xkrq, hbsc, hbpc, hbmc, quyu1, tfqy)
            myCommand.ExecuteNonQuery()

            hthao2 = hthao
            '  Exit While
        End While
        myReader1.Close()


        mySqlConc.Close()
        mySqlConc1.Close()

        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Cursor.Current = Cursors.WaitCursor
        Dim ConnectString As String = "Password=www518;Persist Security Info=True;User ID=sa;Initial Catalog=iposter;Data Source=192.168.1.9"
        Dim myCommand1 As New SqlClient.SqlCommand
        Dim mySqlConc1 As New SqlClient.SqlConnection
        Dim myReader1 As SqlClient.SqlDataReader

        mySqlConc1.ConnectionString = ConnectString
        Try
            mySqlConc1.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand1.Connection = mySqlConc1

        Dim myCommand As New SqlClient.SqlCommand
        Dim mySqlConc As New SqlClient.SqlConnection


        Dim skrq As Date
        Dim xkrq As Date

        Dim yjid As Integer
        Dim hthao2 As String = ""
        Dim leixing, hthao, khmc, hbmc, tfqy, quyu1, htzt, dkqk, beizhu, scyq, bbzt, lxr, sqrm, tel, bm As String
        Dim hbsc As Decimal
        Dim hbpc As Integer
        Dim htks, htjs, sqsj As Date


        mySqlConc.ConnectionString = ConnectString
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc

        myCommand1.CommandText = "select * from 广告发布视图 order by 合同号"
        myReader1 = myCommand1.ExecuteReader()

        While myReader1.Read()
            leixing = Trim(myReader1.GetString(1))
            hthao = Trim(myReader1.GetString(2))
            khmc = Trim(myReader1.GetString(3))
            hbmc = Trim(myReader1.GetString(4))
            tfqy = Trim(myReader1.GetString(5))
            If myReader1.IsDBNull(6) Then
                quyu1 = tfqy
            Else
                quyu1 = Trim(myReader1.GetString(6))
            End If
            hbsc = myReader1.GetDecimal(7)
            hbpc = CInt(myReader1.GetString(8))

            htks = myReader1.GetDateTime(9)
            htjs = myReader1.GetDateTime(10)
            If myReader1.IsDBNull(11) Then
                htzt = ""
            Else
                htzt = Trim(myReader1.GetString(11))
            End If
            If myReader1.IsDBNull(14) Then
                dkqk = ""
            Else
                dkqk = Trim(myReader1.GetString(14))
            End If
            If myReader1.IsDBNull(16) Then
                beizhu = ""
            Else
                beizhu = Trim(myReader1.GetString(16))
            End If
            sqrm = Trim(myReader1.GetString(20))
            sqsj = myReader1.GetDateTime(21)
            If myReader1.IsDBNull(22) Then
                scyq = ""
            Else
                scyq = Trim(myReader1.GetString(22))
            End If
            If myReader1.IsDBNull(23) Then
                bbzt = ""
            Else
                bbzt = Trim(myReader1.GetString(23))
            End If
            skrq = myReader1.GetDateTime(24)
            xkrq = myReader1.GetDateTime(25)
            If myReader1.IsDBNull(26) Then
                lxr = ""
            Else
                lxr = Trim(myReader1.GetString(26))
            End If
            If myReader1.IsDBNull(27) Then
                tel = ""
            Else
                tel = Trim(myReader1.GetString(27))
            End If
            bm = Trim(myReader1.GetString(28)).Replace("商广", "")

            If hthao <> hthao2 Then
                myCommand.CommandText = String.Format("insert into 业绩描述表 (合同号,申请时间,所属区域,媒体类别,发布类别,资源情况,销售姓名,所属部门,客户类型,签约客户名称,客户联系方式,投放客户名称,一级行业,合同起始时间,合同结束时间,监播要求,合同状态,版本状态,到款比例,备注,特殊签批,部门领导,区域总经理,CEO签批,合规部,媒介执行,媒介总经理,行政签批,财务签批,到款日期) values( '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}',{18},'{19}',{20},{21},{22},{23},{24},{25},{26},{27},{28},'{29}')", _
                                                      hthao, sqsj, bm, "商广", leixing, "全资源", sqrm, "销售部", "直客", khmc, lxr + tel, khmc, "", htks, htjs, "", htzt, bbzt, 0, beizhu, 1, 1, 1, 1, 1, 1, 1, 1, 1, "1900-01-01")
                myCommand.ExecuteNonQuery()

                myCommand.CommandText = "SELECT @@IDENTITY AS 'Identity'"
                yjid = myCommand.ExecuteScalar
            End If

            myCommand.CommandText = String.Format("insert into 发布时间表 (业绩表ID,发布时间,截止时间,发布时长,发布频次,发布版本,发布区域名,发布区域码) values( {0},'{1}','{2}',{3},{4},'{5}','{6}','{7}')", _
                                                 yjid, skrq, xkrq, hbsc, hbpc, hbmc, quyu1, tfqy)
            myCommand.ExecuteNonQuery()

            hthao2 = hthao
            '  Exit While
        End While
        myReader1.Close()


        mySqlConc.Close()
        mySqlConc1.Close()

        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        If TextBox2.Text <> "58707780" Then
            MsgBox("验证码错误！")
            Return
        End If

        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始下载过期资源列表......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=0&endDate=" + Now().ToShortDateString

            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)

            str = "http://pv.tihe-china.com/manager/php/city.php?cmd=ListCity&limit=-1"
            data = wc.DownloadData(str)
            Dim citys As String = System.Text.Encoding.UTF8.GetString(data)
            Dim xx() As String
            xx = Split(citys, "[{")
            citys = xx(1).Replace("}]}", "")
            xx = Split(citys, "},{")


            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim prov0, prov1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim panel As Integer = 0
            Dim paneltotal As Integer = 0
            Dim panel2 As Integer = 0
            Dim paneltotal2 As Integer = 0
            Dim theater As Integer = 0
            Dim theatertotal As Integer = 0
            Dim seat As Integer = 0
            Dim seattotal As Integer = 0
            Dim renci As Integer = 0
            Dim rencitotal As Integer = 0
            Dim piaofang As Integer = 0
            Dim piaofangtotal As Integer = 0
            Dim sumrow As Integer = 0
            Dim provrow As Integer = 5
            Dim row As Integer = 4
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()


            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "销售资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 60
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:L").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:L1").Merge()

                .Range("B2:L2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:L2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:L2").Font.Bold = True
                .Range("B2:L2").Font.Size = 18
                .Range("B2:L2").Merge()

                .Range("B4:M4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:M4").Font.Bold = True
                .Range("A4:M2000").Font.Size = 10
                .Range("B4:M4").Font.Color = Color.White
                .Range("B4:M4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:M4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:M4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "附近写字楼、附近商场、百货"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "影厅数"
                .Cells(4, 10) = "座位数"
                .Cells(4, 11) = "人次"
                .Cells(4, 12) = "票房"
                .Cells(4, 13) = "终止日期"

                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CheckBox1.Checked Then
                        If run_time(zzz(28), NumericUpDown1.Value) = False Then
                            Continue For
                        End If
                    End If
                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = theater
                            .Cells(sumrow, 10) = seat
                            .Cells(sumrow, 11) = renci
                            .Cells(sumrow, 12) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(20)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(8)
                    .Cells(row, 10) = zzz(9)
                    .Cells(row, 11) = zzz(10)
                    .Cells(row, 12) = zzz(11)
                    .Cells(row, 13) = zzz(16)
                    If zzz(20) = 1 Then
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":M" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = theater
                    .Cells(sumrow, 10) = seat
                    .Cells(sumrow, 11) = renci
                    .Cells(sumrow, 12) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:L3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:L3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = theatertotal
                .Cells(3, 10) = seattotal
                .Cells(3, 11) = rencitotal
                .Cells(3, 12) = piaofangtotal
                .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "过期影院资源列表" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                .Range("A4:M" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("I5:I" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:M" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:M" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            ListBox1.Items.Add("正在保存过期资源列表.......")
            excel.Worksheets(1).select()
            excel.DisplayAlerts = False
            If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "过期影院资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "过期影院资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\过期影院资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\过期影院资源列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            End If

            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("过期资源列表下载完毕！")
        End If
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始下载资源列表......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If RadioButton1.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=0"
            End If
            If RadioButton2.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=1"
            End If
            If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                str = "http://pv.tihe-china.com/manager/php/region.php?cmd=PvshowListRegionZYA&limit=-1&working_mode=0&tradingId=2"
            End If
            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)

            str = "http://pv.tihe-china.com/manager/php/city.php?cmd=ListCity&limit=-1"
            data = wc.DownloadData(str)
            Dim citys As String = System.Text.Encoding.UTF8.GetString(data)
            Dim xx() As String
            xx = Split(citys, "[{")
            citys = xx(1).Replace("}]}", "")
            xx = Split(citys, "},{")


            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim prov0, prov1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim panel As Integer = 0
            Dim paneltotal As Integer = 0
            Dim panel2 As Integer = 0
            Dim paneltotal2 As Integer = 0
            Dim theater As Integer = 0
            Dim theatertotal As Integer = 0
            Dim seat As Integer = 0
            Dim seattotal As Integer = 0
            Dim renci As Integer = 0
            Dim rencitotal As Integer = 0
            Dim piaofang As Integer = 0
            Dim piaofangtotal As Integer = 0
            Dim sumrow As Integer = 0
            Dim provrow As Integer = 5
            Dim row As Integer = 4
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()


            city0 = ""
            city1 = ""
            prov0 = ""
            prov1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "销售资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:D").ColumnWidth = 8.38
                .Columns("E:E").ColumnWidth = 28.88
                .Columns("F:F").ColumnWidth = 47.5
                .Columns("G:G").ColumnWidth = 60
                .Columns("H:H").ColumnWidth = 7.5
                .Columns("I:L").ColumnWidth = 7

                .Rows("1:1").RowHeight = 14.25
                .Rows("2:2").RowHeight = 86.25
                .Rows("3:3").RowHeight = 16.5
                .Rows("4:4").RowHeight = 14.25
                .Rows("5:2000").RowHeight = 16.5

                .Range("A1:L1").Merge()

                .Range("B2:L2").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:L2").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B2:L2").Font.Bold = True
                .Range("B2:L2").Font.Size = 18
                .Range("B2:L2").Merge()

                .Range("B4:L4").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:L4").VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("B4:L4").Font.Bold = True
                .Range("A4:L2000").Font.Size = 10
                .Range("B4:L4").Font.Color = Color.White
                .Range("B4:L4").Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Range("B4:L4").Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Range("B4:L4").Interior.Color = 812276



                '          .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国单影院资源列表" + Chr(13) + Chr(10)
                .Cells(4, 2) = "省份"
                .Cells(4, 3) = "城市"
                .Cells(4, 4) = "序号"
                .Cells(4, 5) = "影院名称"
                .Cells(4, 6) = "地址"
                .Cells(4, 7) = "点位"
                .Cells(4, 8) = "屏数"
                .Cells(4, 9) = "影厅数"
                .Cells(4, 10) = "座位数"
                .Cells(4, 11) = "人次"
                .Cells(4, 12) = "票房"

                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If CheckBox1.Checked Then
                        If run_time(zzz(28), NumericUpDown1.Value) = False Then
                            Continue For
                        End If
                    End If
                    city1 = zzz(19)
                    prov1 = GetProv(city1, xx)

                    If prov0 <> prov1 Then
                        If prov0 <> "" Then
                            .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                            .Cells(provrow.ToString, 2) = prov0
                            provrow = row + 1
                        End If
                    End If

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Cells(sumrow, 3) = "1"
                            .Cells(sumrow, 4) = "小计："
                            .Cells(sumrow, 5) = region
                            .Cells(sumrow, 8) = panel
                            .Cells(sumrow, 9) = theater
                            .Cells(sumrow, 10) = seat
                            .Cells(sumrow, 11) = renci
                            .Cells(sumrow, 12) = piaofang
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                            .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 3) = city0
                        End If
                        sumcity = sumcity + 1
                        row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                        paneltotal = paneltotal + panel
                        panel = 0
                        theatertotal = theatertotal + theater
                        theater = 0
                        seattotal = seattotal + seat
                        seat = 0
                        rencitotal = rencitotal + renci
                        renci = 0
                        piaofangtotal = piaofangtotal + piaofang
                        piaofang = 0
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.ThemeColor = 7
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternTintAndShade = 0
                    End If


                    row = row + 1
                    region = region + 1
                    If IsNumeric(zzz(2)) Then
                        panel = panel + CInt(zzz(2))
                    End If
                    If IsNumeric(zzz(8)) Then
                        theater = theater + CInt(zzz(8))
                    End If
                    If IsNumeric(zzz(9)) Then
                        seat = seat + CInt(zzz(9))
                    End If
                    If IsNumeric(zzz(10)) Then
                        renci = renci + CInt(zzz(10))
                    End If
                    If IsNumeric(zzz(11)) Then
                        piaofang = piaofang + CInt(zzz(11))
                    End If
                    '   .Cells(row, 2) = prov1
                    '   .Cells(row, 3) = zzz(18)
                    .Cells(row, 4) = region
                    .Cells(row, 5) = zzz(3)
                    .Cells(row, 6) = zzz(5)
                    .Cells(row, 7) = zzz(7)
                    .Cells(row, 8) = zzz(2)
                    .Cells(row, 9) = zzz(8)
                    .Cells(row, 10) = zzz(9)
                    .Cells(row, 11) = zzz(10)
                    .Cells(row, 12) = zzz(11)
                    If zzz(20) = 1 Then
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.ThemeColor = 8
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":L" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    city0 = city1
                    prov0 = prov1
                Next



                If prov0 <> "" Then
                    .Range("B" + provrow.ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + provrow.ToString + ":B" + row.ToString).Merge()
                    .Cells(provrow.ToString, 2) = prov0
                End If



                If city0 <> "" Then
                    .Cells(sumrow, 3) = "1"
                    .Cells(sumrow, 4) = "小计："
                    .Cells(sumrow, 5) = region
                    .Cells(sumrow, 8) = panel
                    .Cells(sumrow, 9) = theater
                    .Cells(sumrow, 10) = seat
                    .Cells(sumrow, 11) = renci
                    .Cells(sumrow, 12) = piaofang
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Bold = True
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Font.Size = 10
                    .Range("C" + (sumrow + 1).ToString + ":C" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 3) = city0
                End If

                regiontotal = regiontotal + region
                paneltotal = paneltotal + panel
                theatertotal = theatertotal + theater
                seattotal = seattotal + seat
                rencitotal = rencitotal + renci
                piaofangtotal = piaofangtotal + piaofang
                .Range("A3:L3").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A3:L3").Font.Bold = True
                .Cells(3, 3) = sumcity
                .Cells(3, 5) = regiontotal
                .Cells(3, 8) = paneltotal
                .Cells(3, 9) = theatertotal
                .Cells(3, 10) = seattotal
                .Cells(3, 11) = rencitotal
                .Cells(3, 12) = piaofangtotal
                If RadioButton1.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton2.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表(储备）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                If RadioButton3.Checked = True Or RadioButton4.Checked = True Then
                    .Cells(2, 2) = "泰和数码海报" + Chr(13) + Chr(10) + "全国影院资源列表（正式+储备）" + Chr(13) + Chr(10) + mydate.ToShortDateString + "（" + sumcity.ToString + "城市" + regiontotal.ToString + "家）"
                End If
                .Range("A4:L" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("E5:E" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("F5:F" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("G5:G" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("I5:I" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("B4:L" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A5").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            ListBox1.Items.Add("正在保存资源点位列表.......")
            excel.Worksheets(1).select()
            excel.DisplayAlerts = False
            If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "全国影院资源点位列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "全国影院资源点位列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\全国影院资源点位列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\全国影院资源点位列表" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            End If

            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("资源点位列表下载完毕！")
        End If
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Cursor.Current = Cursors.WaitCursor
        Dim myCommand As New OdbcCommand
        Dim mySqlConc As New OdbcConnection
        Dim renci As String
        Dim piaofan As String
        Dim regionName As String

        Dim str As String = "Driver={MariaDB ODBC 2.0 Driver};Server=pv.tihe-china.com;Database=pvmanager_new;User=root; Password=519618;"

        mySqlConc.ConnectionString = str
        Try
            mySqlConc.Open()
        Catch ex As Exception
            MsgBox("数据库连接错误！")
            Return
        End Try
        myCommand.Connection = mySqlConc
        ListBox1.Items.Clear()
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim myrow As Integer = 2
            excel.Workbooks.Open(OpenFileDialog1.FileName)

            OrgSheet = excel.Worksheets(1)
            With (OrgSheet)
                While (True)
                    Application.DoEvents()
                    If ESC Then
                        Exit While
                    End If
                    piaofan = Trim(.Cells(myrow, 4).value)
                    renci = Trim(.Cells(myrow, 5).value)
                    regionName = Trim(.Cells(myrow, 6).value)
 
                    If Trim(.Cells(myrow, 3).value) = "" Then
                        Exit While
                    End If
                    If regionName = "" Then
                        myrow += 1
                        Continue While
                    End If
                    '    ListBox1.Items.Add("开始处理 " + regionName)
                    ' myCommand.CommandText = String.Format("Update region set fr_stat = {0},renci = {1} ", 0, 0)
                    ' myCommand.ExecuteNonQuery()
                    ' Exit While
                    myCommand.CommandText = String.Format("Update region set fr_stat = {0},renci = {1} where EnName like '%{2}%'", CInt(piaofan), CInt(renci), regionName)
                    If myCommand.ExecuteNonQuery() = 0 Then
                        ListBox1.Items.Add("导入失败 " + .Cells(myrow, 3).value + regionName)
                    Else
                        ListBox1.Items.Add("导入 " + .Cells(myrow, 3).value + regionName)
                    End If
                    myrow += 1
                End While
            End With
            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
        End If
        ListBox1.Items.Add("导入完毕！")
        mySqlConc.Close()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始下载DSP_ID......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            str = "http://pv.tihe-china.com/manager/php/region.php?cmd=ListZM"

            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)


            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim sumrow As Integer = 0
            '  Dim provrow As Integer = 5
            Dim row As Integer = 1
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()
            Dim yyid1, yyid2 As String

            yyid1 = ""
            yyid2 = ""


            city0 = ""
            city1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "影院资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 5
                .Columns("D:D").ColumnWidth = 60
                .Columns("E:E").ColumnWidth = 20

                .Rows("1:1").RowHeight = 30
                .Rows("2:2000").RowHeight = 20
                .Range("A1:L1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A1:E1").Font.Bold = True

                .Cells(1, 1) = "序号"
                .Cells(1, 2) = "城市"
                .Cells(1, 3) = "序号"
                .Cells(1, 4) = "影院名称"
                .Cells(1, 5) = "影院ID"
               
                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
   
                    city1 = zzz(0)

                    If city0 <> city1 Then
                        If city0 <> "" Then                          
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            '   .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 2) = city0
                        End If
                        '   sumcity = sumcity + 1
                        '   row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                    End If

                    yyid1 = zzz(1)
                    If yyid1 = yyid2 Then
                        ListBox1.Items.Add("重复ID：" + yyid1)
                    End If
                    row = row + 1
                    region = region + 1
                    .Cells(row, 1) = row - 1
                    .Cells(row, 3) = region
                    '  .Cells(row, 4) = zzz(2) + "(" + zzz(1) + ")"
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = zzz(1)
                    .Cells(row, 6) = zzz(3)
                    If InStr(zzz(4), "no program") > 0 Then
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.ThemeColor = 5
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    If InStr(zzz(4), "xml") > 0 Then
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.ThemeColor = 6
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    yyid2 = yyid1
                    city0 = city1
                Next

                If city0 <> "" Then
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    ' .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 2) = city0
                End If


                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A2").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0


            ListBox1.Items.Add("开始下载已安装资源列表......")
            '已安装---------------------------------------------------
            mySqlConc.ConnectionString = dbstr
            Try
                mySqlConc.Open()
            Catch ex As Exception
                MsgBox("数据库连接错误！")
                Return
            End Try
            myCommand.Connection = mySqlConc
            Dim mcount(2000) As Integer
            Dim mmac(2000) As String
            Dim index As Integer = 0
            myCommand.CommandText = String.Format("SELECT count(ID) as ct,mac FROM zmcount WHERE tracking=1 AND ptime > '{0}' AND ptime < '{1}' AND dsp='zm' group by mac", DateTimePicker1.Value.ToShortDateString, DateTimePicker2.Value.AddDays(1).ToShortDateString)
            myReader = myCommand.ExecuteReader()
            While myReader.Read()
                mcount(index) = myReader.GetInt64(0)
                mmac(index) = Trim(myReader.GetString(1))
                index += 1
            End While
            myReader.Close()
            mySqlConc.Close()
            city0 = ""
            city1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0 
            sumrow = 0
            row = 1
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("影院资源列表"))

            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "已安装资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 5
                .Columns("D:D").ColumnWidth = 60
                .Columns("E:E").ColumnWidth = 20
                .Columns("F:F").ColumnWidth = 10

                .Rows("1:1").RowHeight = 30
                .Rows("2:2000").RowHeight = 20
                .Range("A1:L1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A1:F1").Font.Bold = True

                .Cells(1, 1) = "序号"
                .Cells(1, 2) = "城市"
                .Cells(1, 3) = "序号"
                .Cells(1, 4) = "影院名称"
                .Cells(1, 5) = "影院ID"
                .Cells(1, 6) = "展现次数"

                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
    
                    city1 = zzz(0)


                    If city0 <> city1 Then
                        If city0 <> "" And region > 0 Then
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            '   .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 2) = city0
                        End If
                        '   sumcity = sumcity + 1
                        '   row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                    End If

                    yyid1 = zzz(1)

                    yyid2 = yyid1
                    city0 = city1

                    If Trim(zzz(4)) = "" Then
                        Continue For
                    End If
                    row = row + 1
                    region = region + 1
                    .Cells(row, 1) = row - 1
                    .Cells(row, 3) = region
                    ' .Cells(row, 4) = zzz(2) + "(" + zzz(1) + ")"
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = zzz(1)
                    '       .Cells(row, 6) = 0
                    For j As Integer = 0 To 1999
                        If mmac(j) = zzz(1) Then
                            .Cells(row, 6) = mcount(j)
                            Exit For
                        End If
                    Next
                Next

                If city0 <> "" And region > 0 Then
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    ' .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 2) = city0
                End If


                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A2").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载运行资源列表......")
            '已安装---------------------------------------------------

            city0 = ""
            city1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            sumrow = 0
            row = 1
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("已安装资源列表"))

            OrgSheet = excel.Worksheets(3)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "运行资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 5
                .Columns("D:D").ColumnWidth = 60
                .Columns("E:E").ColumnWidth = 20
                .Columns("F:F").ColumnWidth = 10

                .Rows("1:1").RowHeight = 30
                .Rows("2:2000").RowHeight = 20
                .Range("A1:L1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A1:F1").Font.Bold = True

                .Cells(1, 1) = "序号"
                .Cells(1, 2) = "城市"
                .Cells(1, 3) = "序号"
                .Cells(1, 4) = "影院名称"
                .Cells(1, 5) = "影院ID"
                .Cells(1, 6) = "展现次数"

                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next

                    city1 = zzz(0)


                    If city0 <> city1 Then
                        If city0 <> "" And region > 0 Then
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            '   .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 2) = city0
                        End If
                        '   sumcity = sumcity + 1
                        '   row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                    End If

                    yyid1 = zzz(1)

                    yyid2 = yyid1
                    city0 = city1

                    If InStr(zzz(4), "xml") <= 0 Then
                        Continue For
                    End If
                    row = row + 1
                    region = region + 1
                    .Cells(row, 1) = row - 1
                    .Cells(row, 3) = region
                    ' .Cells(row, 4) = zzz(2) + "(" + zzz(1) + ")"
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = zzz(1)
                    '   .Cells(row, 6) = 0
                    For j As Integer = 0 To 1999
                        If mmac(j) = zzz(1) Then
                            .Cells(row, 6) = mcount(j)
                            Exit For
                        End If
                    Next
                Next

                If city0 <> "" And region > 0 Then
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    ' .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 2) = city0
                End If


                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A2").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            mydate = DateTimePicker1.Value
                ListBox1.Items.Add("正在保存DSP_ID.......")
                excel.Worksheets(1).select()
                excel.DisplayAlerts = False
                If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "众盟_ID" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "众盟_ID" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\众盟_ID" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\众盟_ID" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                End If

                excel.Workbooks(1).Close()
                excel.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
                excel = Nothing
                GC.Collect()
                Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("众盟_ID下载完毕！")
        End If
    End Sub
    Public Function MD5(ByVal StrSource As String, ByVal Code As Integer) As String

        Dim str As String = ""



        Dim md5Hasher As New MD5CryptoServiceProvider()


        Dim data As Byte() = md5Hasher.ComputeHash(Encoding.Default.GetBytes(StrSource))


        Dim sBuilder As New StringBuilder()


        Dim i As Integer
        For i = 0 To data.Length - 1
            sBuilder.Append(data(i).ToString("x2"))
        Next i


        Select Case Code
            Case 16
                str = sBuilder.ToString().Substring(0, 16)
            Case 32
                str = sBuilder.ToString().Substring(0, 32)
        End Select


        Return str


    End Function

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始下载DSP_ID......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            str = "http://pv.tihe-china.com/manager/php/region.php?cmd=ListZM"

            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)


            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim sumrow As Integer = 0
            '  Dim provrow As Integer = 5
            Dim row As Integer = 1
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()
            Dim yyid1, yyid2 As String

            yyid1 = ""
            yyid2 = ""


            city0 = ""
            city1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "影院资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 5
                .Columns("D:D").ColumnWidth = 60
                .Columns("E:E").ColumnWidth = 20

                .Rows("1:1").RowHeight = 30
                .Rows("2:2000").RowHeight = 20
                .Range("A1:L1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A1:E1").Font.Bold = True

                .Cells(1, 1) = "序号"
                .Cells(1, 2) = "城市"
                .Cells(1, 3) = "序号"
                .Cells(1, 4) = "影院名称"
                .Cells(1, 5) = "影院ID"

                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next

                    city1 = zzz(0)

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            '   .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 2) = city0
                        End If
                        '   sumcity = sumcity + 1
                        '   row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                    End If

                    yyid1 = zzz(1)
                    If yyid1 = yyid2 Then
                        ListBox1.Items.Add("重复ID：" + yyid1)
                    End If
                    row = row + 1
                    region = region + 1
                    .Cells(row, 1) = row - 1
                    .Cells(row, 3) = region
                    '  .Cells(row, 4) = zzz(2) + "(" + zzz(1) + ")"
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = "2098" + zzz(1)
                    .Cells(row, 6) = zzz(3)
                    If InStr(zzz(4), "no program") > 0 Then
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.ThemeColor = 5
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    If InStr(zzz(4), "xml") > 0 Then
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.ThemeColor = 6
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.TintAndShade = 0.599993896298105
                        .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternTintAndShade = 0
                    End If
                    yyid2 = yyid1
                    city0 = city1
                Next

                If city0 <> "" Then
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    ' .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 2) = city0
                End If


                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A2").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0


            ListBox1.Items.Add("开始下载已安装资源列表......")
            '已安装---------------------------------------------------
            mySqlConc.ConnectionString = dbstr
            Try
                mySqlConc.Open()
            Catch ex As Exception
                MsgBox("数据库连接错误！")
                Return
            End Try
            myCommand.Connection = mySqlConc
            Dim mcount(2000) As Integer
            Dim mmac(2000) As String
            Dim index As Integer = 0
            myCommand.CommandText = String.Format("SELECT count(ID) as ct,mac FROM zmcount WHERE tracking=1 AND ptime > '{0}' AND ptime < '{1}' AND dsp='dp' group by mac", DateTimePicker1.Value.ToShortDateString, DateTimePicker2.Value.AddDays(1).ToShortDateString)
            myReader = myCommand.ExecuteReader()
            While myReader.Read()
                mcount(index) = myReader.GetInt64(0)
                mmac(index) = Trim(myReader.GetString(1))
                index += 1
            End While
            myReader.Close()
            mySqlConc.Close()

            city0 = ""
            city1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            sumrow = 0
            row = 1
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("影院资源列表"))

            OrgSheet = excel.Worksheets(2)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "已安装资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 5
                .Columns("D:D").ColumnWidth = 60
                .Columns("E:E").ColumnWidth = 20
                .Columns("F:F").ColumnWidth = 10

                .Rows("1:1").RowHeight = 30
                .Rows("2:2000").RowHeight = 20
                .Range("A1:L1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A1:F1").Font.Bold = True

                .Cells(1, 1) = "序号"
                .Cells(1, 2) = "城市"
                .Cells(1, 3) = "序号"
                .Cells(1, 4) = "影院名称"
                .Cells(1, 5) = "影院ID"
                .Cells(1, 6) = "展现次数"

                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next

                    city1 = zzz(0)


                    If city0 <> city1 Then
                        If city0 <> "" And region > 0 Then
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            '   .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 2) = city0
                        End If
                        '   sumcity = sumcity + 1
                        '   row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                    End If

                    yyid1 = zzz(1)

                    yyid2 = yyid1
                    city0 = city1

                    If Trim(zzz(4)) = "" Then
                        Continue For
                    End If
                    row = row + 1
                    region = region + 1
                    .Cells(row, 1) = row - 1
                    .Cells(row, 3) = region
                    ' .Cells(row, 4) = zzz(2) + "(" + zzz(1) + ")"
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = "2098" + zzz(1)
                    For j As Integer = 0 To 1999
                        If mmac(j) = zzz(1) Then
                            .Cells(row, 6) = mcount(j)
                            Exit For
                        End If
                    Next
                Next

                If city0 <> "" And region > 0 Then
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    ' .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 2) = city0
                End If


                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A2").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0

            ListBox1.Items.Add("开始下载运行资源列表......")
            '已安装---------------------------------------------------

            city0 = ""
            city1 = ""
            region = 0
            regiontotal = 0
            citycount = 0
            citytotal = 0
            sumrow = 0
            row = 1
            sumcity = 0

            excel.Worksheets.Add(After:=excel.Worksheets("已安装资源列表"))

            OrgSheet = excel.Worksheets(3)
            excel.ActiveWindow.DisplayGridlines = False
            With OrgSheet
                .Name = "运行资源列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 5
                .Columns("D:D").ColumnWidth = 60
                .Columns("E:E").ColumnWidth = 20
                .Columns("F:F").ColumnWidth = 10

                .Rows("1:1").RowHeight = 30
                .Rows("2:2000").RowHeight = 20
                .Range("A1:L1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A1:F1").Font.Bold = True

                .Cells(1, 1) = "序号"
                .Cells(1, 2) = "城市"
                .Cells(1, 3) = "序号"
                .Cells(1, 4) = "影院名称"
                .Cells(1, 5) = "影院ID"
                .Cells(1, 6) = "展现次数"

                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next

                    city1 = zzz(0)


                    If city0 <> city1 Then
                        If city0 <> "" And region > 0 Then
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            '   .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 2) = city0
                        End If
                        '   sumcity = sumcity + 1
                        '   row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                    End If

                    yyid1 = zzz(1)

                    yyid2 = yyid1
                    city0 = city1

                    If InStr(zzz(4), "xml") <= 0 Then
                        Continue For
                    End If
                    row = row + 1
                    region = region + 1
                    .Cells(row, 1) = row - 1
                    .Cells(row, 3) = region
                    ' .Cells(row, 4) = zzz(2) + "(" + zzz(1) + ")"
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = "2098" + zzz(1)
                    For j As Integer = 0 To 1999
                        If mmac(j) = zzz(1) Then
                            .Cells(row, 6) = mcount(j)
                            Exit For
                        End If
                    Next
                Next

                If city0 <> "" And region > 0 Then
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    ' .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 2) = city0
                End If


                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:F" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A2").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            mydate = DateTimePicker1.Value
            ListBox1.Items.Add("正在保存DSP_ID.......")
            excel.Worksheets(1).select()
            excel.DisplayAlerts = False
            If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "点屏_ID" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "点屏_ID" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\点屏_ID" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\点屏_ID" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            End If

            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("点屏_ID下载完毕！")
        End If
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        On Error Resume Next
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With
        With OpenFileDialog2
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If OpenFileDialog2.ShowDialog() = Windows.Forms.DialogResult.OK Then
                ESC = False
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                ListBox1.Items.Clear()

                ProgressBar1.Maximum = 6000
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                Dim excel As New Microsoft.Office.Interop.Excel.Application()
                Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

                Dim excel1 As New Microsoft.Office.Interop.Excel.Application()
                Dim OrgSheet1 As Microsoft.Office.Interop.Excel.Worksheet

                Dim myrow As Integer = 2
                Dim flag As Boolean
                Dim mysn As String
                Dim index As Integer
                Dim tt As String

                excel.Workbooks.Open(OpenFileDialog1.FileName)
                OrgSheet = excel.Worksheets(1)
                excel1.Workbooks.Open(OpenFileDialog2.FileName)
                For i As Integer = 1 To excel1.Worksheets.Count
                    OrgSheet1 = excel1.Worksheets(i)
                    With OrgSheet1
                        index = 1
                        While (True)
                            Application.DoEvents()
                            If ESC Then
                                Exit While
                            End If
                            If myrow <= ProgressBar1.Maximum Then
                                ProgressBar1.Value = myrow
                            End If
                            mysn = Trim(.Cells(index, 4).value)
                            If Trim(mysn) = "" Then
                                Exit While
                            End If
                            ListBox1.Items.Add(mysn)
                            OrgSheet.Cells(myrow, 1) = .Cells(index, 4).value
                            OrgSheet.Cells(myrow, 2) = .Cells(index, 5).value
                            OrgSheet.Cells(myrow, 3) = .Cells(index, 8).value
                            index += 1
                            myrow += 1
                        End While
                    End With
                Next

                ProgressBar1.Visible = False
                ProgressBar1.Value = 0
                Dim s() As String
                Dim path As String
                Dim filename As String
                s = Split(OpenFileDialog1.FileName, "\")
                filename = s(UBound(s))
                ListBox1.Items.Add(" ")
                path = OpenFileDialog1.FileName.Replace(".xlsx", "处理结果.xlsx")
                excel.DisplayAlerts = False

                ListBox1.Items.Add("处理结果保存在： " + path)
                excel.Workbooks(1).SaveAs(path)
                excel.Workbooks(1).Close()

                excel1.Workbooks(1).Close()
                excel1.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1)
                excel1 = Nothing
                excel.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
                excel = Nothing
                GC.Collect()
                Windows.Forms.Cursor.Current = Cursors.Default

                ListBox1.Items.Add("资源列表处理完毕！")

            End If
        End If
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        On Error Resume Next
        With OpenFileDialog1
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With
        With OpenFileDialog2
            .Multiselect = False
            .FileName = ""
            .Filter = "EXCEL文件|*.xlsx;*.xls"
            .FilterIndex = 1
        End With
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If OpenFileDialog2.ShowDialog() = Windows.Forms.DialogResult.OK Then
                ESC = False
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                ListBox1.Items.Clear()

                ProgressBar1.Maximum = 6000
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                Dim excel As New Microsoft.Office.Interop.Excel.Application()
                Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

                Dim excel1 As New Microsoft.Office.Interop.Excel.Application()
                Dim OrgSheet1 As Microsoft.Office.Interop.Excel.Worksheet

                Dim myrow As Integer = 2
                Dim flag As Boolean
                Dim mysn As String
                Dim index As Integer
                Dim tt As String

                excel.Workbooks.Open(OpenFileDialog1.FileName)
                OrgSheet = excel.Worksheets(1)
                excel1.Workbooks.Open(OpenFileDialog2.FileName)
                OrgSheet1 = excel1.Worksheets(1)
                With OrgSheet
                    While (True)
                        Application.DoEvents()
                        If ESC Then
                            Exit While
                        End If
                        If myrow <= ProgressBar1.Maximum Then
                            ProgressBar1.Value = myrow
                        End If
                        mysn = Trim(.Cells(myrow, 1).value)
                        If Trim(mysn) = "" Then
                            Exit While
                        End If
                        ListBox1.Items.Add(mysn)
                        index = 2
                        With OrgSheet1
                            While (True)
                                tt = Trim(.Cells(index, 3).value)
                                If tt = mysn Then
                                    ListBox1.Items.Add("已找到！")
                                    OrgSheet.Cells(myrow, 2) = .Cells(index, 4).value
                                    OrgSheet.Cells(myrow, 3) = .Cells(index, 5).value
                                    OrgSheet.Cells(myrow, 4) = .Cells(index, 6).value
                                    OrgSheet.Cells(myrow, 5) = .Cells(index, 7).value
                                    OrgSheet.Cells(myrow, 6) = .Cells(index, 8).value
                                    OrgSheet.Cells(myrow, 7) = .Cells(index, 9).value
                                    OrgSheet.Cells(myrow, 8) = .Cells(index, 10).value
                                    OrgSheet.Cells(myrow, 9) = .Cells(index, 11).value
                                    OrgSheet.Cells(myrow, 10) = .Cells(index, 12).value
                                    OrgSheet.Cells(myrow, 11) = .Cells(index, 13).value
                                    OrgSheet.Cells(myrow, 12) = .Cells(index, 14).value
                                    OrgSheet.Cells(myrow, 13) = .Cells(index, 15).value
                                    Exit While
                                End If
                                If tt = "" And index > 10 Then
                                    Exit While
                                End If
                                index += 1
                            End While
                        End With

                        myrow = myrow + 1
                    End While
                End With

                ProgressBar1.Visible = False
                ProgressBar1.Value = 0
                Dim s() As String
                Dim path As String
                Dim filename As String
                s = Split(OpenFileDialog1.FileName, "\")
                filename = s(UBound(s))
                ListBox1.Items.Add(" ")
                path = OpenFileDialog1.FileName.Replace(".xlsx", "处理结果.xlsx")
                excel.DisplayAlerts = False

                ListBox1.Items.Add("处理结果保存在： " + path)
                excel.Workbooks(1).SaveAs(path)
                excel.Workbooks(1).Close()

                excel1.Workbooks(1).Close()
                excel1.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1)
                excel1 = Nothing
                excel.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
                excel = Nothing
                GC.Collect()
                Windows.Forms.Cursor.Current = Cursors.Default

                ListBox1.Items.Add("资源列表处理完毕！")

            End If
        End If
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        FolderBrowserDialog1.Description = "选择存放文件夹"
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("开始统计......")

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            str = "http://pv.tihe-china.com/manager/php/region.php?cmd=ListZM"

            Dim wc As New System.Net.WebClient
            Dim data As Byte() = wc.DownloadData(str)
            Dim content As String = System.Text.Encoding.UTF8.GetString(data)


            Dim yy() As String
            Dim zd() As String
            Dim zzz(28) As String
            Dim value() As String
            Dim city0, city1 As String
            Dim region As Integer = 0
            Dim regiontotal As Integer = 0
            Dim citycount As Integer = 0
            Dim citytotal As Integer = 0
            Dim sumrow As Integer = 0
            '  Dim provrow As Integer = 5
            Dim row As Integer = 1
            Dim row2 As Integer = 0
            Dim sumcity As Integer = 0
            Dim mydate As Date = Now()
            Dim yyid1, yyid2 As String
            Dim mcount(2000) As Integer
            Dim mmac(2000) As String
            Dim index As Integer = 0
            Dim mdates As Integer = 0
            Dim regionIDs As String = "," + Trim(TextBox3.Text) + ","

            mdates = DateDiff("d", DateTimePicker1.Value.Date, DateTimePicker2.Value.Date) + 1

            mySqlConc.ConnectionString = dbstr
            Try
                mySqlConc.Open()
            Catch ex As Exception
                MsgBox("数据库连接错误！")
                Return
            End Try
            myCommand.Connection = mySqlConc

            yyid1 = ""
            yyid2 = ""


            city0 = ""
            city1 = ""

            Dim excel As New Microsoft.Office.Interop.Excel.Application()
            Dim OrgSheet As Microsoft.Office.Interop.Excel.Worksheet

            excel.Workbooks.Add()
            excel.ActiveWindow.DisplayGridlines = False
            OrgSheet = excel.Worksheets(1)
            With OrgSheet
                .Name = "影院广告统计列表"
                .Columns("A:A").ColumnWidth = 5
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 5
                .Columns("D:D").ColumnWidth = 60
                .Columns("E:E").ColumnWidth = 16
                .Columns("A:A").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Columns("C:C").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Columns("E:E").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                .Rows("1:1").RowHeight = 30
                .Rows("2:2000").RowHeight = 20
                .Range("A1:L1").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Range("A1:E1").Font.Bold = True

                .Cells(1, 1) = "序号"
                .Cells(1, 2) = "城市"
                .Cells(1, 3) = "序号"
                .Cells(1, 4) = "影院广告统计" + "(" + DateTimePicker1.Value.ToShortDateString + "---" + DateTimePicker2.Value.ToShortDateString + ")"
                .Cells(1, 5) = "广告展现次数"

                yy = Split(content, "[{")
                content = yy(1).Replace("}]}", "")

                yy = Split(content, "},{")
                ProgressBar1.Maximum = UBound(yy)
                ProgressBar1.Minimum = 0
                ProgressBar1.Visible = True
                For i As Integer = 0 To UBound(yy)
                    ProgressBar1.Value = i
                    zd = Split(yy(i).Substring(0, yy(i).Length - 1), "','")
                    For j As Integer = 0 To UBound(zd)
                        value = Split(zd(j), "':'")
                        zzz(j) = value(1)
                    Next
                    If regionIDs <> ",," Then
                        If InStr(regionIDs, "," + zzz(1) + ",") = 0 Then
                            Continue For
                        End If
                    End If
                    index = 0
                    myCommand.CommandText = String.Format("SELECT count(ID) as ct,caption FROM zmcount WHERE ptime > '{0}' AND ptime < '{1}' AND mac='{2}' AND dsp='th' group by caption", DateTimePicker1.Value.ToShortDateString, DateTimePicker2.Value.AddDays(1).ToShortDateString, zzz(1))
                    myReader = myCommand.ExecuteReader()
                    While myReader.Read()
                        mcount(index) = myReader.GetInt64(0) - mdates
                        If mcount(index) < 0 Then
                            mcount(index) = 0
                        End If
                        mmac(index) = Trim(myReader.GetString(1))
                        index += 1
                    End While
                    myReader.Close()


                    city1 = zzz(0)

                    If city0 <> city1 Then
                        If city0 <> "" Then
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            '   .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                            .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                            .Cells((sumrow + 1).ToString, 2) = city0
                        End If
                        '   sumcity = sumcity + 1
                        '   row = row + 1
                        sumrow = row
                        regiontotal = regiontotal + region
                        region = 0
                    End If

                    yyid1 = zzz(1)
                    If yyid1 = yyid2 Then
                        ListBox1.Items.Add("重复ID：" + yyid1)
                    End If
                    row = row + 1
                    row2 = row2 + 1
                    region = region + 1
                    .Cells(row, 1) = row2
                    .Cells(row, 3) = region
                    .Cells(row, 4) = zzz(2)
                    .Cells(row, 5) = "广告展现次数"
                    .Range("C" + row.ToString + ":E" + row.ToString).Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                    .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                    .Range("C" + row.ToString + ":E" + row.ToString).Interior.ThemeColor = 5
                    .Range("C" + row.ToString + ":E" + row.ToString).Interior.TintAndShade = 0.599993896298105
                    .Range("C" + row.ToString + ":E" + row.ToString).Interior.PatternTintAndShade = 0
                    For ii As Integer = 0 To index - 1
                        row = row + 1
                        .Cells(row, 1) = "*"
                        .Cells(row, 3) = "*"
                        .Cells(row, 4) = mmac(ii)
                        .Cells(row, 5) = mcount(ii)
                    Next

                    yyid2 = yyid1
                    city0 = city1
                Next

                If city0 <> "" Then
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    ' .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Bold = True
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Font.Size = 10
                    .Range("B" + (sumrow + 1).ToString + ":B" + row.ToString).Merge()
                    .Cells((sumrow + 1).ToString, 2) = city0
                End If


                .Range("A1:L" + row.ToString).Font.Name = "微软雅黑"
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With
                With .Range("A2:E" + row.ToString).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                End With

            End With
            excel.Range("A2").Select()
            excel.ActiveWindow.FreezePanes = True
            ProgressBar1.Visible = False
            ProgressBar1.Value = 0
            mySqlConc.Close()
            ListBox1.Items.Add("正在保存泰和广告统计.......")
            excel.Worksheets(1).select()
            excel.DisplayAlerts = False
            mydate = DateTimePicker1.Value
            If Mid(FolderBrowserDialog1.SelectedPath, Len(FolderBrowserDialog1.SelectedPath), 1) = "\" Then
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "泰和广告统计" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "泰和广告统计" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            Else
                ListBox1.Items.Add(FolderBrowserDialog1.SelectedPath + "\泰和广告统计" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
                excel.Workbooks(1).SaveAs(FolderBrowserDialog1.SelectedPath + "\泰和广告统计" + mydate.ToShortDateString.Replace("/", "") + ".xlsx")
            End If

            excel.Workbooks(1).Close()
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            excel = Nothing
            GC.Collect()
            Windows.Forms.Cursor.Current = Cursors.Default
            ListBox1.Items.Add("泰和广告统计下载完毕！")
        End If
    End Sub
End Class
