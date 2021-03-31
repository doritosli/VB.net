Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Net.Dns
Imports System.Windows.Forms

Public Class 上傳考勤
    Dim DataAdapterDGV As SqlDataAdapter
    Dim ks1DataSetDGV As DataSet = New DataSet

    Public DBConn1 As SqlConnection
    Public DBConn2 As SqlConnection

    Dim ErrorAmount As Integer = 0
    Dim ErrorString As String = ""

    Private Sub 上傳考勤_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        時間.Text = "登入電腦：" & UCase(GetHostName()) & "      登入時間：" & Format(FileDateTime(Application.StartupPath & "\" & "上傳考勤.exe"), "yyyy-MM-dd hh:mm")
        '公司.SelectedIndex = 0  '預設為KS
        測試連接主機()
    End Sub

    Public Sub 測試連接主機()
        測試連接主機1()  '測試區
    End Sub

    Public Sub 測試連接主機1()
        If My.Computer.Network.Ping(DBServer1) Then  '測試連接資料庫
            'MsgBox("資料庫1連線成功")
            ConnDB1()
            If ErrorAmount = 1 Then
                MsgBox("讀取SERVER 1 資料庫失敗：" & vbCrLf & ErrorString, 16, "錯誤")
                Exit Sub
            End If
        Else
            MsgBox("連接 SERVER 1 失敗，請連絡資訊人員協助!!", 16)
            Exit Sub
        End If
        DBConn1.Close()
    End Sub

    Public Sub ConnDB1()
        ErrorAmount = 0 : ErrorString = ""
        Try
            DBConn1 = New SqlConnection("Data Source = " & DBServer1 & "; Initial Catalog = " & DBName1 & "; User ID = " & DBUser1 & "; Password = " & DBPass1 & "; Connection Timeout = 5")
            DBConn1.Open() '開啟連線
            'DBConn1.Close() '結束連線

        Catch ex As Exception
            'MsgBox("讀取SERVER資料庫失敗：" & vbCrLf & ex.Message, 16, "錯誤")
            ErrorString = ex.Message
            ErrorAmount = 1
            Exit Sub
        End Try
    End Sub

    Private Sub Bu讀取上傳_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Bu讀取上傳.Click
        上傳作業()
    End Sub

    Private Sub 上傳作業() '上傳xls檔
        Dim openfile As New OpenFileDialog()

        openfile.InitialDirectory = "%HOMEPATH%\Desktop\"
        openfile.Filter = "EXCEL files(*.xls)|*.xls" '只能抓excel檔   
        openfile.ShowDialog()

        If openfile.FileName = "" Then : Exit Sub : End If

        Dim connectionString As String = " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & openfile.FileName & ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1' "
        Dim strSQL As String = ""

        strSQL = " SELECT * FROM [Sheet1$] "

        Dim excelConnection As New Data.OleDb.OleDbConnection(connectionString)
        Try
            excelConnection.Open()

            Dim dbCommand As OleDbCommand = New OleDbCommand(strSQL, excelConnection)
            Dim dataAdapter As OleDbDataAdapter = New OleDbDataAdapter(dbCommand)
            Dim dTable As DataTable = New DataTable()
            dataAdapter.Fill(dTable)
            DGV1.DataSource = dTable
            dTable.Dispose()
            dataAdapter.Dispose()
            dbCommand.Dispose()
            excelConnection.Close() : excelConnection.Dispose()

            DGV1欄位() '顯示匯入的資料
            資料回寫() '寫進存_打卡暫存
            資料讀取() '拿存_打卡暫存來串資料
            'DGV2欄位() '顯示串好的資料表
        Catch ex As Exception
            MsgBox("讀取Excel 錯誤：" & ex.Message) : Exit Sub
        End Try
    End Sub

    Private Sub DGV1欄位() '匯入的xls檔內容

        For Each column As DataGridViewColumn In DGV1.Columns
            column.Visible = True
            Select Case column.Name
                Case "F1" : column.HeaderText = "員工編號"
                Case "F2" : column.HeaderText = "員工姓名"
                Case "F3" : column.HeaderText = "地點代碼"
                Case "F4" : column.HeaderText = "刷卡日期"
                Case "F5" : column.HeaderText = "刷卡時間"
                Case "F6" : column.HeaderText = "刷卡代碼"
                Case Else
                    column.Visible = False
            End Select

        Next
        DGV1.AutoResizeColumns()
    End Sub

    Private Sub 資料回寫() '寫回 存_打卡暫存 資料表

        Dim SQLADD As String = "" : Dim SQLDEL As String = ""
        Dim KID As Integer = 0 : Dim SID As Integer = 0
        Dim DocDate As String = "" : Dim DocNo As String = "" : Dim Del As String = "Y" : Dim j = 0

        If DGV1.RowCount > 0 Then
            For i = 0 To DGV1.RowCount - 1

                SQLADD += " INSERT INTO [Payroll].[dbo].[存_打卡暫存] ([員工編號],[員工姓名],[地點代碼],[刷卡日期],"
                SQLADD += " [刷卡時間], [刷卡代碼],[公司代碼],[啟用否],[轉檔否])  "
                SQLADD += " VALUES ('" & DGV1.Rows(i).Cells("F1").Value & "','" & DGV1.Rows(i).Cells("F2").Value & "',"
                SQLADD += "'" & DGV1.Rows(i).Cells("F3").Value & "','" & DGV1.Rows(i).Cells("F4").Value & "',"
                SQLADD += "'" & Format(DGV1.Rows(i).Cells("F5").Value, "HH:mm:ss") & "','" & DGV1.Rows(i).Cells("F6").Value & "','""','Y','N') "

            Next
        End If
        'MsgBox(SQLDEL)
        Try
            Dim DBConn As SqlConnection = DBConn1
            Dim SQLCmd As SqlCommand = New SqlCommand
            SQLCmd.Connection = DBConn : SQLCmd.CommandText = SQLDEL + SQLADD : SQLCmd.CommandTimeout = 100000
            DBConn.Open()
            SQLCmd.Parameters.Clear() : SQLCmd.ExecuteNonQuery()
            DBConn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        MsgBox("上傳成功！")
    End Sub

    Private Sub 資料讀取() '讀取資料show在DGV2

        Dim sql As String
        Dim DBConn As SqlConnection = DBConn1
        Dim SQLCmd As SqlCommand = New SqlCommand

        '用新員編HQ...去 "新舊員編對照" 抓舊員編
        sql = " Select distinct "
        sql += " case when T1.新員工編號 LIKE 'HQ%' then 'KS' "
        sql += " when T1.新員工編號 LIKE 'HM%' then 'HM' "
        sql += " when T1.新員工編號 LIKE 'QT%' then 'QTC' "
        sql += " end as '公司代碼', "
        sql += " T1.[舊員工編號] AS '員工編號', "
        sql += " T0.[員工姓名], "
        sql += " MAX(convert(varchar(100),cast(T0.[刷卡日期] as DATE), 23) + ' ' + SUBSTRING(convert(varchar(50), T0.[刷卡時間], 108),0,9)) as '刷卡時間', "
        sql += " T0.[刷卡代碼], T0.[啟用否], T0.[轉檔否], T0.[地點代碼], "
        sql += " cast(T0.[刷卡日期] as date) as '歸屬日期', "
        sql += " case when T0.[刷卡代碼] = '1' then '上班' "
        sql += " when T0.[刷卡代碼] = '2' then '下班' "
        sql += " when T0.[刷卡代碼] = '3' then '外出' "
        sql += " when T0.[刷卡代碼] = '4' then '外出' "
        sql += " when T0.[刷卡代碼] = '5' then '加班上' "
        sql += " when T0.[刷卡代碼] = '6' then '加班下' "
        sql += " when T0.[刷卡代碼] = '7' then '公差' "
        sql += " when T0.[刷卡代碼] = '8' then '公差' "
        sql += " end as [刷卡作業], "
        sql += " case when T0.[地點代碼] = '1' then '大門' when T0.[地點代碼] = '2' then '守衛室' when T0.[地點代碼] = '3' then '監控室' end as [地點名稱], "
        sql += " getdate() as [系統時間] "
        sql += " from [Payroll].[dbo].[打卡暫存] T0 "
        sql += " left join [Payroll].[dbo].[新舊員編對照] T1 ON T0.員工編號 = T1.新員工編號 "
        sql += " group by 公司代碼,員工編號,T0.員工姓名,刷卡時間,刷卡代碼,啟用否,轉檔否,地點代碼,刷卡日期,T1.舊員工編號,T1.新員工編號 "


        '取新員工編號HQ開頭
        'sql = "select distinct "
        'sql += "case when T0.員工編號 LIKE 'HQ%' then 'KS'"
        'sql += "when T0.員工編號 LIKE 'HM%' then 'HM'"
        'sql += "when T0.員工編號 LIKE 'QT%' then 'QTC' "
        'sql += "end as '公司代碼',T0.[員工編號], T0.[員工姓名],"
        'sql += "MAX(convert(varchar(100),cast(T0.[刷卡日期] as DATE), 23) + ' ' + SUBSTRING(convert(varchar(50), T0.[刷卡時間], 108),0,9)) as '刷卡時間',"
        'sql += "T0.[刷卡代碼], T0.[啟用否], T0.[轉檔否], T0.[地點代碼],"
        'sql += "cast(T0.[刷卡日期] as date) as '歸屬日期',"
        'sql += "case when T0.[刷卡代碼] = '1' then '上班' "
        'sql += "when T0.[刷卡代碼] = '2' then '下班'"
        'sql += "when T0.[刷卡代碼] = '3' then '外出'"   '午休出
        'sql += "when T0.[刷卡代碼] = '4' then '外出'"   '午休回
        'sql += "when T0.[刷卡代碼] = '5' then '加班上'"
        'sql += "when T0.[刷卡代碼] = '6' then '加班下'"
        'sql += "when T0.[刷卡代碼] = '7' then '公差'"   '洽公出
        'sql += "when T0.[刷卡代碼] = '8' then '公差'"   '洽公回
        'sql += "end as [刷卡作業],"
        'sql += "case when T0.[地點代碼] = '1' then '大門' when T0.[地點代碼] = '2' then '守衛室' when T0.[地點代碼] = '3' then '監控室' end as [地點名稱],"
        'sql += "getdate() as [系統時間] "
        'sql += "from [Payroll].[dbo].[存_打卡暫存] T0 "
        'sql += "group by 公司代碼,員工編號,員工姓名,刷卡時間,刷卡代碼,啟用否,轉檔否,地點代碼,刷卡日期"


        SQLCmd.Connection = DBConn
        SQLCmd.CommandText = sql

        DataAdapterDGV = New SqlDataAdapter(sql, DBConn)
        DataAdapterDGV.Fill(ks1DataSetDGV, "DT2")
        DGV2.DataSource = ks1DataSetDGV.Tables("DT2")


        DGV2欄位()

    End Sub

    Private Sub DGV2欄位()

        For Each column As DataGridViewColumn In DGV2.Columns
            column.Visible = True

            Select Case column.Name
                Case "員工編號" : column.HeaderText = "員工編號" : column.DisplayIndex = 0
                Case "地點代碼" : column.HeaderText = "地點代碼" : column.DisplayIndex = 1
                Case "刷卡時間" : column.HeaderText = "刷卡時間" : column.DisplayIndex = 2
                Case "刷卡作業" : column.HeaderText = "刷卡作業" : column.DisplayIndex = 3
                Case "公司代碼" : column.HeaderText = "公司代碼" : column.DisplayIndex = 4
                Case "轉檔否" : column.HeaderText = "轉檔否" : column.DisplayIndex = 5
                Case "刷卡代碼" : column.HeaderText = "刷卡代碼" : column.DisplayIndex = 6
                Case "員工姓名" : column.HeaderText = "員工姓名" : column.DisplayIndex = 7
                Case "地點名稱" : column.HeaderText = "地點名稱" : column.DisplayIndex = 8
                Case "系統時間" : column.HeaderText = "系統時間" : column.DisplayIndex = 9
                Case "啟用否" : column.HeaderText = "啟用否" : column.DisplayIndex = 10
                Case "歸屬日期" : column.HeaderText = "歸屬日期" : column.DisplayIndex = 11
                Case Else
                    column.Visible = False
            End Select
        Next
        DGV2.AutoResizeColumns()

        La1筆數.Text = "筆數：" & DGV2.RowCount

    End Sub

    Private Sub 最後資料回寫() '最後插入"基_刷卡記錄"資料表

        Dim SQLADD As String = "" : Dim SQLDEL As String = ""
        Dim KID As Integer = 0 : Dim SID As Integer = 0
        Dim DocDate As String = "" : Dim DocNo As String = "" : Dim Del As String = "Y" : Dim j = 0

        If DGV1.RowCount > 0 Then
            For i = 0 To DGV1.RowCount - 1

                SQLADD += " INSERT INTO [Payroll].[dbo].[刷卡記錄] ([公司代碼], [刷卡代碼], [刷卡作業],"
                SQLADD += " [啟用否],[轉檔否],[員工編號],[員工姓名],[地點代碼],[地點名稱],[狀態代碼],[刷卡時間],"
                SQLADD += " [系統時間],[歸屬日期],[執行電腦]) "
                SQLADD += " VALUES ( "
                SQLADD += "'" & DGV2.Rows(i).Cells("公司代碼").Value & "',"
                SQLADD += "'" & DGV2.Rows(i).Cells("刷卡代碼").Value & "',"
                SQLADD += "'" & DGV2.Rows(i).Cells("刷卡作業").Value & "',"
                SQLADD += "'Y','N','" & DGV2.Rows(i).Cells("員工編號").Value & "',"
                SQLADD += "'" & DGV2.Rows(i).Cells("員工姓名").Value & "',"
                SQLADD += "'" & DGV2.Rows(i).Cells("地點代碼").Value & "',"
                SQLADD += "'" & DGV2.Rows(i).Cells("地點名稱").Value & "','1',"
                SQLADD += "'" & DGV2.Rows(i).Cells("刷卡時間").Value & "',"
                'SQLADD += "'" & Format(DGV2.Rows(i).Cells("刷卡時間").Value, "yyyy-MM-dd HH:mm:ss") & "',"
                SQLADD += "GETDATE(),'" & DGV2.Rows(i).Cells("歸屬日期").Value & "',"
                SQLADD += "'" & UCase(GetHostName()) & "')"
            Next
        End If
        'MsgBox(SQLDEL)
        Try
            Dim DBConn As SqlConnection = DBConn1
            Dim SQLCmd As SqlCommand = New SqlCommand
            SQLCmd.Connection = DBConn : SQLCmd.CommandText = SQLDEL + SQLADD : SQLCmd.CommandTimeout = 100000
            DBConn.Open() '開啟資料庫連線
            SQLCmd.Parameters.Clear() : SQLCmd.ExecuteNonQuery() '讀取資料庫
            DBConn.Close() '關閉資料庫連線
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        MsgBox("回存成功!")

    End Sub

    Private Sub 清空資料庫_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 清空資料庫.Click
        If DGV1.RowCount > 0 Then : DGV1.DataSource.Clear() : End If '清除DGV1資料
        If DGV2.RowCount > 0 Then : ks1DataSetDGV.Tables("DT2").Clear() : End If '清除DGV2資料

        Dim sql As String
        Dim DBConn As SqlConnection = DBConn1
        Dim SQLCmd As SqlCommand = New SqlCommand

        sql = " TRUNCATE TABLE [Payroll].[dbo].[打卡暫存] "
        'sql += " TRUNCATE TABLE [Payroll].[dbo].[刷卡記錄] "

        SQLCmd.Connection = DBConn
        SQLCmd.CommandText = sql

        DataAdapterDGV = New SqlDataAdapter(sql, DBConn)
        DataAdapterDGV.Fill(ks1DataSetDGV, "DT2")
        DGV2.DataSource = ks1DataSetDGV.Tables("DT2")


        DGV2欄位()
        MsgBox("已清空")

    End Sub

    Private Sub 資料回填_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 資料回填.Click
        If DGV1.RowCount > 0 Then
            最後資料回寫()
            Label1.Visible = True
        Else
            MsgBox("請先上傳資料！", 48, "錯誤訊息")
        End If



    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    'Private Sub 改公司_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 改公司.Click

    '    '公司代碼QT改成HM

    '    Dim i As Integer
    '    For i = 0 To DGV2.Rows.Count - 1
    '        If DGV2.Rows(i).Cells("公司代碼").Value = "QTC" Then
    '            DGV2.Rows(i).Cells("公司代碼").Value = "HM"
    '        End If
    '    Next
    'End Sub
End Class

