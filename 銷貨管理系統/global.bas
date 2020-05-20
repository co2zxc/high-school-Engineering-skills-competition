Attribute VB_Name = "global"
Option Explicit
'Connection 的宣告
Public cn As ADODB.Connection

'*********Recordset的宣告**********
'訂貨單(訂單主檔/明細)
Public rsOrder As ADODB.Recordset
Public rsOrder_Detail As ADODB.Recordset
Public sql_rsOrder As String
Public sql_rsOrder_Detail As String
'銷貨單(銷貨主檔/明細)
Public rsSell As ADODB.Recordset
Public rsSell_Detail As ADODB.Recordset
Public sql_rsSell As String
Public sql_rsSell_Detail As String
'產品庫存統計表
Public rsProdtStore As ADODB.Recordset
Public sql_rsProdtStore As String
'產品盤存表與電腦盤點表
Public rsCheck As ADODB.Recordset
Public sql_rsCheck As String
Public rsComputer As ADODB.Recordset
Public sql_rsComputer As String
Public cmd As ADODB.Command
'員工資料表
Public rsStaff As ADODB.Recordset
Public sql_rsStaff As String
'員工聯絡人資料表
Public rsRelation As ADODB.Recordset
Public sql_rsRelation As String
'客戶資料
Public rsCust As ADODB.Recordset
Public sql_rsCust As String
'產品資料
Public rsProdt As ADODB.Recordset
Public sql_rsProdt As String
'地區表
Public rsArea As ADODB.Recordset
Public sql_rsArea As String
'生產核銷表
Public rsMake As ADODB.Recordset


'********常用的變數宣告*********
Public i As Integer             '用於迴圈的變數
Public oText As TextBox         '用於TextBox的變數
Public pre_prodt As String      '記錄前一個產品
Public bm_move As Variant       '用來記錄目前筆數intRecord
Public intAccount(1) As Integer '訂單明細的前後筆數(核銷用)
Public Add_flag As Integer      '新增鈕的狀態
Public Edit_flag As Integer     '修改鈕的狀態
Public cmdSell_flag As Integer  '銷貨單的來源
'rs共用的RecordSet之index值,部門分類表(1),職稱等級表(2),地區表(3)
Public Menu_index As Integer
Public intNowNum As Integer     '儲存庫存表的暫存量變數

'資料移動時,記錄其目前筆數與總筆數
Public intTotal As Integer
Public intRecord As Integer
'自動編號所使用之變數
Public Main_Head As String
Public Main_No As String
Public Sub_No As String
Public Main_Num As String
Public Main_Len As Integer

'資料移動之副程式
Public Sub rs_Move(inx As Integer, rs As ADODB.Recordset)
    '訂單總筆數
    intTotal = rs.RecordCount
    Select Case inx
        Case 0  '第一筆
            rs.MoveFirst
            intRecord = 1
            
        Case 1  '前一筆
            rs.MovePrevious
            If rs.BOF Then
                rs.MoveFirst
            End If
            intRecord = intRecord - 1
            If intRecord < 1 Then intRecord = 1
            
        Case 2  '下一筆
            rs.MoveNext
            If rs.EOF Then
                rs.MoveLast
            End If
            intRecord = intRecord + 1
            If intRecord > intTotal Then intRecord = intTotal
            
        Case 3  '最末筆
            rs.MoveLast
            intRecord = intTotal
    End Select
    
End Sub

'主檔編號的自動新增
Public Sub Add_No(rs As ADODB.Recordset, indx As Integer, num_index As Integer)
    
    Main_Head = ""
    Select Case indx
        Case 1  '客戶
            Main_Head = "C"
        Case 2  '員工
            Main_Head = "E"
        Case 3  '訂單
            Main_Head = "O"
        Case 4  '銷貨單
            Main_Head = "S"
        Case 5  '產品
            Main_Head = "P"
        Case 6  '地區表
            Main_Head = "A"
    End Select
    
    If rs.RecordCount <> 0 Then
        Select Case indx
            Case 1  '客戶
                rs.Sort = "客戶編號 asc"
            Case 2  '員工
               rs.Sort = "員工編號 asc"
            Case 3  '訂單
                rs.Sort = "訂單編號 asc"
            Case 4  '銷貨單
                rs.Sort = "銷貨單編號 asc"
            Case 5  '產品
                rs.Sort = "產品編號 asc"
            Case 6  '地區表
                rs.Sort = "地區編號 asc"
        End Select
        
        rs.MoveLast
        Select Case num_index
            Case 1      'Main_Num為三位數
                Main_Num = Trim(Int(Right(rs.Fields(0), 3) + 1))
                Main_Len = Len(Main_Num)
                
                Select Case Main_Len
                    Case 1
                        Main_No = Main_Head & "00" & Main_Num
                    Case 2
                        Main_No = Main_Head & "0" & Main_Num
                    Case 3
                        Main_No = Main_Head & Main_Num
                End Select
            Case 2      'Main_Num為二位數
                Main_Num = Trim(Str(Right(rs.Fields(0), 2) + 1))
                Main_Len = Len(Main_Num)
                
                Select Case Main_Len
                    Case 1
                        Main_No = Main_Head & "0" & Main_Num
                    Case 2
                        Main_No = Main_Head & Main_Num
                End Select
        End Select
    Else
        Select Case num_index
            Case 1
                Main_No = Main_Head & "001"
            Case 2
                Main_No = Main_Head & "01"
        End Select
    End If
    rs.AddNew
    
End Sub

'明細編號的自動新增
Public Sub Add_SubNo(rsD As ADODB.Recordset, indx As Integer, txtMain_No As String)
    Dim Sub_Head As String
    Dim Sub_ASC As String
    Dim Sub_Num As Integer
    
    Sub_Head = txtMain_No
    Set rsD = New ADODB.Recordset
    Select Case indx
        Case 1
            rsD.Open "select * from 訂單明細 where 訂單編號= '" & Sub_Head & "'", cn, adOpenDynamic, adLockOptimistic
        Case 2
            rsD.Open "select * from 銷貨明細 where 銷貨單編號='" & Sub_Head & "'", cn, adOpenDynamic, adLockOptimistic
    End Select
    
    '若明細有值
    If rsD.RecordCount <> 0 Then
        Select Case indx
            Case 1  '訂單
                rsD.Sort = "訂單明細編號 asc"
                
            Case 2  '銷貨單
                rsD.Sort = "銷貨單明細編號 asc"
        End Select
        rsD.MoveLast
        Sub_Num = Asc(Right(rsD.Fields(1), 1)) + 1
        Sub_ASC = Chr(Sub_Num)
        Sub_No = Sub_Head & Sub_ASC
    Else
        Sub_No = Sub_Head & "a"
    End If
End Sub


