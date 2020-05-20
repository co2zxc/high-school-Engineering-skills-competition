VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "銷貨管理系統  主畫面"
   ClientHeight    =   4350
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6630
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.Menu mnuInfo 
      Caption         =   "基本資料"
      Begin VB.Menu mnuStaffKind 
         Caption         =   "員工資料管理"
         Begin VB.Menu mnuStaff 
            Caption         =   "員工資料表"
         End
      End
      Begin VB.Menu mnuline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBusiness 
         Caption         =   "客戶資料管理"
         Begin VB.Menu mnuCust 
            Caption         =   "客戶資料表"
         End
      End
      Begin VB.Menu mnuline4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAreaind 
         Caption         =   "地區資料管理"
         Begin VB.Menu mnuArea 
            Caption         =   "地區表"
         End
      End
      Begin VB.Menu mnuline5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObject 
         Caption         =   "產品資料管理"
         Begin VB.Menu mnuProdt 
            Caption         =   "產品資料表"
         End
         Begin VB.Menu mnuline12 
            Caption         =   "-"
         End
         Begin VB.Menu mnub 
            Caption         =   "廠商資料"
         End
         Begin VB.Menu mnuC 
            Caption         =   "類別資料"
         End
      End
      Begin VB.Menu mnuline13 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuSheet 
      Caption         =   "建立單據"
      Begin VB.Menu mnuProdtKind 
         Caption         =   "產品類別"
         Begin VB.Menu mnuOrder 
            Caption         =   "訂貨單"
         End
         Begin VB.Menu mnuline6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSell 
            Caption         =   "銷貨單查詢"
         End
      End
   End
   Begin VB.Menu mnuMkCk 
      Caption         =   "生產核銷"
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "需求查詢"
      Begin VB.Menu mnuProdtStore 
         Caption         =   "產品庫存統計表"
      End
      Begin VB.Menu mnuline7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSafeNum 
         Caption         =   "安全存量的檢查"
      End
      Begin VB.Menu mnuline8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQ1 
         Caption         =   "商品查詢(下拉式)"
      End
      Begin VB.Menu mnuQ2 
         Caption         =   "商品查詢(關鍵字)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "建立報表"
      Begin VB.Menu mnuRptComputer 
         Caption         =   "製作電腦盤存報表"
      End
      Begin VB.Menu mnuline9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "編輯產品盤存表"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "工具"
      Begin VB.Menu mnuQ3 
         Caption         =   "計算機"
      End
      Begin VB.Menu mnuEasy 
         Caption         =   "便條"
      End
      Begin VB.Menu mnuQ4 
         Caption         =   "MP3播放機"
      End
      Begin VB.Menu mnuline10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "資料的備份"
      End
      Begin VB.Menu mnuline14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInitial 
         Caption         =   "建立下期期初環境"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "結束"
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mnuSafeNum_flag As Integer   '表必須生產要補足安全存量之旗標
Dim DBName As String                '回傳所輸入的資料庫檔案路徑之字串
Dim sql_rsCkNo As String
'盤存資料的變數
Dim ck_Head As String
Dim ck_Num As String
Dim ck_Len As Integer
Dim ck_No As String

'資料庫的連結
Private Sub MDIForm_Load()
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    ChDir (App.Path)
    cn.Open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbRose.mdb"
    
    '產品資料常被使用,故在表單載入時即先開啟
    Set rsProdt = New ADODB.Recordset
    rsProdt.Open "select * from 產品資料", cn, adOpenDynamic, adLockOptimistic
        
    '月底盤存功能的狀態
    Call check_status
    '銷貨單查詢功能的狀態
    Call mnuSell_status
End Sub

Private Sub mnub_Click()
Unload Me: frmOrd_Detail2.Show
End Sub

'製作資料的備份
Private Sub mnuBackup_Click()
    '設定備份資料庫之檔案路徑
    DBName = InputBox("請輸入備份資料庫之檔案路徑!!" & "如C:\VB6\DataBase\Backup.mdb", "金光益電器股份有限公司")
    If DBName = "" Then
        Exit Sub
    Else
        If Right(DBName, 4) <> ".mdb" Then
        MsgBox "請為資料庫名稱加上副檔名'.mdb'"
        Call mnuBackup_Click
    End If

    End If
    
    '製作備份資料庫
    'Backup_Database=1時,表放棄覆寫之前所輸入的檔名,故須再輸入新的檔名
    If Backup_Database(DBName) = 1 Then
        Call mnuBackup_Click
    Else
        MsgBox ("資料庫已備份完成!!")
    End If
End Sub


Private Sub mnuC_Click()
Unload Me
frmOrd_Detail1.Show
End Sub

'****部門分類表 職稱等級表與地區表共用一個表單,所以必須有一個Index的值,
'來傳入不同的RecordSet****
'顯示部門分類表
Private Sub mnuDepart_Click()
    Menu_index = 1
    mnuArea.Enabled = False
    mnuDuty.Enabled = False
    frmInfo.Show
    mnuExit.Enabled = False
End Sub

'顯示職稱等級表
Private Sub mnuDuty_Click()
    Menu_index = 2
    mnuArea.Enabled = False
    mnuDepart.Enabled = False
    
    frmInfo.Show
    mnuExit.Enabled = False
End Sub

'顯示地區表
Private Sub mnuArea_Click()
    Menu_index = 3
    
    frmInfo.Show
    mnuExit.Enabled = False
End Sub

'顯示客戶資料表
Private Sub mnuCust_Click()
    mnuExit.Enabled = False
    frmCust.Show
End Sub

Private Sub mnuEasy_Click()
 Unload Me
frmDrawLine.Show
End Sub

'顯示編輯產品盤存表
Private Sub mnuEditCheck_Click()
    frmCheck.Show
    mnuExit.Enabled = False
End Sub

'離開此系統
Private Sub mnuExit_Click()
    cn.Close
    Set cn = Nothing
    End
End Sub

'建立下期期初環境
Private Sub mnuInitial_Click()
   ' Dim old_DBName As String
        
   ' If Add_flag = 1 Then
   '     DBName = old_DBName
   ' End If
   ' Add_flag = 0
        
    '1.設定資料庫的檔案路徑
    DBName = InputBox("請輸入前期歷史檔的備份檔案路徑!" & "如C:\VB6\DataBase\History.mdb", "金光益電器股份有限公司")

    If DBName = "" Then
        Exit Sub
    Else
       ' Add_flag = 1
       ' old_DBName = DBName
        If Right(DBName, 4) <> ".mdb" Then
            MsgBox "請為資料庫名稱加上副檔名'.mdb'"
            Call mnuInitial_Click
        End If
    End If
    
    '2.製作歷史檔的備份
    'Backup_Database=1時,表放棄覆寫之前所輸入的檔名,故須再輸入新的檔名
    If Backup_Database(DBName) = 1 Then
        Call mnuInitial_Click
    Else
        '備份完單據資料後
        '(1).清除訂貨單,銷貨單與產品庫存統計表的資料的資料
        sql_rsOrder = "delete * from 訂單主檔"
        sql_rsOrder_Detail = "delete * from 訂單明細"
        sql_rsSell = "delete * from 銷貨主檔"
        sql_rsSell_Detail = "delete * from 銷貨明細"
        sql_rsProdtStore = "delete * from 產品庫存統計表"
        
        '執行動作查詢
        cn.Execute sql_rsOrder
        cn.Execute sql_rsOrder_Detail
        cn.Execute sql_rsSell
        cn.Execute sql_rsSell_Detail
        cn.Execute sql_rsProdtStore
        
        '(2).將電腦盤點表的產品編號與實際盤存量的值新增進產品庫存統計表中
        sql_rsComputer = "insert into 產品庫存統計表(產品編號,產品名稱,上期庫存量) select 產品編號,產品名稱,實際盤存量 from 電腦盤點表"
        cn.Execute sql_rsComputer
        
        '(3).再將安全存量放進產品庫存統計表中
        Set rsProdtStore = New ADODB.Recordset
        rsProdtStore.Open "select * from 產品庫存統計表", cn, adOpenDynamic, adLockOptimistic

        Do Until rsProdtStore.EOF
            rsProdt.MoveFirst
            rsProdt.Find "產品編號='" & rsProdtStore.Fields("產品編號") & "'", , 1
            rsProdtStore.Fields("安全存量") = rsProdt.Fields("安全存量")
            rsProdtStore.Fields("暫存量") = rsProdtStore.Fields("現存量")
            
            rsProdtStore.MoveNext
        Loop
        
        '(4).將盤點編號記錄於'本月盤點編號'資料表中
        sql_rsCkNo = "insert into 本月盤點編號(盤點編號) select 盤點編號 from 產品盤存表 "
        cn.Execute sql_rsCkNo
        
        '(5).接著即移除產品盤存表與電腦盤點表的資料
        sql_rsCheck = "delete * from 產品盤存表"
        sql_rsComputer = "delete * from 電腦盤點表"
        cn.Execute sql_rsCheck
        cn.Execute sql_rsComputer
        
        '提醒使用者,期初環境已建立完成
        MsgBox ("下期空白單據已建構完成!!")
        '重整主表單中的功能項目
        Call check_status
    End If
End Sub

'顯示生產核銷表
Private Sub mnuMkCk_Click()
    frmMkCk.Show 1
End Sub

'顯示訂貨單
Private Sub mnuOrder_Click()
    frmOrder.Show
    mnuExit.Enabled = False
End Sub

'顯示產品資料表
Private Sub mnuProdt_Click()
    frmProdt.Show
    mnuExit.Enabled = False
End Sub

'顯示產品庫存統計表
Private Sub mnuProdtStore_Click()
    frmProdtStore.Show
    mnuExit.Enabled = False
End Sub

Private Sub mnuQ1_Click()
 Unload Me
    frmBuy_Detail.Show
End Sub

Private Sub mnuQ2_Click()
 Unload Me
    frmBuy1_Detail.Show
End Sub

Private Sub mnuQ3_Click()
 Unload Me
    number.Show
End Sub

Private Sub mnuQ4_Click()
 Unload Me
frmJukebox.Show
End Sub

'建立與顯示電腦盤存報表
Private Sub mnuRptComputer_Click()
    '清除產品盤存表所使用的動作查詢指令
    Dim sql_CheckClear As String
    '代替產品盤存表中的盤點編號欄位之變數
    Dim Check_No As String

    '1.每次點選"建立電腦盤存報表"項目時,都會先將資料清除掉,再重新Insert一次
    sql_CheckClear = "delete from 產品盤存表"
    '執行動作查詢
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = sql_CheckClear
    cmd.Execute
    
    '為原來空的產品盤存表,新增資料記錄
    Call CheckNo_Add
        
    '2.開啟電腦盤點表
    Set rsComputer = New ADODB.Recordset
    sql_rsComputer = "select * from 電腦盤點表"
    rsComputer.Open sql_rsComputer, cn, adOpenDynamic, adLockOptimistic

    '3.製作做查詢指令,將產品庫存統計表的產品編號與現存量的值Insert至電腦盤點表中
    sql_rsProdtStore = "insert into 電腦盤點表(產品編號,產品名稱,應有庫存量) select 產品編號,產品名稱,現存量 from 產品庫存統計表"
    '執行動作查詢
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = sql_rsProdtStore
    cmd.Execute
    '將值填入電腦盤點表中
    rsComputer.Requery
    
    Do Until rsComputer.EOF
        Check_No = rsCheck.Fields("盤點編號")
        '將盤點編號Insert進來
        rsComputer.Fields("盤點編號") = Check_No
        rsComputer.Update
        rsComputer.MoveNext
    Loop
   
    '4.顯示出電腦盤存報表
    RptComputer.Show
End Sub

'新增產品盤存表的盤點編號與盤點日期欄位值
Private Sub CheckNo_Add()
    Set rsCheck = New ADODB.Recordset
    sql_rsCheck = "select * from 產品盤存表"
    rsCheck.Open sql_rsCheck, cn, adOpenDynamic, adLockOptimistic

    ck_Head = Year(Date)
    ck_Num = Month(Date)
    ck_Len = Len(ck_Num)
    
    Select Case ck_Len
        Case 1
            ck_No = "C" & ck_Head & "0" & ck_Num
        Case 2
            ck_No = "C" & ck_Head & ck_Num
    End Select
    rsCheck.AddNew
    rsCheck.Fields("盤點編號") = ck_No
    rsCheck.Fields("盤點日期") = Date
    rsCheck.Update
End Sub

'檢查安全存量是否足夠,決定是否要生產
Private Sub mnuSafeNum_Click()
    Dim self_order_flag As Integer
    '應生產的數量
    Dim intSelf_OrNum As Integer
        
    Set rsOrder = New ADODB.Recordset
    rsOrder.Open "select * from 訂單主檔", cn, adOpenDynamic, adLockOptimistic
    
    Set rsOrder_Detail = New ADODB.Recordset
    rsOrder_Detail.Open "select * from 訂單明細", cn, adOpenDynamic, adLockOptimistic
    
    Set rsProdtStore = New ADODB.Recordset
    rsProdtStore.Open "select * from 產品庫存統計表", cn, adOpenDynamic, adLockOptimistic

    
    '1.檢查每一個產品的現存量是否低於安全存量
    Do Until rsProdtStore.EOF
        '若現存量為負數,即表示庫存量已為0
        If Sgn(rsProdtStore.Fields("暫存量")) = -1 Then
            intNowNum = 0
        Else
            intNowNum = rsProdtStore.Fields("暫存量")
        End If
        '將應生產的數量以變數intSelf_OrNum記錄,以方便使用
        intSelf_OrNum = intNowNum - rsProdtStore.Fields("安全存量")
        '2.當現存量低於安全存量時,即要進行生產,但生產必須依據訂貨單,故此處先產生一張 _
           客戶為公司本身的訂貨單,然後再依據銷貨流程來進行生產
        If Sgn(intSelf_OrNum) = -1 Then
            '訂單主檔的新增
            If self_order_flag = 0 Then
                Call Add_No(rsOrder, 3, 1)
                rsOrder.Fields(0) = Main_No
                rsOrder.Fields(1) = "E017"
                '即為金光益電器股份有限公司本身
                rsOrder.Fields(2) = "C000"
                rsOrder.Fields(3) = Date
                rsOrder.Fields(4) = 0
                rsOrder.Update
                '新增訂單主檔的旗標值
                self_order_flag = 1
            End If
            
            '訂單明細的新增
            Call Add_SubNo(rsOrder_Detail, 1, Main_No)
            rsOrder_Detail.AddNew
            rsOrder_Detail.Fields(0) = Main_No
            rsOrder_Detail.Fields(1) = Sub_No
            rsOrder_Detail.Fields(2) = rsProdtStore.Fields(0)
            rsOrder_Detail.Fields(3) = rsProdtStore.Fields(1)
            rsProdt.MoveFirst
            rsProdt.Find "產品編號='" & rsProdtStore.Fields(0) & "'", , 1
            rsOrder_Detail.Fields(4) = rsProdt.Fields("單位成本")
            rsOrder_Detail.Fields(5) = Abs(intSelf_OrNum)
            rsOrder_Detail.Fields(6) = rsProdt.Fields("單位成本") * Abs(intSelf_OrNum)
            rsOrder_Detail.Fields(7) = 0
            rsOrder_Detail.Fields(8) = 0
            rsOrder_Detail.Update
        End If
        
        rsProdtStore.MoveNext
    Loop
    
    '3.若有需要生產的數量,以補足安全存量,則以frmMake表單呈現出明細來
    If self_order_flag = 1 Then
        '由於frmMake被'訂貨單'與'安全存量的檢查'所共用,此旗標為一個區別值
        mnuSafeNum_flag = 1
        frmMake.Caption = "生產單-補足安全存量"
        frmMake.Show 1
    Else
        MsgBox "目前的安全存量都足夠!", , "金光益電器股份有限公司"
    End If
    
    '重整主表單的功能項目
    Call mnuSell_status
    self_order_flag = 0
End Sub

'顯示銷貨單
Private Sub mnuSell_Click()
    frmSell.Show
    mnuExit.Enabled = False
End Sub

'顯示員工資料表
Private Sub mnuStaff_Click()
    frmStaff.Show
    mnuExit.Enabled = False
End Sub

'製作新資料庫的副程式
Private Function Backup_Database(FileName As String) As Integer
    Dim db As DAO.Database
    Dim sql_rsMake As String
    '1.確認檔案名稱是否存在
    '檔名已存在
    If Dir(FileName) <> "" Then
        If MsgBox("資料庫檔名已存在,是否要覆寫?", 48 + vbYesNo, "金光益電器股份有限公司") = vbYes Then
            '要覆寫,則先將同名之資料庫檔名移除
            Kill FileName
        Else
            '否則即要另外設定檔名
            '要另外為資料庫命名之旗標
            Backup_Database = 1
            Exit Function
        End If
    End If
    
    '2.建立一個資料庫,檔名為變數FileName所對應的路徑與名稱
    Set db = CreateDatabase(FileName, dbLangChineseTraditional)
    db.Close
    
    '3.將原有的資料表內容備份
    sql_rsOrder = "select * into 訂單主檔 in '" & FileName & "' from 訂單主檔"
    sql_rsOrder_Detail = "select * into 訂單明細 in '" & FileName & "' from 訂單明細"
    sql_rsSell = "select * into 銷貨主檔 in '" & FileName & "' from 銷貨主檔"
    sql_rsSell_Detail = "select * into 銷貨明細 in '" & FileName & "' from 銷貨明細"
    sql_rsCheck = "select * into 產品盤存表 in '" & FileName & "' from 產品盤存表"
    sql_rsComputer = "select * into 電腦盤點表 in '" & FileName & "' from 電腦盤點表"
    sql_rsProdtStore = "select * into 產品庫存統計表 in '" & FileName & "' from 產品庫存統計表"
    sql_rsStaff = "select * into 員工資料表 in '" & FileName & "' from 員工資料表"
    sql_rsRelation = "select * into 員工聯絡人資料表 in '" & FileName & "' from 員工聯絡人資料表"
    sql_rsArea = "select * into 地區表 in '" & FileName & "' from 地區表"
    sql_rsDepart = "select * into 部門分類表 in '" & FileName & "' from 部門分類表"
    sql_rsDuty = "select * into 職稱等級表 in '" & FileName & "' from 職稱等級表"
    sql_rsCust = "select * into 客戶資料 in '" & FileName & "' from 客戶資料"
    sql_rsProdt = "select * into 產品資料 in '" & FileName & "' from 產品資料"
    sql_rsCkNo = "select * into 本月盤點編號 in '" & FileName & "' from 本月盤點編號"
    sql_rsMake = "select * into 生產核銷表 in '" & FileName & "' from 生產核銷表"
    
    '4.執行動作查詢
    cn.Execute sql_rsOrder
    cn.Execute sql_rsOrder_Detail
    cn.Execute sql_rsSell
    cn.Execute sql_rsSell_Detail
    cn.Execute sql_rsCheck
    cn.Execute sql_rsComputer
    cn.Execute sql_rsProdtStore
    cn.Execute sql_rsStaff
    cn.Execute sql_rsRelation
    cn.Execute sql_rsArea
    cn.Execute sql_rsDepart
    cn.Execute sql_rsDuty
    cn.Execute sql_rsCust
    cn.Execute sql_rsProdt
    cn.Execute sql_rsCkNo
    cn.Execute sql_rsMake
    
    '表示資料庫中已有備份的資料
    Backup_Database = 0
End Function

'銷貨單查詢與生產核銷功能的狀態
Public Sub mnuSell_status()
    '1.決定銷貨單查詢功能可否被操作
    Set rsSell = New ADODB.Recordset
    rsSell.Open "select * from 銷貨主檔", cn, adOpenDynamic, adLockOptimistic
    '當有銷貨資料時,銷貨單查詢功能才可使用
    If rsSell.RecordCount <> 0 Then
        mnuSell.Enabled = True
    Else
        mnuSell.Enabled = False
    End If
    
    '2.決定生產核銷功能可否被操作
    Set rsOrder = New ADODB.Recordset
    rsOrder.Open "select * from 訂單主檔", cn, adOpenDynamic, adLockOptimistic
    
    rsOrder.Filter = "銷貨核銷=0"
    '若有訂貨單還未核銷(還未有銷貨單),生產核銷表則可查詢
    If rsOrder.RecordCount <> 0 Then
        mnuMkCk.Enabled = True
    Else
        '否則無法使用
        mnuMkCk.Enabled = False
    End If
    rsOrder.Filter = adFilterNone
End Sub

'判斷是否為月底,決定月底盤存的功能選項是否可操作
Private Sub check_status()
    Dim rsCkNo As ADODB.Recordset
    Dim now_CkNo As String

    mnuReport.Enabled = False
    mnuInitial.Enabled = False
    
    Set rsCkNo = New ADODB.Recordset
    rsCkNo.Open "select * from 本月盤點編號", cn, adOpenDynamic, adLockOptimistic
    
    ck_Head = Year(Date)
    ck_Num = Month(Date)
    ck_Len = Len(ck_Num)
    
    Select Case ck_Len
        Case 1
            ck_No = ck_Head & "0" & ck_Num
        Case 2
            ck_No = ck_Head & ck_Num
    End Select
    
    now_CkNo = "C" & ck_No
    
    If Day(Date) >= 28 Then
        If rsCkNo.RecordCount <> 0 Then
            rsCkNo.MoveFirst
        End If
        rsCkNo.Find "盤點編號='" & now_CkNo & "'", , 1
        If rsCkNo.EOF Then
            mnuReport.Enabled = True
            mnuRptComputer.Enabled = True
            mnuEditCheck.Enabled = False
            MsgBox "請準備做月底的盤點作業!", , "金光益電器股份有限公司"
        Else
            MsgBox "本月的盤點作業已完成", , "金光益電器股份有限公司"
        End If
    Else
        mnuReport.Enabled = False
        mnuInitial.Enabled = False
    End If
End Sub
