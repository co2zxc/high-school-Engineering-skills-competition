VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "�P�f�޲z�t��  �D�e��"
   ClientHeight    =   4350
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6630
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '�t�ιw�]��
   WindowState     =   2  '�̤j��
   Begin VB.Menu mnuInfo 
      Caption         =   "�򥻸��"
      Begin VB.Menu mnuStaffKind 
         Caption         =   "���u��ƺ޲z"
         Begin VB.Menu mnuStaff 
            Caption         =   "���u��ƪ�"
         End
      End
      Begin VB.Menu mnuline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBusiness 
         Caption         =   "�Ȥ��ƺ޲z"
         Begin VB.Menu mnuCust 
            Caption         =   "�Ȥ��ƪ�"
         End
      End
      Begin VB.Menu mnuline4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAreaind 
         Caption         =   "�a�ϸ�ƺ޲z"
         Begin VB.Menu mnuArea 
            Caption         =   "�a�Ϫ�"
         End
      End
      Begin VB.Menu mnuline5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObject 
         Caption         =   "���~��ƺ޲z"
         Begin VB.Menu mnuProdt 
            Caption         =   "���~��ƪ�"
         End
         Begin VB.Menu mnuline12 
            Caption         =   "-"
         End
         Begin VB.Menu mnub 
            Caption         =   "�t�Ӹ��"
         End
         Begin VB.Menu mnuC 
            Caption         =   "���O���"
         End
      End
      Begin VB.Menu mnuline13 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuSheet 
      Caption         =   "�إ߳��"
      Begin VB.Menu mnuProdtKind 
         Caption         =   "���~���O"
         Begin VB.Menu mnuOrder 
            Caption         =   "�q�f��"
         End
         Begin VB.Menu mnuline6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSell 
            Caption         =   "�P�f��d��"
         End
      End
   End
   Begin VB.Menu mnuMkCk 
      Caption         =   "�Ͳ��־P"
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "�ݨD�d��"
      Begin VB.Menu mnuProdtStore 
         Caption         =   "���~�w�s�έp��"
      End
      Begin VB.Menu mnuline7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSafeNum 
         Caption         =   "�w���s�q���ˬd"
      End
      Begin VB.Menu mnuline8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQ1 
         Caption         =   "�ӫ~�d��(�U�Ԧ�)"
      End
      Begin VB.Menu mnuQ2 
         Caption         =   "�ӫ~�d��(����r)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "�إ߳���"
      Begin VB.Menu mnuRptComputer 
         Caption         =   "�s�@�q���L�s����"
      End
      Begin VB.Menu mnuline9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "�s�貣�~�L�s��"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "�u��"
      Begin VB.Menu mnuQ3 
         Caption         =   "�p���"
      End
      Begin VB.Menu mnuEasy 
         Caption         =   "�K��"
      End
      Begin VB.Menu mnuQ4 
         Caption         =   "MP3�����"
      End
      Begin VB.Menu mnuline10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "��ƪ��ƥ�"
      End
      Begin VB.Menu mnuline14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInitial 
         Caption         =   "�إߤU����������"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "����"
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mnuSafeNum_flag As Integer   '�����Ͳ��n�ɨ��w���s�q���X��
Dim DBName As String                '�^�ǩҿ�J����Ʈw�ɮ׸��|���r��
Dim sql_rsCkNo As String
'�L�s��ƪ��ܼ�
Dim ck_Head As String
Dim ck_Num As String
Dim ck_Len As Integer
Dim ck_No As String

'��Ʈw���s��
Private Sub MDIForm_Load()
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    ChDir (App.Path)
    cn.Open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbRose.mdb"
    
    '���~��Ʊ`�Q�ϥ�,�G�b�����J�ɧY���}��
    Set rsProdt = New ADODB.Recordset
    rsProdt.Open "select * from ���~���", cn, adOpenDynamic, adLockOptimistic
        
    '�멳�L�s�\�઺���A
    Call check_status
    '�P�f��d�ߥ\�઺���A
    Call mnuSell_status
End Sub

Private Sub mnub_Click()
Unload Me: frmOrd_Detail2.Show
End Sub

'�s�@��ƪ��ƥ�
Private Sub mnuBackup_Click()
    '�]�w�ƥ���Ʈw���ɮ׸��|
    DBName = InputBox("�п�J�ƥ���Ʈw���ɮ׸��|!!" & "�pC:\VB6\DataBase\Backup.mdb", "�����q�q���ѥ��������q")
    If DBName = "" Then
        Exit Sub
    Else
        If Right(DBName, 4) <> ".mdb" Then
        MsgBox "�Ь���Ʈw�W�٥[�W���ɦW'.mdb'"
        Call mnuBackup_Click
    End If

    End If
    
    '�s�@�ƥ���Ʈw
    'Backup_Database=1��,�����мg���e�ҿ�J���ɦW,�G���A��J�s���ɦW
    If Backup_Database(DBName) = 1 Then
        Call mnuBackup_Click
    Else
        MsgBox ("��Ʈw�w�ƥ�����!!")
    End If
End Sub


Private Sub mnuC_Click()
Unload Me
frmOrd_Detail1.Show
End Sub

'****���������� ¾�ٵ��Ū�P�a�Ϫ�@�Τ@�Ӫ��,�ҥH�������@��Index����,
'�ӶǤJ���P��RecordSet****
'��ܳ���������
Private Sub mnuDepart_Click()
    Menu_index = 1
    mnuArea.Enabled = False
    mnuDuty.Enabled = False
    frmInfo.Show
    mnuExit.Enabled = False
End Sub

'���¾�ٵ��Ū�
Private Sub mnuDuty_Click()
    Menu_index = 2
    mnuArea.Enabled = False
    mnuDepart.Enabled = False
    
    frmInfo.Show
    mnuExit.Enabled = False
End Sub

'��ܦa�Ϫ�
Private Sub mnuArea_Click()
    Menu_index = 3
    
    frmInfo.Show
    mnuExit.Enabled = False
End Sub

'��ܫȤ��ƪ�
Private Sub mnuCust_Click()
    mnuExit.Enabled = False
    frmCust.Show
End Sub

Private Sub mnuEasy_Click()
 Unload Me
frmDrawLine.Show
End Sub

'��ܽs�貣�~�L�s��
Private Sub mnuEditCheck_Click()
    frmCheck.Show
    mnuExit.Enabled = False
End Sub

'���}���t��
Private Sub mnuExit_Click()
    cn.Close
    Set cn = Nothing
    End
End Sub

'�إߤU����������
Private Sub mnuInitial_Click()
   ' Dim old_DBName As String
        
   ' If Add_flag = 1 Then
   '     DBName = old_DBName
   ' End If
   ' Add_flag = 0
        
    '1.�]�w��Ʈw���ɮ׸��|
    DBName = InputBox("�п�J�e�����v�ɪ��ƥ��ɮ׸��|!" & "�pC:\VB6\DataBase\History.mdb", "�����q�q���ѥ��������q")

    If DBName = "" Then
        Exit Sub
    Else
       ' Add_flag = 1
       ' old_DBName = DBName
        If Right(DBName, 4) <> ".mdb" Then
            MsgBox "�Ь���Ʈw�W�٥[�W���ɦW'.mdb'"
            Call mnuInitial_Click
        End If
    End If
    
    '2.�s�@���v�ɪ��ƥ�
    'Backup_Database=1��,�����мg���e�ҿ�J���ɦW,�G���A��J�s���ɦW
    If Backup_Database(DBName) = 1 Then
        Call mnuInitial_Click
    Else
        '�ƥ�����ڸ�ƫ�
        '(1).�M���q�f��,�P�f��P���~�w�s�έp����ƪ����
        sql_rsOrder = "delete * from �q��D��"
        sql_rsOrder_Detail = "delete * from �q�����"
        sql_rsSell = "delete * from �P�f�D��"
        sql_rsSell_Detail = "delete * from �P�f����"
        sql_rsProdtStore = "delete * from ���~�w�s�έp��"
        
        '����ʧ@�d��
        cn.Execute sql_rsOrder
        cn.Execute sql_rsOrder_Detail
        cn.Execute sql_rsSell
        cn.Execute sql_rsSell_Detail
        cn.Execute sql_rsProdtStore
        
        '(2).�N�q���L�I�����~�s���P��ڽL�s�q���ȷs�W�i���~�w�s�έp��
        sql_rsComputer = "insert into ���~�w�s�έp��(���~�s��,���~�W��,�W���w�s�q) select ���~�s��,���~�W��,��ڽL�s�q from �q���L�I��"
        cn.Execute sql_rsComputer
        
        '(3).�A�N�w���s�q��i���~�w�s�έp��
        Set rsProdtStore = New ADODB.Recordset
        rsProdtStore.Open "select * from ���~�w�s�έp��", cn, adOpenDynamic, adLockOptimistic

        Do Until rsProdtStore.EOF
            rsProdt.MoveFirst
            rsProdt.Find "���~�s��='" & rsProdtStore.Fields("���~�s��") & "'", , 1
            rsProdtStore.Fields("�w���s�q") = rsProdt.Fields("�w���s�q")
            rsProdtStore.Fields("�Ȧs�q") = rsProdtStore.Fields("�{�s�q")
            
            rsProdtStore.MoveNext
        Loop
        
        '(4).�N�L�I�s���O����'����L�I�s��'��ƪ�
        sql_rsCkNo = "insert into ����L�I�s��(�L�I�s��) select �L�I�s�� from ���~�L�s�� "
        cn.Execute sql_rsCkNo
        
        '(5).���ۧY�������~�L�s��P�q���L�I�����
        sql_rsCheck = "delete * from ���~�L�s��"
        sql_rsComputer = "delete * from �q���L�I��"
        cn.Execute sql_rsCheck
        cn.Execute sql_rsComputer
        
        '�����ϥΪ�,�������Ҥw�إߧ���
        MsgBox ("�U���ťճ�ڤw�غc����!!")
        '����D��椤���\�ඵ��
        Call check_status
    End If
End Sub

'��ܥͲ��־P��
Private Sub mnuMkCk_Click()
    frmMkCk.Show 1
End Sub

'��ܭq�f��
Private Sub mnuOrder_Click()
    frmOrder.Show
    mnuExit.Enabled = False
End Sub

'��ܲ��~��ƪ�
Private Sub mnuProdt_Click()
    frmProdt.Show
    mnuExit.Enabled = False
End Sub

'��ܲ��~�w�s�έp��
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

'�إ߻P��ܹq���L�s����
Private Sub mnuRptComputer_Click()
    '�M�����~�L�s��ҨϥΪ��ʧ@�d�߫��O
    Dim sql_CheckClear As String
    '�N�����~�L�s�����L�I�s����줧�ܼ�
    Dim Check_No As String

    '1.�C���I��"�إ߹q���L�s����"���خ�,���|���N��ƲM����,�A���sInsert�@��
    sql_CheckClear = "delete from ���~�L�s��"
    '����ʧ@�d��
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = sql_CheckClear
    cmd.Execute
    
    '����ӪŪ����~�L�s��,�s�W��ưO��
    Call CheckNo_Add
        
    '2.�}�ҹq���L�I��
    Set rsComputer = New ADODB.Recordset
    sql_rsComputer = "select * from �q���L�I��"
    rsComputer.Open sql_rsComputer, cn, adOpenDynamic, adLockOptimistic

    '3.�s�@���d�߫��O,�N���~�w�s�έp�����~�s���P�{�s�q����Insert�ܹq���L�I��
    sql_rsProdtStore = "insert into �q���L�I��(���~�s��,���~�W��,�����w�s�q) select ���~�s��,���~�W��,�{�s�q from ���~�w�s�έp��"
    '����ʧ@�d��
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = sql_rsProdtStore
    cmd.Execute
    '�N�ȶ�J�q���L�I��
    rsComputer.Requery
    
    Do Until rsComputer.EOF
        Check_No = rsCheck.Fields("�L�I�s��")
        '�N�L�I�s��Insert�i��
        rsComputer.Fields("�L�I�s��") = Check_No
        rsComputer.Update
        rsComputer.MoveNext
    Loop
   
    '4.��ܥX�q���L�s����
    RptComputer.Show
End Sub

'�s�W���~�L�s���L�I�s���P�L�I�������
Private Sub CheckNo_Add()
    Set rsCheck = New ADODB.Recordset
    sql_rsCheck = "select * from ���~�L�s��"
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
    rsCheck.Fields("�L�I�s��") = ck_No
    rsCheck.Fields("�L�I���") = Date
    rsCheck.Update
End Sub

'�ˬd�w���s�q�O�_����,�M�w�O�_�n�Ͳ�
Private Sub mnuSafeNum_Click()
    Dim self_order_flag As Integer
    '���Ͳ����ƶq
    Dim intSelf_OrNum As Integer
        
    Set rsOrder = New ADODB.Recordset
    rsOrder.Open "select * from �q��D��", cn, adOpenDynamic, adLockOptimistic
    
    Set rsOrder_Detail = New ADODB.Recordset
    rsOrder_Detail.Open "select * from �q�����", cn, adOpenDynamic, adLockOptimistic
    
    Set rsProdtStore = New ADODB.Recordset
    rsProdtStore.Open "select * from ���~�w�s�έp��", cn, adOpenDynamic, adLockOptimistic

    
    '1.�ˬd�C�@�Ӳ��~���{�s�q�O�_�C��w���s�q
    Do Until rsProdtStore.EOF
        '�Y�{�s�q���t��,�Y��ܮw�s�q�w��0
        If Sgn(rsProdtStore.Fields("�Ȧs�q")) = -1 Then
            intNowNum = 0
        Else
            intNowNum = rsProdtStore.Fields("�Ȧs�q")
        End If
        '�N���Ͳ����ƶq�H�ܼ�intSelf_OrNum�O��,�H��K�ϥ�
        intSelf_OrNum = intNowNum - rsProdtStore.Fields("�w���s�q")
        '2.��{�s�q�C��w���s�q��,�Y�n�i��Ͳ�,���Ͳ������̾ڭq�f��,�G���B�����ͤ@�i _
           �Ȥᬰ���q�������q�f��,�M��A�̾ھP�f�y�{�Ӷi��Ͳ�
        If Sgn(intSelf_OrNum) = -1 Then
            '�q��D�ɪ��s�W
            If self_order_flag = 0 Then
                Call Add_No(rsOrder, 3, 1)
                rsOrder.Fields(0) = Main_No
                rsOrder.Fields(1) = "E017"
                '�Y�������q�q���ѥ��������q����
                rsOrder.Fields(2) = "C000"
                rsOrder.Fields(3) = Date
                rsOrder.Fields(4) = 0
                rsOrder.Update
                '�s�W�q��D�ɪ��X�Э�
                self_order_flag = 1
            End If
            
            '�q����Ӫ��s�W
            Call Add_SubNo(rsOrder_Detail, 1, Main_No)
            rsOrder_Detail.AddNew
            rsOrder_Detail.Fields(0) = Main_No
            rsOrder_Detail.Fields(1) = Sub_No
            rsOrder_Detail.Fields(2) = rsProdtStore.Fields(0)
            rsOrder_Detail.Fields(3) = rsProdtStore.Fields(1)
            rsProdt.MoveFirst
            rsProdt.Find "���~�s��='" & rsProdtStore.Fields(0) & "'", , 1
            rsOrder_Detail.Fields(4) = rsProdt.Fields("��즨��")
            rsOrder_Detail.Fields(5) = Abs(intSelf_OrNum)
            rsOrder_Detail.Fields(6) = rsProdt.Fields("��즨��") * Abs(intSelf_OrNum)
            rsOrder_Detail.Fields(7) = 0
            rsOrder_Detail.Fields(8) = 0
            rsOrder_Detail.Update
        End If
        
        rsProdtStore.MoveNext
    Loop
    
    '3.�Y���ݭn�Ͳ����ƶq,�H�ɨ��w���s�q,�h�HfrmMake���e�{�X���Ө�
    If self_order_flag = 1 Then
        '�ѩ�frmMake�Q'�q�f��'�P'�w���s�q���ˬd'�Ҧ@��,���X�Ь��@�ӰϧO��
        mnuSafeNum_flag = 1
        frmMake.Caption = "�Ͳ���-�ɨ��w���s�q"
        frmMake.Show 1
    Else
        MsgBox "�ثe���w���s�q������!", , "�����q�q���ѥ��������q"
    End If
    
    '����D��檺�\�ඵ��
    Call mnuSell_status
    self_order_flag = 0
End Sub

'��ܾP�f��
Private Sub mnuSell_Click()
    frmSell.Show
    mnuExit.Enabled = False
End Sub

'��ܭ��u��ƪ�
Private Sub mnuStaff_Click()
    frmStaff.Show
    mnuExit.Enabled = False
End Sub

'�s�@�s��Ʈw���Ƶ{��
Private Function Backup_Database(FileName As String) As Integer
    Dim db As DAO.Database
    Dim sql_rsMake As String
    '1.�T�{�ɮצW�٬O�_�s�b
    '�ɦW�w�s�b
    If Dir(FileName) <> "" Then
        If MsgBox("��Ʈw�ɦW�w�s�b,�O�_�n�мg?", 48 + vbYesNo, "�����q�q���ѥ��������q") = vbYes Then
            '�n�мg,�h���N�P�W����Ʈw�ɦW����
            Kill FileName
        Else
            '�_�h�Y�n�t�~�]�w�ɦW
            '�n�t�~����Ʈw�R�W���X��
            Backup_Database = 1
            Exit Function
        End If
    End If
    
    '2.�إߤ@�Ӹ�Ʈw,�ɦW���ܼ�FileName�ҹ��������|�P�W��
    Set db = CreateDatabase(FileName, dbLangChineseTraditional)
    db.Close
    
    '3.�N�즳����ƪ��e�ƥ�
    sql_rsOrder = "select * into �q��D�� in '" & FileName & "' from �q��D��"
    sql_rsOrder_Detail = "select * into �q����� in '" & FileName & "' from �q�����"
    sql_rsSell = "select * into �P�f�D�� in '" & FileName & "' from �P�f�D��"
    sql_rsSell_Detail = "select * into �P�f���� in '" & FileName & "' from �P�f����"
    sql_rsCheck = "select * into ���~�L�s�� in '" & FileName & "' from ���~�L�s��"
    sql_rsComputer = "select * into �q���L�I�� in '" & FileName & "' from �q���L�I��"
    sql_rsProdtStore = "select * into ���~�w�s�έp�� in '" & FileName & "' from ���~�w�s�έp��"
    sql_rsStaff = "select * into ���u��ƪ� in '" & FileName & "' from ���u��ƪ�"
    sql_rsRelation = "select * into ���u�p���H��ƪ� in '" & FileName & "' from ���u�p���H��ƪ�"
    sql_rsArea = "select * into �a�Ϫ� in '" & FileName & "' from �a�Ϫ�"
    sql_rsDepart = "select * into ���������� in '" & FileName & "' from ����������"
    sql_rsDuty = "select * into ¾�ٵ��Ū� in '" & FileName & "' from ¾�ٵ��Ū�"
    sql_rsCust = "select * into �Ȥ��� in '" & FileName & "' from �Ȥ���"
    sql_rsProdt = "select * into ���~��� in '" & FileName & "' from ���~���"
    sql_rsCkNo = "select * into ����L�I�s�� in '" & FileName & "' from ����L�I�s��"
    sql_rsMake = "select * into �Ͳ��־P�� in '" & FileName & "' from �Ͳ��־P��"
    
    '4.����ʧ@�d��
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
    
    '��ܸ�Ʈw���w���ƥ������
    Backup_Database = 0
End Function

'�P�f��d�߻P�Ͳ��־P�\�઺���A
Public Sub mnuSell_status()
    '1.�M�w�P�f��d�ߥ\��i�_�Q�ާ@
    Set rsSell = New ADODB.Recordset
    rsSell.Open "select * from �P�f�D��", cn, adOpenDynamic, adLockOptimistic
    '���P�f��Ʈ�,�P�f��d�ߥ\��~�i�ϥ�
    If rsSell.RecordCount <> 0 Then
        mnuSell.Enabled = True
    Else
        mnuSell.Enabled = False
    End If
    
    '2.�M�w�Ͳ��־P�\��i�_�Q�ާ@
    Set rsOrder = New ADODB.Recordset
    rsOrder.Open "select * from �q��D��", cn, adOpenDynamic, adLockOptimistic
    
    rsOrder.Filter = "�P�f�־P=0"
    '�Y���q�f���٥��־P(�٥����P�f��),�Ͳ��־P��h�i�d��
    If rsOrder.RecordCount <> 0 Then
        mnuMkCk.Enabled = True
    Else
        '�_�h�L�k�ϥ�
        mnuMkCk.Enabled = False
    End If
    rsOrder.Filter = adFilterNone
End Sub

'�P�_�O�_���멳,�M�w�멳�L�s���\��ﶵ�O�_�i�ާ@
Private Sub check_status()
    Dim rsCkNo As ADODB.Recordset
    Dim now_CkNo As String

    mnuReport.Enabled = False
    mnuInitial.Enabled = False
    
    Set rsCkNo = New ADODB.Recordset
    rsCkNo.Open "select * from ����L�I�s��", cn, adOpenDynamic, adLockOptimistic
    
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
        rsCkNo.Find "�L�I�s��='" & now_CkNo & "'", , 1
        If rsCkNo.EOF Then
            mnuReport.Enabled = True
            mnuRptComputer.Enabled = True
            mnuEditCheck.Enabled = False
            MsgBox "�зǳư��멳���L�I�@�~!", , "�����q�q���ѥ��������q"
        Else
            MsgBox "���몺�L�I�@�~�w����", , "�����q�q���ѥ��������q"
        End If
    Else
        mnuReport.Enabled = False
        mnuInitial.Enabled = False
    End If
End Sub
