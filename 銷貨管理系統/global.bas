Attribute VB_Name = "global"
Option Explicit
'Connection ���ŧi
Public cn As ADODB.Connection

'*********Recordset���ŧi**********
'�q�f��(�q��D��/����)
Public rsOrder As ADODB.Recordset
Public rsOrder_Detail As ADODB.Recordset
Public sql_rsOrder As String
Public sql_rsOrder_Detail As String
'�P�f��(�P�f�D��/����)
Public rsSell As ADODB.Recordset
Public rsSell_Detail As ADODB.Recordset
Public sql_rsSell As String
Public sql_rsSell_Detail As String
'���~�w�s�έp��
Public rsProdtStore As ADODB.Recordset
Public sql_rsProdtStore As String
'���~�L�s��P�q���L�I��
Public rsCheck As ADODB.Recordset
Public sql_rsCheck As String
Public rsComputer As ADODB.Recordset
Public sql_rsComputer As String
Public cmd As ADODB.Command
'���u��ƪ�
Public rsStaff As ADODB.Recordset
Public sql_rsStaff As String
'���u�p���H��ƪ�
Public rsRelation As ADODB.Recordset
Public sql_rsRelation As String
'�Ȥ���
Public rsCust As ADODB.Recordset
Public sql_rsCust As String
'���~���
Public rsProdt As ADODB.Recordset
Public sql_rsProdt As String
'�a�Ϫ�
Public rsArea As ADODB.Recordset
Public sql_rsArea As String
'�Ͳ��־P��
Public rsMake As ADODB.Recordset


'********�`�Ϊ��ܼƫŧi*********
Public i As Integer             '�Ω�j�骺�ܼ�
Public oText As TextBox         '�Ω�TextBox���ܼ�
Public pre_prodt As String      '�O���e�@�Ӳ��~
Public bm_move As Variant       '�ΨӰO���ثe����intRecord
Public intAccount(1) As Integer '�q����Ӫ��e�ᵧ��(�־P��)
Public Add_flag As Integer      '�s�W�s�����A
Public Edit_flag As Integer     '�ק�s�����A
Public cmdSell_flag As Integer  '�P�f�檺�ӷ�
'rs�@�Ϊ�RecordSet��index��,����������(1),¾�ٵ��Ū�(2),�a�Ϫ�(3)
Public Menu_index As Integer
Public intNowNum As Integer     '�x�s�w�s���Ȧs�q�ܼ�

'��Ʋ��ʮ�,�O����ثe���ƻP�`����
Public intTotal As Integer
Public intRecord As Integer
'�۰ʽs���ҨϥΤ��ܼ�
Public Main_Head As String
Public Main_No As String
Public Sub_No As String
Public Main_Num As String
Public Main_Len As Integer

'��Ʋ��ʤ��Ƶ{��
Public Sub rs_Move(inx As Integer, rs As ADODB.Recordset)
    '�q���`����
    intTotal = rs.RecordCount
    Select Case inx
        Case 0  '�Ĥ@��
            rs.MoveFirst
            intRecord = 1
            
        Case 1  '�e�@��
            rs.MovePrevious
            If rs.BOF Then
                rs.MoveFirst
            End If
            intRecord = intRecord - 1
            If intRecord < 1 Then intRecord = 1
            
        Case 2  '�U�@��
            rs.MoveNext
            If rs.EOF Then
                rs.MoveLast
            End If
            intRecord = intRecord + 1
            If intRecord > intTotal Then intRecord = intTotal
            
        Case 3  '�̥���
            rs.MoveLast
            intRecord = intTotal
    End Select
    
End Sub

'�D�ɽs�����۰ʷs�W
Public Sub Add_No(rs As ADODB.Recordset, indx As Integer, num_index As Integer)
    
    Main_Head = ""
    Select Case indx
        Case 1  '�Ȥ�
            Main_Head = "C"
        Case 2  '���u
            Main_Head = "E"
        Case 3  '�q��
            Main_Head = "O"
        Case 4  '�P�f��
            Main_Head = "S"
        Case 5  '���~
            Main_Head = "P"
        Case 6  '�a�Ϫ�
            Main_Head = "A"
    End Select
    
    If rs.RecordCount <> 0 Then
        Select Case indx
            Case 1  '�Ȥ�
                rs.Sort = "�Ȥ�s�� asc"
            Case 2  '���u
               rs.Sort = "���u�s�� asc"
            Case 3  '�q��
                rs.Sort = "�q��s�� asc"
            Case 4  '�P�f��
                rs.Sort = "�P�f��s�� asc"
            Case 5  '���~
                rs.Sort = "���~�s�� asc"
            Case 6  '�a�Ϫ�
                rs.Sort = "�a�Ͻs�� asc"
        End Select
        
        rs.MoveLast
        Select Case num_index
            Case 1      'Main_Num���T���
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
            Case 2      'Main_Num���G���
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

'���ӽs�����۰ʷs�W
Public Sub Add_SubNo(rsD As ADODB.Recordset, indx As Integer, txtMain_No As String)
    Dim Sub_Head As String
    Dim Sub_ASC As String
    Dim Sub_Num As Integer
    
    Sub_Head = txtMain_No
    Set rsD = New ADODB.Recordset
    Select Case indx
        Case 1
            rsD.Open "select * from �q����� where �q��s��= '" & Sub_Head & "'", cn, adOpenDynamic, adLockOptimistic
        Case 2
            rsD.Open "select * from �P�f���� where �P�f��s��='" & Sub_Head & "'", cn, adOpenDynamic, adLockOptimistic
    End Select
    
    '�Y���Ӧ���
    If rsD.RecordCount <> 0 Then
        Select Case indx
            Case 1  '�q��
                rsD.Sort = "�q����ӽs�� asc"
                
            Case 2  '�P�f��
                rsD.Sort = "�P�f����ӽs�� asc"
        End Select
        rsD.MoveLast
        Sub_Num = Asc(Right(rsD.Fields(1), 1)) + 1
        Sub_ASC = Chr(Sub_Num)
        Sub_No = Sub_Head & Sub_ASC
    Else
        Sub_No = Sub_Head & "a"
    End If
End Sub


