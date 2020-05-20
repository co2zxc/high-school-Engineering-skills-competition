VERSION 5.00
Begin VB.Form frmQ1 
   Caption         =   "產品包裝"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdChangeData 
      Caption         =   "ChangeData"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   744
      Width           =   1095
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "=>計算=>"
      Height          =   495
      Left            =   1793
      TabIndex        =   2
      Top             =   1851
      Width           =   1095
   End
   Begin VB.TextBox txtOut 
      Height          =   2895
      Left            =   2993
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   1
      Top             =   98
      Width           =   1575
   End
   Begin VB.TextBox txtIn 
      Height          =   2895
      Left            =   113
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   0
      Text            =   "frmQ1.frx":0000
      Top             =   98
      Width           =   1575
   End
End
Attribute VB_Name = "frmQ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim flag As Boolean

Private Sub cmdChangeData_Click()
txtIn = ""
flag = Not flag
If flag Then
    txtIn = "2250 50 60 2" & vbCrLf & "8 7 15 12" & vbCrLf & "10 300 23 5" & vbCrLf
Else
    txtIn = "5 4 8 2" & vbCrLf & "12 14 32 7" & vbCrLf
End If
End Sub

Private Sub cmdCount_Click()
Dim data, tmp, prducOut%()

data = Split(txtIn, vbCrLf)
ReDim prducOut(UBound(data) - 1, 3)
Dim i%, j%
For i = 0 To UBound(data) - 1
    tmp = Split(data(i))
    For j = 0 To 3
        prducOut(i, j) = tmp(j)
    Next j
Next i
Dim sum&, t$
Dim P1%, P2%, P3%, P4% '每筆資料中,邊長N公分的產品數目
For i = 0 To UBound(prducOut, 1)
    P1 = prducOut(i, 0)
    P2 = prducOut(i, 1)
    P4 = prducOut(i, 3)
    P3 = prducOut(i, 2)
    
    '4*4*4
    sum = P4
    P4 = 0
    
    '3*3*3
    Debug.Print IIf(P1 \ 37 <= P3, P1 \ 37, P3)
    sum = sum + P3 + IIf(P1 \ 37 <= P3, P1 \ 37, P3)
    P1 = IIf(P1 \ 37 < P3, 0, P1 - P3 * 37)
    P3 = 0
    'P1 = P1 - IIf(P1 \ 37 <= P3, P1 \ 37, P3) 'IIf(P1 \ 37 <= P3, P1 \ 37, P3)
    
    '2*2*2
    sum = sum + P2 \ 8
    Do Until P2 < 8
        P2 = P2 - 8
    Loop
    If P1 = 0 Then
        sum = sum + 1
        P2 = 0
    Else
        If P2 \ 8 = 0 Then
            If P1 <= (4 ^ 3 - P2 * 2 ^ 3) Then
                sum = sum + 1
                P2 = 0
                P1 = 0
            Else
                sum = sum + 1 + (4 ^ 3 - P2 * 2 ^ 3)
                P2 = 0
                P1 = P1 - (4 ^ 3 - P2 * 2 ^ 3)
            End If
        ElseIf P2 \ 8 = 1 Then
            sum = sum + 1
            P2 = 0
        Else
            sum = sum + P2 \ 8
            P2 = P2 - P2 \ 8
            If P2 \ 8 = 0 Then
                If P1 <= (4 ^ 3 - P2 * 2 ^ 3) Then
                    sum = sum + 1
                    P2 = 0
                    P1 = 0
                Else
                    sum = sum + 1 + (4 ^ 3 - P2 * 2 ^ 3)
                    P2 = 0
                    P1 = P1 - (4 ^ 3 - P2 * 2 ^ 3)
                End If
            End If
        End If
    End If
goto_1:
    '1*1*1
    sum = sum + P1
    P1 = 0
    
    t = t & sum & vbCrLf
Next i
txtOut = t
End Sub
