VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �P�꒠ 
   Caption         =   "�P�꒠"
   ClientHeight    =   5976
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9768
   OleObjectBlob   =   "�P�꒠.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�P�꒠"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Dim i As Integer  '���̍Ō�̗�
    Dim j As Integer  '���̍ŏ��̗�
    Dim k As Integer  '�J�E���^�[
    Dim A As Integer  '���ԍ������Ɍ���
    Dim B As String   '�����i�[
    Dim C As Integer  '��萔
    Dim D As String
    Dim E As String
    Dim F As String
    Dim AA As Integer
    
    
Private Sub CommandButton1_Click()
    
    '�V�[�g1��ϐ��֑��
    Dim WS1 As Worksheet
    Set WS1 = Worksheets("sheet1")
    
    '�V�[�g3��ϐ��֑��
    Dim WS3 As Worksheet
    Set WS3 = Worksheets("sheet3")
    
    'WS1.Activate
    
    '�y�ϐ��z�u�c��̉𓚁v�L�ڍs
    Dim N As Integer
    N = 1
    Do Until WS1.Cells(7, N) = "�c���"
        N = N + 1
    Loop
    
    D = Replace(StrConv(LCase(TextBox2), vbNarrow), " ", "")
    E = Replace(StrConv(LCase(WS1.Cells(j + k, 4)), vbNarrow), " ", "")
    F = Replace(StrConv(LCase(WS1.Cells(j + k, 5)), vbNarrow), " ", "")
    
    If (D <> "") And (D = E Or D = F) Then
        WS3.Activate
        
        MsgBox "����!!"
        
        'Worksheets("sheet1").Activate
        WS1.Cells(k + j, N - 3) = WS1.Cells(k + j, N - 3) + 1
        WS1.Cells(k + j, N) = ""
    
    Else
        B = WS1.Cells(k + j, 4)
        WS3.Activate
        MsgBox "�s����!!", vbCritical
        MsgBox "������" & vbCrLf & B
        'Worksheets("sheet1").Activate
        WS1.Cells(k + j, N - 2) = WS1.Cells(k + j, N - 2) + 1
        WS1.Cells(k + j, N) = "�s����"
        
        If WS1.Cells(5, N - 3) > 0 Then
        
            Dim z As Integer
            Dim x As Integer
            Dim KAKUNIN As String
            z = Val(WS1.Cells(5, N - 3))
            MsgBox "���K" & "�񓚂�" & vbCrLf & z & "����͂��܂��傤�B", vbInformation
            For x = 1 To z
                WS3.Activate
                KAKUNIN = InputBox(WS1.Cells(j + k, 3).Text, "��" & z - x + 1 & "��", "�񓚂��L�����ĉ������B")
                D = Replace(StrConv(LCase(KAKUNIN), vbNarrow), " ", "")
                If KAKUNIN = "" Then
                    Exit For
                ElseIf (D <> E) And (D <> F) Then
                    MsgBox "�񓚂��Ԉ���Ă��܂��B", vbCritical
                    MsgBox "������������" & vbCrLf & WS1.Cells(j + k, 4).Text
                    x = x - 1
                End If
            Next x
        
        End If
            
        'Worksheets("sheet1").Activate
        
    End If
    
    TextBox2 = ""
    TextBox2.SetFocus

    
    A = A + 1
    Label3.Caption = C & " �� �� " & A & " ��"
    AA = WS1.Cells(6, N)
    Label4.Caption = "�c��񓚐� " & AA & " ��"
    
    Dim YOKO As Integer
    YOKO = WS1.Cells(Rows.Count, N).End(xlUp).Row
    If YOKO = 7 Then
        MsgBox "���͂��ׂĐ������܂����B" & vbCrLf & "�����I�����܂��B"
        Unload Me
        
        Worksheets("sheet1").Activate
        
        Exit Sub
        
    End If
        
        If A > C Then
            Unload Me
            MsgBox "�s��������ďo�肵�܂��B"
            
            Call �s����
            
        End If
        
        For k = 1 To i
            If WS1.Cells(k + j, 1) = A Then
                TextBox1 = WS1.Cells(k + j, 3).Text
                    GoTo nxt2
            End If
        Next k
nxt2:
        
        WS3.Activate
        
End Sub

Private Sub CommandButton2_Click()

    Unload Me
    Sheets(1).Activate
    
    
End Sub


Private Sub TextBox1_Change()

End Sub

'**********�@���o��A�c��𓚐��A�𓚐� **********

Private Sub UserForm_Activate()

    Dim N As Integer '�c��񓚂̋L�ڍs
    
    Dim WS1 As Worksheet
    Set WS1 = Worksheets("sheet1")
    
    Dim WS3 As Worksheet
    Set WS3 = Worksheets("sheet3")
    
    N = 1
    Do Until WS1.Cells(7, N) = "�c���"
        N = N + 1
    Loop
    
    'Worksheets("sheet1").Activate
    
    i = WS1.Cells(Rows.Count, 2).End(xlUp).Row
    j = WS1.Cells(1, 2).End(xlDown).Row
    
    For k = j + 1 To i
        If WS1.Cells(k, 1) <> "" Then
            C = C + 1
        End If
    Next k
    
    A = 1
    For k = 1 To i
        If WS1.Cells(k + j, 1) = A Then
            TextBox1 = WS1.Cells(k + j, 3).Text
            GoTo nxt1
        End If
    Next k
    
nxt1:

    WS3.Activate
    TextBox2.SetFocus
    Label3.Caption = C & " �� �� " & A & " ��"
    AA = WS1.Cells(6, N)
    Label4.Caption = "�c��񓚐� " & AA & " ��"
    
End Sub


