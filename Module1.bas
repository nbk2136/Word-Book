Attribute VB_Name = "Module1"
Sub �A��()

    Dim i As Integer
    Dim j As Integer
    Dim N As Integer '�c��񓚂̋L�ڍs
    
    N = 1
    Do Until Cells(7, N) = "�c���"
        N = N + 1
    Loop
    
    i = Cells(1, 3).End(xlDown).Offset(1, 0).Row
    j = 1
    
    Do While Cells(7 + j, 2) <> ""
        Cells(7 + j, 2) = ""
        Cells(7 + j, N) = ""
        j = j + 1
    Loop
    
    j = 1
    
    Do While Cells(i, 3) <> ""
        Cells(i, 2) = j
        Cells(i, N) = "�s����"
        i = i + 1
        j = j + 1
    Loop
        
    
    
End Sub

Sub ����()

    Dim �ő�l As Integer
    Dim ���[ As Integer
    Dim A() As Integer
    Dim ���� As Integer
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer '�A�Ԃ̓��ʒu������

    Dim tmp As Integer
    
    j = Cells(1, 2).End(xlDown).Row
    �ő�l = Cells(Rows.Count, 2).End(xlUp).Value
    ���[ = Cells(Rows.Count, 3).End(xlUp).Row
    
    For k = j To ���[
        Cells(k, 1) = ""
    Next k
    
    ReDim A(1 To �ő�l)     'ReDim ���I�z��
    
    Randomize                '�����W�F�l���[�^��������
    
    k = 1
    For i = 1 To ���[
        If Cells(i + j, 2) <> "" Then
            A(k) = Cells(i + j, 2)
                k = k + 1
        End If
    Next i
    
    For k = 1 To �ő�l
        ���� = Int((�ő�l - 1 + 1) * Rnd + 1)
            
        tmp = A(k)
        A(k) = A(����)
        A(����) = tmp
        
    Next k
    
    k = 1
    For i = 1 To ���[
        If Cells(i + j, 2) <> "" Then
            Cells(i + j, 1) = A(k)
            k = k + 1
        End If
    Next i
    
    

End Sub

Sub ���o��()

    Dim A As Date
    Dim B As Date
    Dim C As Integer
    Dim D As Integer
    

    Call �A��

    Call ����
    
    A = Now()

    �P�꒠.Show
    
    B = Now()
    C = DateDiff("n", A, B)
    MsgBox "�񓚎��Ԃ́A" & C & "���ł��B"
    
End Sub

Sub �ĊJ()

    Dim A As Date
    Dim B As Date
    Dim C As Integer
    
    A = Now()
    
    �s����.Show
    
    B = Now()
    C = DateDiff("n", A, B)
    MsgBox "�𓚎��Ԃ́A" & C & "�@���ł��B"
    
    
End Sub


Sub �ďo()

    Dim A As Date
    Dim B As Date
    Dim C As Integer
    
    �ďo��.Show
    
    Call ����

    A = Now()

    �P�꒠.Show
    
    B = Now()
    C = DateDiff("n", A, B)
    MsgBox "�񓚎��Ԃ́A" & C & "���ł��B"
    
    
End Sub

Sub �s����()
'�ďo���A2�s�ڂ֕s�����̔ԍ���U��

    Dim i As Integer
    Dim YOKO As Integer
    Dim A As Integer
    Dim N As Integer '�c��񓚂̋L�ڍs
    
    N = 1
    Do Until Cells(7, N) = "�c���"
        N = N + 1
    Loop


    '�A�Ԃ̍ŏI�s���擾
    YOKO = Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 8 To YOKO
        Cells(i, 2) = ""
    Next i
    
    A = 1
    For i = 8 To YOKO
        If Cells(i, N) = "�s����" Then
            Cells(i, 2) = A
            A = A + 1
        End If
    Next i
    
    Call ����
    
    �P�꒠.Show
    
End Sub


