Attribute VB_Name = "Module1"
Sub 連番()

    Dim i As Integer
    Dim j As Integer
    Dim N As Integer '残り回答の記載行
    
    N = 1
    Do Until Cells(7, N) = "残り回答"
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
        Cells(i, N) = "不正解"
        i = i + 1
        j = j + 1
    Loop
        
    
    
End Sub

Sub 乱数()

    Dim 最大値 As Integer
    Dim 下端 As Integer
    Dim A() As Integer
    Dim 乱数 As Integer
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer '連番の頭位置を検索

    Dim tmp As Integer
    
    j = Cells(1, 2).End(xlDown).Row
    最大値 = Cells(Rows.Count, 2).End(xlUp).Value
    下端 = Cells(Rows.Count, 3).End(xlUp).Row
    
    For k = j To 下端
        Cells(k, 1) = ""
    Next k
    
    ReDim A(1 To 最大値)     'ReDim 動的配列
    
    Randomize                '乱数ジェネレータを初期化
    
    k = 1
    For i = 1 To 下端
        If Cells(i + j, 2) <> "" Then
            A(k) = Cells(i + j, 2)
                k = k + 1
        End If
    Next i
    
    For k = 1 To 最大値
        乱数 = Int((最大値 - 1 + 1) * Rnd + 1)
            
        tmp = A(k)
        A(k) = A(乱数)
        A(乱数) = tmp
        
    Next k
    
    k = 1
    For i = 1 To 下端
        If Cells(i + j, 2) <> "" Then
            Cells(i + j, 1) = A(k)
            k = k + 1
        End If
    Next i
    
    

End Sub

Sub 問題出題()

    Dim A As Date
    Dim B As Date
    Dim C As Integer
    Dim D As Integer
    

    Call 連番

    Call 乱数
    
    A = Now()

    単語帳.Show
    
    B = Now()
    C = DateDiff("n", A, B)
    MsgBox "回答時間は、" & C & "分です。"
    
End Sub

Sub 再開()

    Dim A As Date
    Dim B As Date
    Dim C As Integer
    
    A = Now()
    
    不正解.Show
    
    B = Now()
    C = DateDiff("n", A, B)
    MsgBox "解答時間は、" & C & "　分です。"
    
    
End Sub


Sub 再出()

    Dim A As Date
    Dim B As Date
    Dim C As Integer
    
    再出題.Show
    
    Call 乱数

    A = Now()

    単語帳.Show
    
    B = Now()
    C = DateDiff("n", A, B)
    MsgBox "回答時間は、" & C & "分です。"
    
    
End Sub

Sub 不正解()
'再出時、2行目へ不正解の番号を振る

    Dim i As Integer
    Dim YOKO As Integer
    Dim A As Integer
    Dim N As Integer '残り回答の記載行
    
    N = 1
    Do Until Cells(7, N) = "残り回答"
        N = N + 1
    Loop


    '連番の最終行を取得
    YOKO = Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 8 To YOKO
        Cells(i, 2) = ""
    Next i
    
    A = 1
    For i = 8 To YOKO
        If Cells(i, N) = "不正解" Then
            Cells(i, 2) = A
            A = A + 1
        End If
    Next i
    
    Call 乱数
    
    単語帳.Show
    
End Sub


