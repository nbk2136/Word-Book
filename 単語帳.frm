VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 単語帳 
   Caption         =   "単語帳"
   ClientHeight    =   5976
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9768
   OleObjectBlob   =   "単語帳.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "単語帳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Dim i As Integer  '問題の最後の列
    Dim j As Integer  '問題の最初の列
    Dim k As Integer  'カウンター
    Dim A As Integer  '問題番号を順に検索
    Dim B As String   '正解格納
    Dim C As Integer  '問題数
    Dim D As String
    Dim E As String
    Dim F As String
    Dim AA As Integer
    
    
Private Sub CommandButton1_Click()
    
    'シート1を変数へ代入
    Dim WS1 As Worksheet
    Set WS1 = Worksheets("sheet1")
    
    'シート3を変数へ代入
    Dim WS3 As Worksheet
    Set WS3 = Worksheets("sheet3")
    
    'WS1.Activate
    
    '【変数】「残りの解答」記載行
    Dim N As Integer
    N = 1
    Do Until WS1.Cells(7, N) = "残り回答"
        N = N + 1
    Loop
    
    D = Replace(StrConv(LCase(TextBox2), vbNarrow), " ", "")
    E = Replace(StrConv(LCase(WS1.Cells(j + k, 4)), vbNarrow), " ", "")
    F = Replace(StrConv(LCase(WS1.Cells(j + k, 5)), vbNarrow), " ", "")
    
    If (D <> "") And (D = E Or D = F) Then
        WS3.Activate
        
        MsgBox "正解!!"
        
        'Worksheets("sheet1").Activate
        WS1.Cells(k + j, N - 3) = WS1.Cells(k + j, N - 3) + 1
        WS1.Cells(k + j, N) = ""
    
    Else
        B = WS1.Cells(k + j, 4)
        WS3.Activate
        MsgBox "不正解!!", vbCritical
        MsgBox "正解は" & vbCrLf & B
        'Worksheets("sheet1").Activate
        WS1.Cells(k + j, N - 2) = WS1.Cells(k + j, N - 2) + 1
        WS1.Cells(k + j, N) = "不正解"
        
        If WS1.Cells(5, N - 3) > 0 Then
        
            Dim z As Integer
            Dim x As Integer
            Dim KAKUNIN As String
            z = Val(WS1.Cells(5, N - 3))
            MsgBox "練習" & "回答を" & vbCrLf & z & "回入力しましょう。", vbInformation
            For x = 1 To z
                WS3.Activate
                KAKUNIN = InputBox(WS1.Cells(j + k, 3).Text, "後" & z - x + 1 & "回", "回答を記入して下さい。")
                D = Replace(StrConv(LCase(KAKUNIN), vbNarrow), " ", "")
                If KAKUNIN = "" Then
                    Exit For
                ElseIf (D <> E) And (D <> F) Then
                    MsgBox "回答が間違っています。", vbCritical
                    MsgBox "正しい答えは" & vbCrLf & WS1.Cells(j + k, 4).Text
                    x = x - 1
                End If
            Next x
        
        End If
            
        'Worksheets("sheet1").Activate
        
    End If
    
    TextBox2 = ""
    TextBox2.SetFocus

    
    A = A + 1
    Label3.Caption = C & " 問 中 " & A & " 問"
    AA = WS1.Cells(6, N)
    Label4.Caption = "残り回答数 " & AA & " 問"
    
    Dim YOKO As Integer
    YOKO = WS1.Cells(Rows.Count, N).End(xlUp).Row
    If YOKO = 7 Then
        MsgBox "問題はすべて正解しました。" & vbCrLf & "問題を終了します。"
        Unload Me
        
        Worksheets("sheet1").Activate
        
        Exit Sub
        
    End If
        
        If A > C Then
            Unload Me
            MsgBox "不正解問題を再出題します。"
            
            Call 不正解
            
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

'**********　問題出題、残り解答数、解答数 **********

Private Sub UserForm_Activate()

    Dim N As Integer '残り回答の記載行
    
    Dim WS1 As Worksheet
    Set WS1 = Worksheets("sheet1")
    
    Dim WS3 As Worksheet
    Set WS3 = Worksheets("sheet3")
    
    N = 1
    Do Until WS1.Cells(7, N) = "残り回答"
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
    Label3.Caption = C & " 問 中 " & A & " 問"
    AA = WS1.Cells(6, N)
    Label4.Caption = "残り回答数 " & AA & " 問"
    
End Sub


