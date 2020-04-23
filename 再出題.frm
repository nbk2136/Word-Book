VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 再出題 
   Caption         =   "何%以下をソート??"
   ClientHeight    =   1704
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3468
   OleObjectBlob   =   "再出題.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "再出題"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim i As Integer
    Dim j As Integer
    Dim YOKO As Integer
    Dim N As Integer '残り回答の記載行
    Dim M As Integer '正解率の記載行
    
    N = 1
    Do Until Cells(7, N) = "残り回答"
        N = N + 1
    Loop
    
    M = 1
    Do Until Cells(7, M) = "正解率"
        M = M + 1
    Loop
    
    YOKO = Cells(Rows.Count, 3).End(xlUp).Row
    
    j = 1
    For i = 8 To YOKO
        If Cells(i, M) <= TextBox1 / 100 Then
            Cells(i, 2) = j
            Cells(i, N) = "不正解"
                j = j + 1
        Else
            Cells(i, 2) = ""
            Cells(i, N) = ""
        End If
    Next i
        
    Unload Me
    
    
End Sub


Private Sub CommandButton2_Click()
    
    Unload Me
    
End Sub

Private Sub UserForm_Click()

End Sub
