VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �ďo�� 
   Caption         =   "��%�ȉ����\�[�g??"
   ClientHeight    =   1704
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3468
   OleObjectBlob   =   "�ďo��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�ďo��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim i As Integer
    Dim j As Integer
    Dim YOKO As Integer
    Dim N As Integer '�c��񓚂̋L�ڍs
    Dim M As Integer '���𗦂̋L�ڍs
    
    N = 1
    Do Until Cells(7, N) = "�c���"
        N = N + 1
    Loop
    
    M = 1
    Do Until Cells(7, M) = "����"
        M = M + 1
    Loop
    
    YOKO = Cells(Rows.Count, 3).End(xlUp).Row
    
    j = 1
    For i = 8 To YOKO
        If Cells(i, M) <= TextBox1 / 100 Then
            Cells(i, 2) = j
            Cells(i, N) = "�s����"
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
