Dim Name As String
Dim FontSize2 As Integer
Dim FontSize1 As Integer

Private Sub UserForm_Initialize()
'//フォントサイズの初期化
FontSize2 = 12
FontSize1 = 12
ActiveCell.Font.Size = 12
Worksheets("座席").Cells.Font.Name = "MS 明朝"
End Sub

Private Sub CheckBox1_Click()
    
If CheckBox1.Value = True Then
    If OptionButton1.Value = True Then
        Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 3).Value & vbLf _
        & Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 1).Value & " " & ComboBox1.Value
    Else
        Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 1).Value & " " & ComboBox1.Value
    End If
Else
    If OptionButton1.Value = True Then
        Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 3).Value & vbLf & ComboBox1.Value
    Else
        Name = ComboBox1.Value
    End If
End If

ActiveCell.Value = Name
'//色
If Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 4) = "M" Then
    ActiveCell.Interior.ColorIndex = 20
Else
    ActiveCell.Interior.ColorIndex = 38
End If

End Sub

Private Sub ComboBox1_Change()
Dim R As Integer
Dim C As Integer
Dim Data As String
    
'// 格納済みかどうかの判定
For C = 1 To 6 Step 1
    For R = 4 To 11 Step 1
        Data = Worksheets("座席").Cells(R, C).Value
            
        If Data Like "*" & ComboBox1.Value Then
            MsgBox "既に格納済みです。"
            Exit Sub
        End If
    Next
Next

Name = ComboBox1.Value

If CheckBox1.Value = True Then
    If OptionButton1.Value = True Then
        Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 3).Value & vbLf _
        & Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 1).Value & " " & ComboBox1.Value
    Else
        Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 1).Value & " " & ComboBox1.Value
    End If
Else
    If OptionButton1.Value = True Then
        Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 3).Value & vbLf & ComboBox1.Value
    End If
End If

ActiveCell.Value = Name

'//色
If Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 4) = "M" Then
    ActiveCell.Interior.ColorIndex = 20
Else
    ActiveCell.Interior.ColorIndex = 38
End If
  
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm
End Sub

Private Sub CommandButton3_Click()
    ActiveCell.ClearContents
    ActiveCell.Interior.ColorIndex = 2
End Sub

'// かな氏名格納
Private Sub OptionButton1_Click()

If CheckBox1.Value = True Then
    Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 3).Value & vbLf _
    & Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 1).Value & " " & ComboBox1.Value
Else
    Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 3).Value & vbLf & ComboBox1.Value
End If

ActiveCell.Value = Name

End Sub

'//　かな氏名格納しない
Private Sub OptionButton2_Click()

If CheckBox1.Value = True Then
    Name = Worksheets("学生一覧").Cells(ComboBox1.ListIndex + 2, 1).Value & " " & ComboBox1.Value
Else
    Name = ComboBox1.Value
End If
    ActiveCell = Name
    
End Sub

'// かな氏名
Private Sub SpinButton2_SpinDown()
Dim N As Long

'//かな氏名格納時のみフォントサイズ変更
If OptionButton1.Value = True Then
    FontSize2 = FontSize2 - 1

    N = InStr(ActiveCell, vbLf)

    ActiveCell.Characters(Length:=N).Font.Size = FontSize2
    TextBox2.Value = FontSize2

    If FontSize2 < 6 Then
        FontSize2 = 6
        ActiveCell.Characters(Length:=N).Font.Size = FontSize2
        TextBox2.Value = FontSize2
        MsgBox "これ以上小さくできません"
        Exit Sub
    End If
End If
End Sub

Private Sub SpinButton2_SpinUp()
Dim N As Long

If OptionButton1.Value = True Then
    FontSize2 = FontSize2 + 1

    N = InStr(ActiveCell, vbLf)

    ActiveCell.Characters(Length:=N).Font.Size = FontSize2
    TextBox2.Value = FontSize2

    If FontSize2 > 18 Then
        FontSize2 = 18
        ActiveCell.Characters(Length:=N).Font.Size = FontSize2
        TextBox2.Value = FontSize2
        MsgBox "これ以上大きくできません"
        Exit Sub
    End If
End If
End Sub

'//漢字番号

Private Sub SpinButton1_SpinDown()
Dim N As Long

FontSize1 = FontSize1 - 1

N = InStr(ActiveCell, vbLf)

ActiveCell.Characters(Start:=N + 1).Font.Size = FontSize1
TextBox1.Value = FontSize1

If FontSize1 < 6 Then
    FontSize1 = 6
    ActiveCell.Characters(Start:=N + 1).Font.Size = FontSize1
    TextBox1.Value = FontSize1
    MsgBox "これ以上小さくできません"
    Exit Sub
End If
End Sub

Private Sub SpinButton1_SpinUp()
Dim N As Long

FontSize1 = FontSize1 + 1

N = InStr(ActiveCell, vbLf)

ActiveCell.Characters(Start:=N + 1).Font.Size = FontSize1
TextBox1.Value = FontSize1

If FontSize1 > 18 Then
    FontSize1 = 18
    ActiveCell.Characters(Start:=N + 1).Font.Size = FontSize1
    TextBox1.Value = FontSize1
    MsgBox "これ以上大きくできません"
    Exit Sub
End If
End Sub
