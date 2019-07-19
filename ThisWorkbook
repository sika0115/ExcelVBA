Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim セルアドレス As String
    Dim シート名 As String
    Dim R As Integer
    Dim C As Integer
    
    
    '// このイベント処理中に発生する他のイベントを無効にする
    Application.EnableEvents = False
    '// イベント処理
    セルアドレス = Target.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
    シート名 = Sh.Name
     
    R = Target.Row
    C = Target.Column
    
    If シート名 = "座席" Then
        If 3 < R And R < 12 And 0 < C And C < 7 Then
            UserForm.Show
        End If
    End If
    '// Excel本体のマウス右ボタンクリック時の機能を実行させない
    Cancel = True

    
    '// これ以降のイベントを有効にする
    Application.EnableEvents = True
    
    
End Sub
