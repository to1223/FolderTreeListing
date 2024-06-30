Attribute VB_Name = "Main"
Option Explicit

Private Type UserInput

    RootPath As String
    TriggerFolderName As String
    TriggerSearchLevel As Long
    FolderListLevel As Long
    IgnoreCase As Boolean

End Type


Public Sub Update()

On Error GoTo ErrorHandling

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject

    Dim target_sheet As Worksheet
    Set target_sheet = ThisWorkbook.ActiveSheet

    '' 入力値の取得と検証
    Dim user_input As UserInput
    
    With target_sheet
        user_input.RootPath = .Range("RootPath")
        user_input.TriggerFolderName = .Range("TriggerFolderName")
        user_input.TriggerSearchLevel = CLng(.Range("TriggerSearchLevel"))
        user_input.FolderListLevel = CLng(.Range("FolderListLevel"))
        
        If .Range("IgnoreCase") = "On" Then
            user_input.IgnoreCase = True
        Else
            user_input.IgnoreCase = False
        End If
    End With
    
    If Not fso.FolderExists(user_input.RootPath) Then
        Err.Raise ErrorTypeFolderNotFound
    End If
    
    '' データ作成
    Dim builder As DataBuilder
    Set builder = New DataBuilder
    
    builder.Init _
        user_input.RootPath, _
        user_input.TriggerFolderName, _
        user_input.TriggerSearchLevel, _
        user_input.FolderListLevel, _
        user_input.IgnoreCase
    
    builder.Build
    
    Dim data_array() As Variant
    data_array = TransposeData(builder)
    
    '' エクセルへの書き込み
    ' 既存のデータ削除
    target_sheet.Range("OutputHeaderPosition").CurrentRegion.Clear
    
    ' ヘッダの書き込み
    Dim i As Long
    
    With target_sheet
        For i = 1 To UBound(data_array, 2)
            .Range("OutputHeaderPosition").Offset(0, i - 1) = "Level" & CStr(i)
        Next
        
        .Range("OutputHeaderPosition").Offset(0, UBound(data_array, 2)) = "Link"
    End With
    
    ' ボディの書き込み
    target_sheet.Range( _
        target_sheet.Range("OutputHeaderPosition").Offset(1, 0), _
        target_sheet.Range("OutputHeaderPosition").Offset( _
            1 + UBound(data_array, 1), UBound(data_array, 2) _
        ) _
    ) = data_array

    ' 出力日時の書き込み
    target_sheet.Range("OutputDateTime").Value = Now()
    
    '' 書式設定
    With target_sheet.Range("OutputHeaderPosition").CurrentRegion
        .HorizontalAlignment = xlCenter
        .AutoFilter
    End With

    GoTo Closing

ErrorHandling:
    Select Case Err.Number
        ' クラスが発生させるエラーコードをクラス内部で定義した場合、ここ（プロシージャレベルのエラーハンドリング）で例外処理しようとすると、エラーコードの重複の可能性がある。
        ' そういう意味で、そのクラスを使用しているコード直下に例外処理を書いた方がいいと思われる。
        Case ErrorTypeFolderNotFound
            MsgBox "指定のフォルダが見つかりませんでした。入力が正しいか確認してください。"
            GoTo Closing
        Case Else
            MsgBox "エラーが発生しました。管理者にお知らせください。"
            GoTo Closing
    End Select

Closing:
    Set fso = Nothing
    Exit Sub

End Sub


''' 元配列を出力用に変換する。
''' source_row: (path, level1, level2, ...) -> result_col: (level1, level2, ... , link)
Private Function TransposeData(builder As DataBuilder) As Variant

    Dim source_array() As String
    Dim result_array() As Variant ' 型をStringにすると、数式表現でも文字列として扱われてしまう。
    
    source_array = builder.Data
    
    ReDim result_array(UBound(source_array, 2), UBound(source_array, 1))
    
    Dim source_row As Long
    Dim source_col As Long
    
    For source_col = LBound(source_array, 2) To UBound(source_array, 2)
        For source_row = LBound(source_array, 1) To UBound(source_array, 1)
            ' ハイパーリンクをつける
            If source_row = 0 Then
                If source_array(source_row, source_col) <> "" Then
                    result_array(source_col, UBound(source_array, 1)) = "=HYPERLINK(""" & source_array(source_row, source_col) & """, ""→"")"
                Else
                    result_array(source_col, UBound(source_array, 1)) = source_array(source_row, source_col)
                End If
            Else
                result_array(source_col, source_row - 1) = source_array(source_row, source_col)
            End If
        Next
    Next

    TransposeData = result_array

End Function
