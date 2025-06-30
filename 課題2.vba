Option Explicit
'インデントは普段書いてるプログラムのノリでやってます VBAですがオールマンスタイルです
'ライセンス:GPLv3

Sub ボタン1_Click()
	'配列作成
	'変数類
	Dim subjct_1a_name() As String  '区分用stirng型
	Dim string_buffer As String 'subject_1aで処理する文字列用
	Dim columns_counter As Long '行数計算用
	columns_counter = Worksheets("３月実績").ListObjects("テーブル1").ListRows.Count    '行数を代入
	Dim loop_counter As Long    'ループ回数識別用
	Dim var_counter As Long '配列カウントアップ用変数
	Dim bound_counter As Long   '配列確認ループ用
	Dim columns_counter_absolute As Long    '行絶対値指定用
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		string_buffer = Worksheets("３月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("３月実績").Cells(columns_counter_absolute, "E").Text 'D列+E列統合
		If (Not subjct_1a_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1a_name(0)    '配列拡張
			subjct_1a_name(0) = string_buffer   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1a_name) Then    '配列に一致データ無し
				For bound_counter = 0 To UBound(subjct_1a_name) '配列データの末尾まで検索
					If subjct_1a_name(bound_counter) = string_buffer Then   '保険
						Exit For
					Else
						ReDim Preserve subjct_1a_name(UBound(subjct_1a_name) + 1)   '配列拡張
						subjct_1a_name(UBound(subjct_1a_name)) = string_buffer  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1a_name(var_counter) = string_buffer Then '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
		If string_buffer = " - " Then
			subjct_1a_name(UBound(subjct_1a_name)) = "合計"
		End If
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1a_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		string_buffer = Worksheets("３月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("３月実績").Cells(columns_counter_absolute, "E").Text 'D列+E列統合
		If (Not subjct_1a_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1a_time(0)    '配列拡張
			subjct_1a_time(0) = Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_1  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1a_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1a_name(var_counter) = string_buffer Then '名前とセルの値が一致
				If UBound(subjct_1a_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1a_time(UBound(subjct_1a_time) + 1)   '配列拡張
				End If
				subjct_1a_time(var_counter) = subjct_1a_time(var_counter) + Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1a_name(var_counter) <> string_buffer And var_counter = UBound(subjct_1a_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1a_time(UBound(subjct_1a_time) + 1)   '配列拡張
				subjct_1a_time(UBound(subjct_1a_time)) = subjct_1a_time(UBound(subjct_1a_time)) + Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_1:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数類
	Dim rtn As Integer
	Dim tbl As ListObject
	'変数リセット
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "B").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(B7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("B" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("B7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("B6:C" & columns_counter_absolute + UBound(subjct_1a_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_1
		End Select
	End If
	ReDim Preserve subjct_1a_name(UBound(subjct_1a_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "B").Value = "分類・業務区分別勤務時間"
	Cells(columns_counter_absolute, "C").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1a_name)
		Cells(columns_counter_absolute, "B").Value = subjct_1a_name(loop_counter)
		Cells(columns_counter_absolute, "C").Value = subjct_1a_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("B7").CurrentRegion
	Set tbl = Range("B7").ListObject
	tbl.ListColumns(1).Name = "分類・業務区分別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.Name = "課題2_a"
	tbl.ShowTotals = True
	'ソート
	tbl.Range.Sort key1:=Range("B7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final_1:

	'配列作成
	'変数類
	Dim subjct_1b_name() As String  '区分用stirng型
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		string_buffer = Worksheets("３月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("３月実績").Cells(columns_counter_absolute, "E").Text & " - " & Worksheets("３月実績").Cells(columns_counter_absolute, "F").Text  'D列+E列+F列統合
		If (Not subjct_1b_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1b_name(0)    '配列拡張
			subjct_1b_name(0) = string_buffer   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1b_name) Then    '配列に一致データ無し
				For bound_counter = 0 To UBound(subjct_1b_name) '配列データの末尾まで検索
					If subjct_1b_name(bound_counter) = string_buffer Then   '保険
						Exit For
					Else
						ReDim Preserve subjct_1b_name(UBound(subjct_1b_name) + 1)   '配列拡張
						subjct_1b_name(UBound(subjct_1b_name)) = string_buffer  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1b_name(var_counter) = string_buffer Then '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
		If string_buffer = " - " Then
			subjct_1b_name(UBound(subjct_1b_name)) = "合計"
		End If
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1b_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		string_buffer = Worksheets("３月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("３月実績").Cells(columns_counter_absolute, "E").Text & " - " & Worksheets("３月実績").Cells(columns_counter_absolute, "F").Text  'D列+E列+F列統合
		If (Not subjct_1b_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1b_time(0)    '配列拡張
			subjct_1b_time(0) = Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_2  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1b_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1b_name(var_counter) = string_buffer Then '名前とセルの値が一致
				If UBound(subjct_1b_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1b_time(UBound(subjct_1b_time) + 1)   '配列拡張
				End If
				subjct_1b_time(var_counter) = subjct_1b_time(var_counter) + Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value  '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1b_name(var_counter) <> string_buffer And var_counter = UBound(subjct_1b_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1b_time(UBound(subjct_1b_time) + 1)   '配列拡張
				subjct_1b_time(UBound(subjct_1b_time)) = subjct_1b_time(UBound(subjct_1b_time)) + Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value    '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_2:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数リセット
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "E").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(E7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("E" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("E7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("E7:F" & columns_counter_absolute + UBound(subjct_1b_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_2
		End Select
	End If
	ReDim Preserve subjct_1b_name(UBound(subjct_1b_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "E").Value = "分類・業務区分別勤務時間"
	Cells(columns_counter_absolute, "F").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1b_name)
		Cells(columns_counter_absolute, "E").Value = subjct_1b_name(loop_counter)
		Cells(columns_counter_absolute, "F").Value = subjct_1b_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("E7").CurrentRegion
	Set tbl = Range("E7").ListObject
	tbl.ListColumns(1).Name = "分類・業務・工程区分別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.ShowTotals = True
	tbl.Name = "課題2_b"
	'ソート
	tbl.Range.Sort key1:=Range("E7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final_2:

	'配列作成
	'変数類
	Dim subjct_1c_name() As String  '区分用stirng型
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		If (Not subjct_1c_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1c_name(0)    '配列拡張
			subjct_1c_name(0) = Worksheets("３月実績").Cells(columns_counter_absolute, "B").Text   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1c_name) Then    '保険
				For bound_counter = 0 To UBound(subjct_1c_name)
					If subjct_1c_name(bound_counter) = Worksheets("３月実績").Cells(columns_counter_absolute, "B").Text Then
						Exit For
					Else
						ReDim Preserve subjct_1c_name(UBound(subjct_1c_name) + 1)   '配列拡張
						subjct_1c_name(UBound(subjct_1c_name)) = Worksheets("３月実績").Cells(columns_counter_absolute, "B")  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1c_name(var_counter) = Worksheets("３月実績").Cells(columns_counter_absolute, "B") Then   '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1c_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		If (Not subjct_1c_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1c_time(0)    '配列拡張
			subjct_1c_time(0) = Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_3    '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1c_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1c_name(var_counter) = Worksheets("３月実績").Cells(columns_counter_absolute, "B").Value Then '名前とセルの値が一致
				If UBound(subjct_1c_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1c_time(UBound(subjct_1c_time) + 1)   '配列拡張
				End If
				subjct_1c_time(var_counter) = subjct_1c_time(var_counter) + Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1c_name(var_counter) <> Worksheets("３月実績").Cells(columns_counter_absolute, "B").Text And var_counter = UBound(subjct_1c_time) Then    '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1c_time(UBound(subjct_1c_time) + 1)   '配列拡張
				subjct_1c_time(UBound(subjct_1c_time)) = subjct_1c_time(UBound(subjct_1c_time)) + Worksheets("３月実績").Cells(columns_counter_absolute, "I").Value '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_3:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "H").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(H7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("H" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("H7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("H7:I" & columns_counter_absolute + UBound(subjct_1c_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_3
		End Select
	End If
	ReDim Preserve subjct_1c_name(UBound(subjct_1c_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "H").Value = "社員別勤務時間"
	Cells(columns_counter_absolute, "I").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1c_name)
		Cells(columns_counter_absolute, "H").Value = subjct_1c_name(loop_counter)
		Cells(columns_counter_absolute, "I").Value = subjct_1c_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("H7").CurrentRegion
	Set tbl = Range("H7").ListObject
	tbl.ListColumns(1).Name = "社員別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.ShowTotals = True
	tbl.Name = "課題2_c"
	'ソート
	tbl.Range.Sort key1:=Range("H7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	Range("B6").Value = "3月分統計"
	final_3:
End Sub

Sub ボタン2_Click()
	'配列作成
	'変数類
	Dim subjct_1a_name() As String  '区分用stirng型
	Dim string_buffer As String 'subject_1aで処理する文字列用
	Dim columns_counter As Long '行数計算用
	columns_counter = Worksheets("４月実績").ListObjects("テーブル2").ListRows.Count    '行数を代入
	Dim loop_counter As Long    'ループ回数識別用
	Dim var_counter As Long '配列カウントアップ用変数
	Dim bound_counter As Long   '配列確認ループ用
	Dim columns_counter_absolute As Long    '行絶対値指定用
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		string_buffer = Worksheets("４月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("４月実績").Cells(columns_counter_absolute, "E").Text 'D列+E列統合
		If (Not subjct_1a_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1a_name(0)    '配列拡張
			subjct_1a_name(0) = string_buffer   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1a_name) Then    '保険
				For bound_counter = 0 To UBound(subjct_1a_name)
					If subjct_1a_name(bound_counter) = string_buffer Then
						Exit For
					Else
						ReDim Preserve subjct_1a_name(UBound(subjct_1a_name) + 1)   '配列拡張
						subjct_1a_name(UBound(subjct_1a_name)) = string_buffer  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1a_name(var_counter) = string_buffer Then '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
		If string_buffer = " - " Then
			subjct_1a_name(UBound(subjct_1a_name)) = "合計"
		End If
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1a_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		string_buffer = Worksheets("４月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("４月実績").Cells(columns_counter_absolute, "E").Text 'D列+E列統合
		If (Not subjct_1a_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1a_time(0)    '配列拡張
			subjct_1a_time(0) = Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_1  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1a_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1a_name(var_counter) = string_buffer Then '名前とセルの値が一致
				If UBound(subjct_1a_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1a_time(UBound(subjct_1a_time) + 1)   '配列拡張
				End If
				subjct_1a_time(var_counter) = subjct_1a_time(var_counter) + Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1a_name(var_counter) <> string_buffer And var_counter = UBound(subjct_1a_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1a_time(UBound(subjct_1a_time) + 1)   '配列拡張
				subjct_1a_time(UBound(subjct_1a_time)) = subjct_1a_time(UBound(subjct_1a_time)) + Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_1:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数類
	Dim rtn As Integer
	Dim tbl As ListObject
	'変数リセット
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "K").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(K7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("K" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("K7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("K6:L" & columns_counter_absolute + UBound(subjct_1a_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_1
		End Select
	End If
	ReDim Preserve subjct_1a_name(UBound(subjct_1a_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "K").Value = "分類・業務区分別勤務時間"
	Cells(columns_counter_absolute, "L").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1a_name)
		Cells(columns_counter_absolute, "K").Value = subjct_1a_name(loop_counter)
		Cells(columns_counter_absolute, "L").Value = subjct_1a_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("K7").CurrentRegion
	Set tbl = Range("K7").ListObject
	tbl.ListColumns(1).Name = "分類・業務区分別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.Name = "課題2_d"
	tbl.ShowTotals = True
	'ソート
	tbl.Range.Sort key1:=Range("K7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final_1:

	'配列作成
	'変数類
	Dim subjct_1b_name() As String  '区分用stirng型
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		string_buffer = Worksheets("４月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("４月実績").Cells(columns_counter_absolute, "E").Text & " - " & Worksheets("４月実績").Cells(columns_counter_absolute, "F").Text  'D列+E列+F列統合
		If (Not subjct_1b_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1b_name(0)    '配列拡張
			subjct_1b_name(0) = string_buffer   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1b_name) Then    '保険
				For bound_counter = 0 To UBound(subjct_1b_name)
					If subjct_1b_name(bound_counter) = string_buffer Then
						Exit For
					Else
						ReDim Preserve subjct_1b_name(UBound(subjct_1b_name) + 1)   '配列拡張
						subjct_1b_name(UBound(subjct_1b_name)) = string_buffer  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1b_name(var_counter) = string_buffer Then '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
		If string_buffer = " - " Then
			subjct_1b_name(UBound(subjct_1b_name)) = "合計"
		End If
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1b_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		string_buffer = Worksheets("４月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("４月実績").Cells(columns_counter_absolute, "E").Text & " - " & Worksheets("４月実績").Cells(columns_counter_absolute, "F").Text  'D列+E列+F列統合
		If (Not subjct_1b_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1b_time(0)    '配列拡張
			subjct_1b_time(0) = Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_2  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1b_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1b_name(var_counter) = string_buffer Then '名前とセルの値が一致
				If UBound(subjct_1b_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1b_time(UBound(subjct_1b_time) + 1)   '配列拡張
				End If
				subjct_1b_time(var_counter) = subjct_1b_time(var_counter) + Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value  '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1b_name(var_counter) <> string_buffer And var_counter = UBound(subjct_1b_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1b_time(UBound(subjct_1b_time) + 1)   '配列拡張
				subjct_1b_time(UBound(subjct_1b_time)) = subjct_1b_time(UBound(subjct_1b_time)) + Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value    '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_2:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数リセット
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "N").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(N7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("N" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("N7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("N7:O" & columns_counter_absolute + UBound(subjct_1b_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_2
		End Select
	End If
	ReDim Preserve subjct_1b_name(UBound(subjct_1b_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "N").Value = "分類・業務区分別勤務時間"
	Cells(columns_counter_absolute, "O").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1b_name)
		Cells(columns_counter_absolute, "N").Value = subjct_1b_name(loop_counter)
		Cells(columns_counter_absolute, "O").Value = subjct_1b_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("N7").CurrentRegion
	Set tbl = Range("O7").ListObject
	tbl.ListColumns(1).Name = "分類・業務・工程区分別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.ShowTotals = True
	tbl.Name = "課題2_e"
	'ソート
	tbl.Range.Sort key1:=Range("N7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final_2:

	'配列作成
	'変数類
	Dim subjct_1c_name() As String  '区分用stirng型
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		If (Not subjct_1c_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1c_name(0)    '配列拡張
			subjct_1c_name(0) = Worksheets("４月実績").Cells(columns_counter_absolute, "B").Text   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1c_name) Then    '保険
				For bound_counter = 0 To UBound(subjct_1c_name)
					If subjct_1c_name(bound_counter) = Worksheets("４月実績").Cells(columns_counter_absolute, "B").Text Then
						Exit For
					Else
						ReDim Preserve subjct_1c_name(UBound(subjct_1c_name) + 1)   '配列拡張
						subjct_1c_name(UBound(subjct_1c_name)) = Worksheets("４月実績").Cells(columns_counter_absolute, "B")  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1c_name(var_counter) = Worksheets("４月実績").Cells(columns_counter_absolute, "B") Then   '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1c_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		If (Not subjct_1c_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1c_time(0)    '配列拡張
			subjct_1c_time(0) = Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_3    '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1c_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1c_name(var_counter) = Worksheets("４月実績").Cells(columns_counter_absolute, "B").Value Then '名前とセルの値が一致
				If UBound(subjct_1c_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1c_time(UBound(subjct_1c_time) + 1)   '配列拡張
				End If
				subjct_1c_time(var_counter) = subjct_1c_time(var_counter) + Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1c_name(var_counter) <> Worksheets("４月実績").Cells(columns_counter_absolute, "B").Text And var_counter = UBound(subjct_1c_time) Then    '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1c_time(UBound(subjct_1c_time) + 1)   '配列拡張
				subjct_1c_time(UBound(subjct_1c_time)) = subjct_1c_time(UBound(subjct_1c_time)) + Worksheets("４月実績").Cells(columns_counter_absolute, "I").Value '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_3:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "Q").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(Q7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("Q" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("Q7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("Q7:R" & columns_counter_absolute + UBound(subjct_1c_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_3
		End Select
	End If
	ReDim Preserve subjct_1c_name(UBound(subjct_1c_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "Q").Value = "社員別勤務時間"
	Cells(columns_counter_absolute, "R").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1c_name)
		Cells(columns_counter_absolute, "Q").Value = subjct_1c_name(loop_counter)
		Cells(columns_counter_absolute, "R").Value = subjct_1c_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("Q7").CurrentRegion
	Set tbl = Range("Q7").ListObject
	tbl.ListColumns(1).Name = "社員別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.ShowTotals = True
	tbl.Name = "課題2_f"
	'ソート
	tbl.Range.Sort key1:=Range("Q7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final_3:

	Range("K6").Value = "4月分統計"
End Sub

Sub ボタン3_Click()
	'配列作成
	'変数類
	Dim subjct_1a_name() As String  '区分用stirng型
	Dim string_buffer As String 'subject_1aで処理する文字列用
	Dim columns_counter As Long '行数計算用
	columns_counter = Worksheets("５月実績").ListObjects("テーブル3").ListRows.Count    '行数を代入
	Dim loop_counter As Long    'ループ回数識別用
	Dim var_counter As Long '配列カウントアップ用変数
	Dim bound_counter As Long   '配列確認ループ用
	Dim columns_counter_absolute As Long    '行絶対値指定用
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		string_buffer = Worksheets("５月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("５月実績").Cells(columns_counter_absolute, "E").Text 'D列+E列統合
		If (Not subjct_1a_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1a_name(0)    '配列拡張
			subjct_1a_name(0) = string_buffer   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1a_name) Then    '保険
				For bound_counter = 0 To UBound(subjct_1a_name)
					If subjct_1a_name(bound_counter) = string_buffer Then
						Exit For
					Else
						ReDim Preserve subjct_1a_name(UBound(subjct_1a_name) + 1)   '配列拡張
						subjct_1a_name(UBound(subjct_1a_name)) = string_buffer  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1a_name(var_counter) = string_buffer Then '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
		If string_buffer = " - " Then
			subjct_1a_name(UBound(subjct_1a_name)) = "合計"
		End If
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1a_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		string_buffer = Worksheets("５月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("５月実績").Cells(columns_counter_absolute, "E").Text 'D列+E列統合
		If (Not subjct_1a_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1a_time(0)    '配列拡張
			subjct_1a_time(0) = Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_1  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1a_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1a_name(var_counter) = string_buffer Then '名前とセルの値が一致
				If UBound(subjct_1a_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1a_time(UBound(subjct_1a_time) + 1)   '配列拡張
				End If
				subjct_1a_time(var_counter) = subjct_1a_time(var_counter) + Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1a_name(var_counter) <> string_buffer And var_counter = UBound(subjct_1a_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1a_time(UBound(subjct_1a_time) + 1)   '配列拡張
				subjct_1a_time(UBound(subjct_1a_time)) = subjct_1a_time(UBound(subjct_1a_time)) + Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_1:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数類
	Dim rtn As Integer
	Dim tbl As ListObject
	'変数リセット
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "T").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(T7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("T" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("T7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("T6:U" & columns_counter_absolute + UBound(subjct_1a_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_1
		End Select
	End If
	ReDim Preserve subjct_1a_name(UBound(subjct_1a_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "T").Value = "分類・業務区分別勤務時間"
	Cells(columns_counter_absolute, "U").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1a_name)
		Cells(columns_counter_absolute, "T").Value = subjct_1a_name(loop_counter)
		Cells(columns_counter_absolute, "U").Value = subjct_1a_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("T7").CurrentRegion
	Set tbl = Range("T7").ListObject
	tbl.ListColumns(1).Name = "分類・業務区分別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.Name = "課題2_g"
	tbl.ShowTotals = True
	'ソート
	tbl.Range.Sort key1:=Range("T7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final_1:

	'配列作成
	'変数類
	Dim subjct_1b_name() As String  '区分用stirng型
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		string_buffer = Worksheets("５月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("５月実績").Cells(columns_counter_absolute, "E").Text & " - " & Worksheets("５月実績").Cells(columns_counter_absolute, "F").Text  'D列+E列+F列統合
		If (Not subjct_1b_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1b_name(0)    '配列拡張
			subjct_1b_name(0) = string_buffer   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1b_name) Then    '保険
				For bound_counter = 0 To UBound(subjct_1b_name)
					If subjct_1b_name(bound_counter) = string_buffer Then
						Exit For
					Else
						ReDim Preserve subjct_1b_name(UBound(subjct_1b_name) + 1)   '配列拡張
						subjct_1b_name(UBound(subjct_1b_name)) = string_buffer  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1b_name(var_counter) = string_buffer Then '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
		If string_buffer = " - " Then
			subjct_1b_name(UBound(subjct_1b_name)) = "合計"
		End If
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1b_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		string_buffer = Worksheets("５月実績").Cells(columns_counter_absolute, "D").Text & " - " & Worksheets("５月実績").Cells(columns_counter_absolute, "E").Text & " - " & Worksheets("５月実績").Cells(columns_counter_absolute, "F").Text  'D列+E列+F列統合
		If (Not subjct_1b_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1b_time(0)    '配列拡張
			subjct_1b_time(0) = Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_2  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1b_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1b_name(var_counter) = string_buffer Then '名前とセルの値が一致
				If UBound(subjct_1b_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1b_time(UBound(subjct_1b_time) + 1)   '配列拡張
				End If
				subjct_1b_time(var_counter) = subjct_1b_time(var_counter) + Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value  '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1b_name(var_counter) <> string_buffer And var_counter = UBound(subjct_1b_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1b_time(UBound(subjct_1b_time) + 1)   '配列拡張
				subjct_1b_time(UBound(subjct_1b_time)) = subjct_1b_time(UBound(subjct_1b_time)) + Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value    '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_2:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数リセット
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "W").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(W7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("W" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("W7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("W7:X" & columns_counter_absolute + UBound(subjct_1b_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_2
		End Select
	End If
	ReDim Preserve subjct_1b_name(UBound(subjct_1b_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "W").Value = "分類・業務区分別勤務時間"
	Cells(columns_counter_absolute, "X").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1b_name)
		Cells(columns_counter_absolute, "W").Value = subjct_1b_name(loop_counter)
		Cells(columns_counter_absolute, "X").Value = subjct_1b_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("W7").CurrentRegion
	Set tbl = Range("W7").ListObject
	tbl.ListColumns(1).Name = "分類・業務・工程区分別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.ShowTotals = True
	tbl.Name = "課題2_h"
	'ソート
	tbl.Range.Sort key1:=Range("W7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final_2:

	'配列作成
	'変数類
	Dim subjct_1c_name() As String  '区分用stirng型
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		If (Not subjct_1c_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1c_name(0)    '配列拡張
			subjct_1c_name(0) = Worksheets("５月実績").Cells(columns_counter_absolute, "B").Text   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1c_name) Then    '保険
				For bound_counter = 0 To UBound(subjct_1c_name)
					If subjct_1c_name(bound_counter) = Worksheets("５月実績").Cells(columns_counter_absolute, "B").Text Then
						Exit For
					Else
						ReDim Preserve subjct_1c_name(UBound(subjct_1c_name) + 1)   '配列拡張
						subjct_1c_name(UBound(subjct_1c_name)) = Worksheets("５月実績").Cells(columns_counter_absolute, "B")  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1c_name(var_counter) = Worksheets("５月実績").Cells(columns_counter_absolute, "B") Then   '配列に一致するデータが存在しないか
				Exit For
			End If
		Next var_counter
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter   'ループここまで

	'作業時間加算
	'変数類
	Dim subjct_1c_time() As Long    '時間カウント用
	'変数リセット
	columns_counter_absolute = 3

	'処理部分
	For loop_counter = 0 To columns_counter
		If (Not subjct_1c_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1c_time(0)    '配列拡張
			subjct_1c_time(0) = Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time_3    '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1c_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1c_name(var_counter) = Worksheets("５月実績").Cells(columns_counter_absolute, "B").Value Then '名前とセルの値が一致
				If UBound(subjct_1c_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1c_time(UBound(subjct_1c_time) + 1)   '配列拡張
				End If
				subjct_1c_time(var_counter) = subjct_1c_time(var_counter) + Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value   '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1c_name(var_counter) <> Worksheets("５月実績").Cells(columns_counter_absolute, "B").Text And var_counter = UBound(subjct_1c_time) Then    '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1c_time(UBound(subjct_1c_time) + 1)   '配列拡張
				subjct_1c_time(UBound(subjct_1c_time)) = subjct_1c_time(UBound(subjct_1c_time)) + Worksheets("５月実績").Cells(columns_counter_absolute, "I").Value '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time_3:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	columns_counter_absolute = 7

	'処理部分
	If Cells(7, "Z").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(Z7)にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveSheet.Range("Z" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = Range("Z7").ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("Z7:AA" & columns_counter_absolute + UBound(subjct_1c_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final_3
		End Select
	End If
	ReDim Preserve subjct_1c_name(UBound(subjct_1c_name) - 1)   '合計行だけ消去
	Cells(columns_counter_absolute, "Z").Value = "社員別勤務時間"
	Cells(columns_counter_absolute, "AA").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1c_name)
		Cells(columns_counter_absolute, "Z").Value = subjct_1c_name(loop_counter)
		Cells(columns_counter_absolute, "AA").Value = subjct_1c_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveSheet.ListObjects.Add 1, Range("Z7").CurrentRegion
	Set tbl = Range("Z7").ListObject
	tbl.ListColumns(1).Name = "社員別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.ShowTotals = True
	tbl.Name = "課題2_i"
	'ソート
	tbl.Range.Sort key1:=Range("Z7"), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	Range("T6").Value = "5月分統計"
	final_3:
End Sub