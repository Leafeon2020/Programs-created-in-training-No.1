Option Explicit
'インデントは普段書いてるプログラムのノリでやってます VBAですがオールマンスタイルです
'ライセンス:GPLv3

Sub ボタン1_Click() '課題1-a
	'配列作成
	'変数類
	Dim subjct_1a_name() As String  '区分用stirng型
	Dim string_buffer As String 'subject_1aで処理する文字列用
	Dim columns_counter As Long '行数計算用
	columns_counter = ThisWorkBook.Worksheets(1).ListObjects("テーブル1").ListRows.Count   '行数を代入
	Dim loop_counter As Long    'ループ回数識別用
	Dim var_counter As Long '配列カウントアップ用変数
	Dim bound_counter As Long   '配列確認ループ用
	Dim columns_counter_absolute As Long    '行絶対値指定用
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		string_buffer = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "D").Text & " - " & ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "E").Text   'D列+E列統合
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
		string_buffer = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "D").Text & " - " & Cells(columns_counter_absolute, "E").Text   'D列+E列統合
		If (Not subjct_1a_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1a_time(0)    '配列拡張
			subjct_1a_time(0) = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "I").Value  '時間を加算
			GoTo skip_time  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1a_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1a_name(var_counter) = string_buffer Then '名前とセルの値が一致
				If UBound(subjct_1a_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1a_time(UBound(subjct_1a_time) + 1)   '配列拡張
				End If
				subjct_1a_time(var_counter) = subjct_1a_time(var_counter) + ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "I").Value  '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1a_name(var_counter) <> string_buffer And var_counter = UBound(subjct_1a_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1a_time(UBound(subjct_1a_time) + 1)   '配列拡張
				subjct_1a_time(UBound(subjct_1a_time)) = subjct_1a_time(UBound(subjct_1a_time)) + ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "I").Value    '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数類
	Dim rtn As Integer
	Dim tbl As ListObject
	'変数リセット
	columns_counter_absolute = columns_counter + 5

	'処理部分
	If ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "K").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(K" & columns_counter_absolute & ")にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveWorkBook.Worksheets(1).Range("K" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = ActiveWorkBook.Worksheets(1).Range("K" & columns_counter_absolute).ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				Range("K" & columns_counter_absolute & ":L" & columns_counter_absolute + UBound(subjct_1a_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final
		End Select
	End If
	ReDim Preserve subjct_1a_name(UBound(subjct_1a_name) - 1)   '合計行だけ消去
	ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "K").Value = "分類・業務区分別勤務時間"
	ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "L").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1a_name)
		ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "K").Value = subjct_1a_name(loop_counter)
		ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "L").Value = subjct_1a_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveWorkBook.Worksheets(1).ListObjects.Add 1, ActiveWorkBook.Worksheets(1).Range("K" & columns_counter + 5).CurrentRegion
	Set tbl = Range("K" & columns_counter_absolute).ListObject
	tbl.ListColumns(1).Name = "分類・業務区分別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.Name = "課題1_a"
	tbl.ShowTotals = True
	'ソート
	tbl.Range.Sort key1: = ActiveWorkBook.Worksheets(1).Range("K" & columns_counter + 5), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin

	final:
End Sub

Sub ボタン2_Click() '課題1-b
	'配列作成
	'変数類
	Dim subjct_1b_name() As String  '区分用stirng型
	Dim string_buffer As String 'subject_1bで処理する文字列用
	Dim columns_counter As Long '行数計算用
	columns_counter = ThisWorkBook.Worksheets(1).ListObjects("テーブル1").ListRows.Count   '行数を代入
	Dim loop_counter As Long    'ループ回数識別用
	Dim var_counter As Long '配列カウントアップ用変数
	Dim bound_counter As Long   '配列確認ループ用
	Dim columns_counter_absolute As Long    '行絶対値指定用
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		string_buffer = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "D").Text & " - " & ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "E").Text & " - " & ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "F").Text   'D列+E列+F列統合
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
		string_buffer = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "D").Text & " - " & ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "E").Text & " - " & ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "F").Text   'D列+E列+F列統合
		If (Not subjct_1b_time) = -1 Then   '初回分岐
			ReDim Preserve subjct_1b_time(0)    '配列拡張
			subjct_1b_time(0) = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "I").Value   '時間を加算
			GoTo skip_time  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1b_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1b_name(var_counter) = string_buffer Then '名前とセルの値が一致
				If UBound(subjct_1b_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1b_time(UBound(subjct_1b_time) + 1)   '配列拡張
				End If
				subjct_1b_time(var_counter) = subjct_1b_time(var_counter) + ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "I").Value  '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1b_name(var_counter) <> string_buffer And var_counter = UBound(subjct_1b_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1b_time(UBound(subjct_1b_time) + 1)   '配列拡張
				subjct_1b_time(UBound(subjct_1b_time)) = subjct_1b_time(UBound(subjct_1b_time)) + ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "I").Value    '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数類
	Dim rtn As Integer
	Dim tbl As ListObject
	'変数リセット
	columns_counter_absolute = columns_counter + 5

	'処理部分
	If ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "N").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(N" & columns_counter_absolute & ")にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveWorkBook.Worksheets(1).Range("N" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = ActiveWorkBook.Worksheets(1).Range("N" & columns_counter_absolute).ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				ActiveWorkBook.Worksheets(1).Range("N" & columns_counter_absolute & ":O" & columns_counter_absolute + UBound(subjct_1b_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final
		End Select
	End If
	ReDim Preserve subjct_1b_name(UBound(subjct_1b_name) - 1)   '合計行だけ消去
	ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "N").Value = "分類・業務区分別勤務時間"
	ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "O").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1b_name)
		ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "N").Value = subjct_1b_name(loop_counter)
		ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "O").Value = subjct_1b_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveWorkBook.Worksheets(1).ListObjects.Add 1, ActiveWorkBook.Worksheets(1).Range("N" & columns_counter + 5).CurrentRegion
	Set tbl = Range("N" & columns_counter_absolute).ListObject
	tbl.ListColumns(1).Name = "分類・業務・工程区分別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.ShowTotals = True
	tbl.Name = "課題1_b"
	'ソート
	tbl.Range.Sort key1: = ActiveWorkBook.Worksheets(1).Range("N" & columns_counter + 5), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final:
End Sub

Sub ボタン3_Click()
	'配列作成
	'変数類
	Dim subjct_1c_name() As String  '区分用stirng型
	Dim columns_counter As Long '行数計算用
	columns_counter = ThisWorkBook.Worksheets(1).ListObjects("テーブル1").ListRows.Count   '行数を代入
	Dim loop_counter As Long    'ループ回数識別用
	Dim var_counter As Long '配列カウントアップ用変数
	Dim bound_counter As Long   '配列確認ループ用
	Dim columns_counter_absolute As Long    '行絶対値指定用
	columns_counter_absolute = 3    '初期値指定

	'処理部分
	For loop_counter = 0 To columns_counter '行数までループ
		If (Not subjct_1c_name) = -1 Then   '初回条件分岐
			ReDim Preserve subjct_1c_name(0)    '配列拡張
			subjct_1c_name(0) = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "B").Text   '変数代入
		End If
		For var_counter = 0 To columns_counter
			If var_counter > UBound(subjct_1c_name) Then    '配列に一致データ無し
				For bound_counter = 0 To UBound(subjct_1c_name) '配列データの末尾まで検索
					If subjct_1c_name(bound_counter) = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "B").Text Then    '保険
						Exit For
					Else
						ReDim Preserve subjct_1c_name(UBound(subjct_1c_name) + 1)   '配列拡張
						subjct_1c_name(UBound(subjct_1c_name)) = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "B")  '変数代入
						Exit For
					End If
				Next bound_counter
			End If
			If subjct_1c_name(var_counter) = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "B") Then '配列に一致するデータが存在しないか
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
			subjct_1c_time(0) = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "I").Value  '時間を加算
			GoTo skip_time  '後の処理全部すっ飛ばしてループの終了処理
		End If
		For var_counter = 0 To UBound(subjct_1c_name)   '名前テーブルを走査
			If var_counter > columns_counter Then   '例外対策
				Exit For
			ElseIf subjct_1c_name(var_counter) = ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "B").Value Then '名前とセルの値が一致
				If UBound(subjct_1c_time) < var_counter Then    '配列の数が配列カウンターを上回ったら例外処理
					ReDim Preserve subjct_1c_time(UBound(subjct_1c_time) + 1)   '配列拡張
				End If
				subjct_1c_time(var_counter) = subjct_1c_time(var_counter) + ThisWorkBook.Worksheets(1).Cells(columns_counter_absolute, "I").Value  '時間を加算
				Exit For    'ループ終了
			ElseIf subjct_1c_name(var_counter) <> Cells(columns_counter_absolute, "B").Text And var_counter = UBound(subjct_1c_time) Then   '時間テーブルに合致するパターン無し
				ReDim Preserve subjct_1c_time(UBound(subjct_1c_time) + 1)   '配列拡張
				subjct_1c_time(UBound(subjct_1c_time)) = subjct_1c_time(UBound(subjct_1c_time)) + Cells(columns_counter_absolute, "I").Value    '時間を加算
				Exit For    'ループ終了
			End If
		Next var_counter    '走査終わってなかったらもう1周
		skip_time:          '初回分岐用
		columns_counter_absolute = columns_counter_absolute + 1 '行数インクリメント
	Next loop_counter

	'出力
	'変数類
	Dim rtn As Integer
	Dim tbl As ListObject
	'変数リセット
	columns_counter_absolute = columns_counter + 5

	'処理部分
	If ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "Q").Value <> "" Then
		rtn = MsgBox("データ記入予定のセル(Q" & columns_counter_absolute & ")にデータが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				If Not ActiveWorkBook.Worksheets(1).Range("Q" & columns_counter_absolute).ListObject Is Nothing Then
					Set tbl = ActiveWorkBook.Worksheets(1).Range("Q" & columns_counter_absolute).ListObject
					tbl.TableStyle = ""
					tbl.Delete
				End If
				ActiveWorkBook.Worksheets(1).Range("Q" & columns_counter_absolute & ":R" & columns_counter_absolute + UBound(subjct_1c_name)).Value = ""
			Case vbNo   '処理中断
				GoTo final
		End Select
	End If
	ReDim Preserve subjct_1c_name(UBound(subjct_1c_name) - 1)   '合計行だけ消去
	ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "Q").Value = "社員別勤務時間"
	ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "R").Value = "勤務時間(分)"
	'数値入力
	For loop_counter = 0 To UBound(subjct_1c_name)
		ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "Q").Value = subjct_1c_name(loop_counter)
		ActiveWorkBook.Worksheets(1).Cells(columns_counter_absolute, "R").Value = subjct_1c_time(loop_counter)
		columns_counter_absolute = columns_counter_absolute + 1
	Next
	'テーブル作成
	ActiveWorkBook.Worksheets(1).ListObjects.Add 1, ActiveWorkBook.Worksheets(1).Range("Q" & columns_counter + 5).CurrentRegion
	Set tbl = ActiveWorkBook.Worksheets(1).Range("Q" & columns_counter_absolute).ListObject
	tbl.ListColumns(1).Name = "社員別勤務時間"
	tbl.ListColumns(2).Name = "勤務時間(分)"
	tbl.ShowTotals = True
	tbl.Name = "課題1_c"
	'ソート
	tbl.Range.Sort key1: = ActiveWorkBook.Worksheets(1).Range("Q" & columns_counter + 5), _
		order1:=xlAscending, _
		Header:=xlYes, _
		Orientation:=xlTopToBottom, _
		SortMethod:=xlPinYin
	final:
End Sub

Sub Reset()
	On Error Resume Next    '究極のゴリ押しを実現する魔法の呪文 C#で言うなら例外処理を処理に組み込んでます
	Dim tbl As ListObject
	Set tbl = ThisWorkBook.Worksheets(1).Range("課題1_a").ListObject  'エラーが出ようが出まいが変数に代入する
	If Err.Number = 0 Then  '正常に処理された(=テーブルが存在した)時の条件分岐
		tbl.TableStyle = ""
		tbl.Delete
		Err.Clear
	ElseIf Err.Number = 1004 Then   'テーブルが無かったら1004吐くのでそれを検出する
		Err.Clear
	Else    'それ以外
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Set tbl = ThisWorkBook.Worksheets(1).Range("課題1_b").ListObject  'エラーが出ようが出まいが変数に代入する
	If Err.Number = 0 Then  '正常に処理された(=テーブルが存在した)時の条件分岐
		tbl.TableStyle = ""
		tbl.Delete
		Err.Clear
	ElseIf Err.Number = 1004 Then   'テーブルが無かったら1004吐くのでそれを検出する
		Err.Clear
	Else    'それ以外
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Set tbl = ThisWorkBook.Worksheets(1).Range("課題1_c").ListObject  'エラーが出ようが出まいが変数に代入する
	If Err.Number = 0 Then  '正常に処理された(=テーブルが存在した)時の条件分岐
		tbl.TableStyle = ""
		tbl.Delete
		Err.Clear
	ElseIf Err.Number = 1004 Then   'テーブルが無かったら1004吐くのでそれを検出する
		Err.Clear
	Else    'それ以外
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
End Sub