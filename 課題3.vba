Option Explicit
'インデントは普段書いてるプログラムのノリでやってます VBAですがオールマンスタイルです
'ライセンス:GPLv3

Sub ボタン1_Click()
	Dim rtn As Integer
	Dim tbl As ListObject
	' Workbooks.Open(Thisworkbook.Path & "\業務実績_3月.xlsm")
	' Workbooks.Open(Thisworkbook.Path & "\業務実績_4月.xlsm")
	' Workbooks.Open(Thisworkbook.Path & "\業務実績_5月.xlsm")
	On Error Resume Next    '究極のゴリ押しを実現する魔法の呪文 C#で言うなら例外処理を処理に組み込んでます
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_a").ListObject    'エラーが出ようが出まいが変数に代入する
	If Err.Number = 0 Then  '正常に処理された(=テーブルが存在した)時の条件分岐
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_3月.xlsm!Reset")
				Call Application.Run("業務実績_3月.xlsm!ボタン1_Click") '呼び出し
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_a").Range.Cut Range("B6")   '移動
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then   'テーブルが無かったら1004吐くのでそれを検出する
		Call Application.Run("業務実績_3月.xlsm!Reset")
		Call Application.Run("業務実績_3月.xlsm!ボタン1_Click") '呼び出し
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_a").Range.Cut Range("B6")   '移動
		Err.Clear
	Else    'それ以外
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	'以下全部同じ処理の繰り返し
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_b").ListObject
	If Err.Number = 0 Then
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_3月.xlsm!Reset")
				Call Application.Run("業務実績_3月.xlsm!ボタン2_Click")
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_b").Range.Cut Range("E6")
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then
		Call Application.Run("業務実績_3月.xlsm!Reset")
		Call Application.Run("業務実績_3月.xlsm!ボタン2_Click")
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_b").Range.Cut Range("E6")
		Err.Clear
	Else
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_c").ListObject
	If Err.Number = 0 Then
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_3月.xlsm!Reset")
				Call Application.Run("業務実績_3月.xlsm!ボタン3_Click")
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_c").Range.Cut Range("H6")
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then
		Call Application.Run("業務実績_3月.xlsm!Reset")
		Call Application.Run("業務実績_3月.xlsm!ボタン3_Click")
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_c").Range.Cut Range("H6")
		Err.Clear
	Else
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Range("B5").Value = "3月分統計"
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_d").ListObject
	If Err.Number = 0 Then
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_4月.xlsm!Reset")
				Call Application.Run("業務実績_4月.xlsm!ボタン1_Click")
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_d").Range.Cut Range("K6")
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then
		Call Application.Run("業務実績_4月.xlsm!Reset")
		Call Application.Run("業務実績_4月.xlsm!ボタン1_Click")
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_d").Range.Cut Range("K6")
		Err.Clear
	Else
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_e").ListObject
	If Err.Number = 0 Then
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_4月.xlsm!Reset")
				Call Application.Run("業務実績_4月.xlsm!ボタン2_Click")
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_e").Range.Cut Range("N6")
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then
		Call Application.Run("業務実績_4月.xlsm!Reset")
		Call Application.Run("業務実績_4月.xlsm!ボタン2_Click")
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_e").Range.Cut Range("N6")
		Err.Clear
	Else
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_f").ListObject
	If Err.Number = 0 Then
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_4月.xlsm!Reset")
				Call Application.Run("業務実績_4月.xlsm!ボタン3_Click")
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_f").Range.Cut Range("Q6")
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then
		Call Application.Run("業務実績_4月.xlsm!Reset")
		Call Application.Run("業務実績_4月.xlsm!ボタン3_Click")
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_f").Range.Cut Range("Q6")
		Err.Clear
	Else
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Range("K5").Value = "4月分統計"
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_g").ListObject
	If Err.Number = 0 Then
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_5月.xlsm!Reset")
				Call Application.Run("業務実績_5月.xlsm!ボタン1_Click")
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_g").Range.Cut Range("T6")
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then
		Call Application.Run("業務実績_5月.xlsm!Reset")
		Call Application.Run("業務実績_5月.xlsm!ボタン1_Click")
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_g").Range.Cut Range("T6")
		Err.Clear
	Else
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_h").ListObject
	If Err.Number = 0 Then
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_5月.xlsm!Reset")
				Call Application.Run("業務実績_5月.xlsm!ボタン2_Click")
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_h").Range.Cut Range("W6")
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then
		Call Application.Run("業務実績_5月.xlsm!Reset")
		Call Application.Run("業務実績_5月.xlsm!ボタン2_Click")
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_h").Range.Cut Range("W6")
		Err.Clear
	Else
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Set tbl = ActiveWorkBook.Worksheets(1).Range("課題1_i").ListObject
	If Err.Number = 0 Then
		rtn = MsgBox("競合するテーブルが存在します。削除してもよろしいですか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認")
		Select Case rtn '消去確認
			Case vbYes  'テーブル消去
				tbl.TableStyle = ""
				tbl.Delete
				Call Application.Run("業務実績_5月.xlsm!Reset")
				Call Application.Run("業務実績_5月.xlsm!ボタン3_Click")
				ActiveWorkBook.ActiveSheet.ListObjects("課題1_i").Range.Cut Range("Z6")
				Err.Clear
		End Select
	ElseIf Err.Number = 1004 Then
		Call Application.Run("業務実績_5月.xlsm!Reset")
		Call Application.Run("業務実績_5月.xlsm!ボタン3_Click")
		ActiveWorkBook.ActiveSheet.ListObjects("課題1_i").Range.Cut Range("Z6")
		Err.Clear
	Else
		MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, Buttons:=vbCritical, Title:="ERROR"
		Err.Clear
	End If
	Range("T5").Value = "5月分統計"
End Sub