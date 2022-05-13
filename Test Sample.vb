' ========== Test Sample ==========
Path="D:\RPAWORK\RPATEST\과제관리\20220105\Input\Test.xlsx"
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(Path)
Set objWorkSheet = objWorkbook.Sheets(1)
objExcel.ScreenUpdating = False '스크린 업데이트 없이 진행
objExcel.Visible = False
objExcel.DisplayAlerts = False '팝업해제

On Error Resume Next
objWorkSheet.Activate
'elea = eroro() 'Error 처리 테스트용 코드
' ========== Last Row and Column ==========
'Range 첫 문자열은 컬럼 문자
'최대범위에서 컬럼 문자 기준으로 마지막 데이터가 있는 Row를 변수에 저장
Last_row = objWorkSheet.Range("C65536").End(-4162).row
'Range 마지막 슷자은 Row
'최대범위에서 Row 번호 기준으로 마지막 데이터가 있는 Col을 변수에 저장
Last_col = objWorkSheet.Range("XFD4").End(-4159).Column

' ========== 값입력과 수식입력 ==========
For idx = "%RowNum%" to 6	
	objWorkSheet.Range("H" & idx).Value = "X"
	objWorkSheet.Range("I" & idx).Value = "X"
	objWorkSheet.Range("J" & idx).Value = "X"
	objWorkSheet.Range("G7").formula = "=SUM(G4:G6)"
Next

'작업 수행 후 저장 예시
'objWorkbook.Save

'다른이름으로 저장 예시
objWorkbook.SaveCopyAs "%Excel_SaveAs_Root%"

'Excel Workbook 닫기
'objWorkbook.Close False '저장하지 않고 종료
objWorkbook.Close

' ========== 결과전달 케이스1 ==========
If Err.Number <> 0 Then
	Wscript.Echo "Error"
Else
	'결과 값 전달 케이스 1
	Wscript.Echo "OK"
End If

' ========== 결과전달 케이스2 ==========
'If Err.Number <> 0 Then
'	Wscript.Echo "Error"
'Else
'	'결과 값 전달 케이스 2
'	Wscript.Echo Last_row & "|" & Last_col
'End If