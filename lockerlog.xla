Sub visit()
	Dim f3 As Worksheet 
		Set f3 = ThisWorkbook.Sheets("3f")	' main sheet
	Dim f3log As Worksheet
		Set f3log = ThisWorkbook.Sheets("3flog")	' log sheet

	Dim btn As Object
		Set btn = ActiveSheet.Buttons(Application.Caller)

	Dim btnCell As Range
		Set btnCell = btn.TopLeftCell.Offset(0, 0)	' save button location

		Dim name As Range	
			Set name = f3.Cells(btnCell.Row, btnCell.Column - 3)	

		Dim phoneno As Range
			Set phoneno = f3.Cells(btnCell.Row, btnCell.Column - 2)

		' saving basic information on locker users
		f3log.Cells(btnCell.Row, 3).Value = phoneno.Value
		f3log.Cells(btnCell.Row, 2).Value = name.Value
		f3log.Cells(btnCell.Row, 1).Value = f3.Cells(btnCell.Row, btnCell.Column - 6)	' lockernumber

	Dim targetCell As Range
		Set targetCell = f3log.Cells(btnCell.Row, 4)

	Do While targetCell.Value <> ""	' while loop while targetCell.Value != N/A

		If targetCell.Value = Date Then
			MsgBox "Error: today's log has been already updated", vbExclamation, "Duplicate Date"
			Exit Do
		End If

		Set targetCell = targetCell.Offset(0, 1)	' move to next column
	Loop

	targetCell.Value = Date ' establish log

	End Sub
