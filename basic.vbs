Dim objExcel, objWorkbook

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Swetha Panchumarthi\GitTest\Book1.xlsm")
objExcel.Run "test"
objWorkbook.Save
objWorkbook.Close
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing

