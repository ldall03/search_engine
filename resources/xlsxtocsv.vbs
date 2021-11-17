
csv_format = 6

Set objFSO = CreateObject("Scripting.FileSystemObject")

src_file = objFSO.GetAbsolutePathName("resources/FPS-R30 Ambatovy Hazardous Substance Register Rev01_2021.xlsx")
dest_file = Replace(Replace(src_file,".xlsx",".tmp.csv"),".xls",".tmp.csv")

Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)

oBook.SaveAs dest_file, csv_format

oBook.Close False
oExcel.Quit


' Delete First n Lines of a Text File

Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = "resources/FPS-R30 Ambatovy Hazardous Substance Register Rev01_2021.tmp.csv"
iNumberOfLinesToDelete = 3

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

arrLines = Split(strContents, vbNewLine)
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)

For i=0 To UBound(arrLines)
   If i > (iNumberOfLinesToDelete - 1) Then
      objTS.WriteLine arrLines(i)
   End If
Next
