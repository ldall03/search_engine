csv_format = 6

Set objFSO = CreateObject("Scripting.FileSystemObject")

src_file = objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
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
iNumberOfLinesToDelete = 3

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(dest_file, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

arrLines = Split(strContents, vbNewLine)
Set objTS = objFS.OpenTextFile(dest_file, FOR_WRITING)

For i=0 To UBound(arrLines)
   If i > (iNumberOfLinesToDelete - 1) Then
      objTS.WriteLine arrLines(i)
   End If
Next
