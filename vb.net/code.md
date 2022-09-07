# 기본 코드

- ### Excel
  - 엑셀 파일 열고 닫는 코드
  ```vb
  Dim excelApp As Application
  Dim wb As Workbook
  Dim filePath As String = "C:\Temp.xlsx"
  Dim sheetName As String = "Sheet1"
  Dim valueString As String = "Hello World!"
  
  Try
    excelApp = New Application
    wb = ExcelApp.Workbooks.Open(filePath)
    With CType(wb.Worksheets(sheetName), Worksheet)
      .Range("A1").Value = valueString
    End With
    wb.Save
  Catch e As Exception
    Console.WriteLine(e.ToString)
  Finally
    Try
      wb.Close(SaveChanges:=False)
    Catch e As Exception
    End Try
    Try
      excelApp.Quit
    Catch e As Exception
    End Try
  End Try
  ```

- ### PowerPoint
  - 파워포인트 파일 열고 닫는 코드
  ```vb
  Dim pptApp As Microsoft.Office.Interop.PowerPoint
  Dim pre As Microsoft.Office.Interop.PowerPoint.Presentation
  Dim filePath As String = "C:\Temp.pptx"
  Dim shapeName As String = "table12"
  Dim valueString As String = "Hello World!"
  Try
    pptApp = New Application
    pre = pptApp.Presentations.Open(filePath)
    With pre
      .Slides(1).Shapes(shapeName).TextFrame.TextRange.Text = valueString
    End With
    pre.Save
  Catch e As Exception
    Console.WriteLine(e.ToString)
  Finally
    Try
      pre.Close(SaveChanges:=False)
    Catch e As Exception
    End Try
    Try
      pptApp.Quit
    Catch e As Exception
    End Try
  End Try
  ```
