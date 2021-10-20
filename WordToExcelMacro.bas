Attribute VB_Name = "Module21"

    

Sub WordToExcel()
    
    Dim wdDoc As Word.Document
    Dim xlsDoc As Workbook
    Dim WordFileLocation As String
    Dim ExcelFileLocation As String
    Dim oSection As Section
    Dim oHeader As HeaderFooter
    Dim ParagraphText As String
    Dim ParagraphStyle As String
    Dim i As Integer
    
    WordFileLocation = "C:\Users\Boban\Desktop\Document Management Specialist -task - input - Copy (2).docx"
    ExcelFileLocation = "C:\Users\Boban\Documents\WordToExcel.xlsx"
    
    
    Rem Opens Word and Excel Applications if none are present
    
    On Error Resume Next
    Set objExcel = GetObject(, "Excel.Application")

    Set objWord = GetObject(, "Word.Application")

    If objWord Is Nothing Then
        Set objWord = CreateObject("Word.Application")
    End If
    
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.Application")
    End If
    
    On Error GoTo 0
    
    Rem Opens Word and Excel Documents
    
    Set wdDoc = objWord.Documents.Open(WordFileLocation)
    Set xlsDoc = objExcel.Workbooks.Add
      
    
    Rem Attempt to identify an image in Word header and place it in an Excel cell
    
'     For Each oSection In ActiveDocument.Sections
'        For Each oHeader In oSection.Headers
'           For Each oShape In oHeader.Range.InlineShapes
'                    i = i + 1
'
'                    oShape.Select
'                    Selection.Copy
'
'                   ( Couldn't place an image in Excel document after )
'
'        Next oHeader
'    Next oSection
    
    
    Rem loops through header paragraps, manipulates them and places them to an Excel document in order defined by given instructions
    
    i = 0
    
    For Each oSection In ActiveDocument.Sections
        For Each oHeader In oSection.Headers
            For Each Paragraph In oHeader.Range.Paragraphs
            
                If Len(Paragraph.Range.Text) > 1 And Left(Paragraph.Range.Text, Len("/")) <> "/" Then
                    i = i + 1
                        xlsDoc.Worksheets("Sheet1").Range("A" & i).Value = Paragraph.Range.Text
                End If
                
            Next
        Next oHeader
    Next oSection
    
    Rem loops through document paragraphs, manipulates them and places them to an Excel document in order defined by given instructions
    
    For Each Paragraph In wdDoc.Paragraphs
        
         ParagraphText = Paragraph.Range.Text
         ParagraphStyle = Paragraph.Range.Style
           
            If ParagraphStyle = "Heading 1" Then
                i = i + 1
                    xlsDoc.Worksheets("Sheet1").Range("A" & i).Value = UCase(ParagraphText)
                
            End If
            
            If ParagraphStyle = "Heading 2" Then
                i = i + 1
                    xlsDoc.Worksheets("Sheet1").Range("A" & i).Value = ParagraphText
                        xlsDoc.Worksheets("Sheet1").Range("A" & i).Font.Bold = True
                            xlsDoc.Worksheets("Sheet1").Range("A" & i).Font.Underline = xlUnderlineStyleSingle
                       
                
            End If
            
            
            If Len(ParagraphText) > 1 And Left(ParagraphStyle, Len("Heading")) <> "Heading" Then
                i = i + 1
                    xlsDoc.Worksheets("Sheet1").Range("A" & i).Value = ParagraphText
                
            End If
    Next
        
    Rem Closes Word Application and document
    
    wdDoc.Close
    Set objWord = Nothing
    
    Rem Saves Output in given Excel location and closes Excel Application and document
    
    xlsDoc.SaveAs (ExcelFileLocation)
    xlsDoc.Close
    Set objExcel = Nothing
    


End Sub



