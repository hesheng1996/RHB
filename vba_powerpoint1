

Sub CopyDataIntoPowerpoint()
  
  Dim path As String
  Dim file As String

 
  
  path = Worksheets("Control").Range("C27").Value
  file = Worksheets("Control").Range("C28").Value
  
  Dim ppt As PowerPoint.Application
  Dim pres As PowerPoint.Presentation
  Dim pptSlide As PowerPoint.Slide
  Dim ppTbl  As PowerPoint.Shape
  
  Set ppt = New PowerPoint.Application
  Set pres = ppt.Presentations.Open(path & "\" & file & ".pptx")
  'Set pptSlide = pres.Slides(13)

  
   
  Dim x As Integer
  x = Worksheets("Control").Range("C29").Value
  
  Application.DisplayAlerts = False
  
  ppt.ActivePresentation.Slides(x).Select
  
  
  
   'Corp********************************************************************************************************************************************************************
 
  Dim excel_Corp, ppt_Corp, Num_A, A
    
    Num_A = Array(1, 2, 3, 4, 5)
    excel_Corp = Array("D3:D6", "G3:G6", "J3:J6", "M3:M6", "P3:P6")
    ppt_Corp = Array("185", "181", "199", "322", "201")
    
    For A = 0 To UBound(Num_A)
    
  ThisWorkbook.Sheets("Template").Range(excel_Corp(A)).Copy
  ppt.ActivePresentation.Slides(x).Shapes("Table " & ppt_Corp(A)).Table.Rows(1).Cells.Item(2).Select
  ppt.CommandBars.ExecuteMso ("PasteExcelTableDestinationTableStyle")
     
             ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
    Next A
                
    'Comm************************************************************************************************************************************************************************
  
  Dim excel_Comm, ppt_Comm, Num_B, B
    
    Num_B = Array(1, 2, 3, 4, 5)
    excel_Comm = Array("D7:D10", "G7:G10", "J7:J10", "M7:M10", "P7:P10")
    ppt_Comm = Array("228", "223", "229", "325", "231")
    
    For B = 0 To UBound(Num_B)
    
  ThisWorkbook.Sheets("Template").Range(excel_Comm(B)).Copy
  ppt.ActivePresentation.Slides(x).Shapes("Table " & ppt_Comm(B)).Table.Rows(1).Cells.Item(2).Select
  ppt.CommandBars.ExecuteMso ("PasteExcelTableDestinationTableStyle")
     
             ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
    Next B
  'MM****************************************************************************************************************************************************************************
   Dim excel_MM, ppt_MM, Num_M, M
    
    Num_M = Array(1, 2, 3, 4, 5)
    excel_MM = Array("D11:D14", "G11:G14", "J11:J14", "M11:M14", "P11:P14")
    ppt_MM = Array("280", "265", "281", "327", "282")
    
    For M = 0 To UBound(Num_M)
    
  ThisWorkbook.Sheets("Template").Range(excel_MM(M)).Copy
  ppt.ActivePresentation.Slides(x).Shapes("Table " & ppt_MM(M)).Table.Rows(1).Cells.Item(2).Select
  ppt.CommandBars.ExecuteMso ("PasteExcelTableDestinationTableStyle")
     
             ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
    Next M
    'RSME****************************************************************************************************************************************************************************
   Dim excel_RSME, ppt_RSME, Num_C, C
    
    Num_C = Array(1, 2, 3, 4, 5)
    excel_SME = Array("D15:D18", "G15:G18", "J15:J18", "M15:M18", "P15:P18")
    ppt_SME = Array("351", "352", "353", "356", "357")
    
    For C = 0 To UBound(Num_C)
    
  ThisWorkbook.Sheets("Template").Range(excel_SME(C)).Copy
  ppt.ActivePresentation.Slides(x).Shapes("Table " & ppt_SME(C)).Table.Rows(1).Cells.Item(2).Select
  ppt.CommandBars.ExecuteMso ("PasteExcelTableDestinationTableStyle")
     
             ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
    Next C
    
   
  
    'Retail****************************************************************************************************************************************************************************
  Dim excel_Retail, ppt_Retail, Num_D, D
    
    Num_D = Array(1, 2, 3, 4, 5)
    excel_Retail = Array("D19:D22", "G19:G22", "J19:J22", "M19:M22", "P19:P22")
    ppt_Retail = Array("424", "433", "434", "437", "438")
    
    For D = 0 To UBound(Num_D)
    
  ThisWorkbook.Sheets("Template").Range(excel_Retail(D)).Copy
  ppt.ActivePresentation.Slides(x).Shapes("Table " & ppt_Retail(D)).Table.Rows(1).Cells.Item(2).Select
  ppt.CommandBars.ExecuteMso ("PasteExcelTableDestinationTableStyle")
     
             ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
    Next D
    
    

    
    
   'ScoreCard Table shape****************************************************************************************************************************************************************
    
    Dim excel_Table, ppt_Table, Num_F, F
    
    Num_F = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    excel_Table = Array("S3", "S6", "S7", "S10", "S11", "S14", "S15", "S18", "S19", "S22")
    ppt_Table = Array("337", "347", "339", "338", "340", "341", "343", "361", "407", "442")
    
    For F = 0 To UBound(Num_F)
  ThisWorkbook.Sheets("Template").Range(excel_Table(F)).Copy
  ppt.ActivePresentation.Slides(x).Shapes("Table " & ppt_Table(F)).Table.Rows(1).Cells.Item(1).Select
  ppt.CommandBars.ExecuteMso ("PasteTextOnly")
  
        ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
                
    Next F
    
    
   'Table Colour*************************************************************************************************************************************************************************
   
   Dim Table_excel, Table_ppt, Num_FF, FF
    
    
    
    Num_FF = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    Table_excel = Array("S3", "S6", "S7", "S10", "S11", "S14", "S15", "S18", "S19", "S22")
    Table_ppt = Array("337", "347", "339", "338", "340", "341", "343", "361", "407", "442")
    
    For FF = 0 To UBound(Num_FF)
    
   If ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) <= 1 Then
          ppt.ActivePresentation.Slides(x).Shapes("Table " & Table_ppt(FF)).Fill.ForeColor.RGB = RGB(255, 0, 0)
          
    ElseIf ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) > 1 And ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) <= 2 Then
          ppt.ActivePresentation.Slides(x).Shapes("Table " & Table_ppt(FF)).Fill.ForeColor.RGB = RGB(255, 192, 0)
          
    ElseIf ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) > 2 And ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) <= 3 Then
          ppt.ActivePresentation.Slides(x).Shapes("Table " & Table_ppt(FF)).Fill.ForeColor.RGB = RGB(0, 255, 0)
          
    ElseIf ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) > 3 And ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) <= 4 Then
          ppt.ActivePresentation.Slides(x).Shapes("Table " & Table_ppt(FF)).Fill.ForeColor.RGB = RGB(111, 205, 228)
     
    ElseIf ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) > 4 And ThisWorkbook.Sheets("Template").Range(Table_excel(FF)) <= 5 Then
          ppt.ActivePresentation.Slides(x).Shapes("Table " & Table_ppt(FF)).Fill.ForeColor.RGB = RGB(223, 192, 255)
     
  End If
           ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
    Next FF
    
  'ScoreCard Oval shape****************************************************************************************************************************************************************
    
    Dim excel_Oval, ppt_Oval, Num_G, G
    
    Num_G = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25)
    excel_Oval = Array("E3", "H3", "K3", "N3", "Q3", "E7", "H7", "K7", "N7", "Q7", "E11", "H11", "K11", "N11", "Q11", "E15", "H15", "K15", "N15", "Q15", "E19", "H19", "K19", "N19", "Q19")
    ppt_Oval = Array("289", "330", "197", "350", "200", "291", "295", "227", "324", "296", "297", "331", "333", "328", "279", "315", "332", "355", "329", "359", "372", "379", "436", "378", "440")
    For G = 0 To UBound(Num_G)
    
   ThisWorkbook.Sheets("Template").Range(excel_Oval(G)).Copy
       
   ppt.ActivePresentation.Slides(x).Shapes("Oval " & ppt_Oval(G)).TextFrame.TextRange.Paste
 
        ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
                
  
   Next G

  'Oval Colour*************************************************************************************************************************************************************************
   
   Dim Oval_num, Num_J, J
    
    Num_J = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25)
    Oval_num = Array("289", "330", "197", "350", "200", "291", "295", "227", "324", "296", "297", "331", "333", "328", "279", "315", "332", "355", "329", "359", "372", "379", "436", "378", "440")
    
    For J = 0 To UBound(Num_J)
    
   If ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).TextFrame.TextRange = 1 Then
          ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).Fill.ForeColor.RGB = RGB(255, 0, 0)
          
    ElseIf ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).TextFrame.TextRange = 2 Then
          ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).Fill.ForeColor.RGB = RGB(255, 192, 0)
          
    ElseIf ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).TextFrame.TextRange = 3 Then
          ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).Fill.ForeColor.RGB = RGB(0, 255, 0)
          
    ElseIf ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).TextFrame.TextRange = 4 Then
          ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).Fill.ForeColor.RGB = RGB(111, 205, 228)
     
    ElseIf ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).TextFrame.TextRange = 5 Then
          ppt.ActivePresentation.Slides(x).Shapes("Oval " & Oval_num(J)).Fill.ForeColor.RGB = RGB(223, 192, 255)
     
  End If
          ' Introduce delay to let paste action happen before moving on
                Wait = Timer
                While Timer < Wait + 0.5
                DoEvents
                Wend
    Next J
  
  
  MsgBox "Finished!"
  
  

End Sub

