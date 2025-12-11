Attribute VB_Name = "NewMacros"
Sub ExportAllFiguresAsPDFs_SelectFolder()
    Dim shp As InlineShape
    Dim doc As Document
    Dim newDoc As Document
    Dim i As Long
    Dim pdfPath As String
    Dim fileBase As String
    Dim outFolder As String
    Dim picWidthPts As Single, picHeightPts As Single
    Dim figRange As Range
    Dim newShp As InlineShape
    
    ' Ask user for destination folder
    outFolder = SelectFolder("Select destination folder for exported PDFs:")
    
    If outFolder = "" Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Ensure folder ends with a backslash
    If Right(outFolder, 1) <> "\" Then outFolder = outFolder & "\"
    
    Set doc = ActiveDocument
    
    If doc.InlineShapes.Count = 0 Then
        MsgBox "No figures found in this document.", vbExclamation
        Exit Sub
    End If
    
    i = 1
    
    For Each shp In doc.InlineShapes
        
        ' Get the range of the figure (NO CLIPBOARD!)
        Set figRange = shp.Range.Duplicate
        
        ' Create new doc
        Set newDoc = Documents.Add
        
        ' Insert figure directly
        newDoc.Content.FormattedText = figRange.FormattedText
        
        ' Reference the inserted shape
        Set newShp = newDoc.InlineShapes(1)
        
        ' Get size
        picWidthPts = newShp.Width
        picHeightPts = newShp.Height
        
        ' Fit page to image
        With newDoc.Sections(1).PageSetup
            .TopMargin = 1
            .BottomMargin = 1
            .LeftMargin = 1
            .RightMargin = 1
            
            .PageWidth = picWidthPts + .LeftMargin + .RightMargin
            .PageHeight = picHeightPts + .TopMargin + .BottomMargin
        End With
        
        ' File name
        fileBase = "image" & i & ".pdf"
        pdfPath = outFolder & fileBase
        
        ' Save PDF
        newDoc.ExportAsFixedFormat OutputFileName:=pdfPath, _
            ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
            OptimizeFor:=wdExportOptimizeForPrint
        
        ' Close temp doc
        newDoc.Close SaveChanges:=wdDoNotSaveChanges
        
        i = i + 1
    Next shp
    
    MsgBox "All figures exported successfully!", vbInformation
End Sub


'==============================
'  Folder Picker Function
'==============================
Function SelectFolder(prompt As String) As String
    Dim fldr As FileDialog
    Dim result As String
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .title = prompt
        .AllowMultiSelect = False
        If .Show <> -1 Then
            SelectFolder = ""
            Exit Function
        End If
        result = .SelectedItems(1)
    End With
    
    SelectFolder = result
End Function


