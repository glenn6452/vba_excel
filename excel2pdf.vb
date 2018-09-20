Public Sub ASMReportPDF()
Dim sC As SlicerCache
Set wbA = ActiveWorkbook
Set wsA = ActiveSheet
strTime = Format(Now(), "yyyymmdd")
Set sC = wbA.SlicerCaches("Slicer_ASM")




  'This reminds the user to only select the first slicer item
   If sC.VisibleSlicerItems.Count <> 1 Or sC.SlicerItems(1).Selected = False Then
      MsgBox "Please Only Select ASM 1"
      Exit Sub
   End If


For i = 1 To sC.SlicerItems.Count

    'Do not clear ilter as it causes to select all of the items (sC.ClearManualFilter)

    sC.SlicerItems(i).Selected = True
    If i <> 1 Then sC.SlicerItems(i - 1).Selected = False
    Range("D29") = sC.SlicerItems(i).Name

'get active workbook folder, if saved
strPath = wbA.Path
If strPath = "" Then
  strPath = Application.DefaultFilePath
End If
strPath = strPath & "\"

'replace spaces and periods in sheet name
strName = Replace(Range("D29").Value, " ", "_")
strName = Replace(strName, ".", "_")
strName = Replace(strName, "-", "_")

'create default name for savng file
strFile = strName & "_" & strTime & ".pdf"
strPathFile = strPath & strFile

 'export to PDF in current folder
    wsA.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=strPathFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    'confirmation message with file info
Next

End Sub