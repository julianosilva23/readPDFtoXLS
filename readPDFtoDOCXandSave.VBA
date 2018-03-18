Sub ConvertPDFsToWord2()
    Dim path As String
    'Manually edit path in the next line before running
    path = "C:\Users\Admin\Desktop\fichas1\"

    Dim file As String
    Dim doc As Word.Document
    Dim regValPDF As Integer
    Dim originalAlertLevel As WdAlertLevel

'Generate string for getting all PDFs with Dir command
    'Check for terminal \
    If Right(path, 1) <> "\" Then path = path & "\"
    'Append file type with wildcard
    file = path & "*.pdf"

    'Get path for first PDF (blank string if no PDFs exist)
    file = Dir(file)

    originalAlertLevel = Application.DisplayAlerts
    Application.DisplayAlerts = wdAlertsNone

    If file <> "" Then regValPDF = TogglePDFWarning(1)

    Do While file <> ""
        'Open method will automatically convert PDF for editing
        Set doc = Documents.Open(path & file, False)

        'Save and close document
        doc.SaveAs2 path & Replace(file, ".pdf", ".docx"), _
                    fileformat:=wdFormatDocumentDefault
        doc.Close False

        'Get path for next PDF (blank string if no PDFs remain)
        file = Dir
    Loop

CleanUp:
    On Error Resume Next 'Ignore errors during cleanup
    doc.Close False
    'Restore registry value, if necessary
    If regValPDF <> 1 Then TogglePDFWarning regValPDF
    Application.DisplayAlerts = originalAlertLevel

End Sub

Private Function TogglePDFWarning(newVal As Integer) As Integer
'This function reads and writes the registry value that controls
'the dialog displayed when Word opens (and converts) a PDF file
    Dim wShell As Object
    Dim regKey As String
    Dim regVal As Variant

    'setup shell object and string for key
    Set wShell = CreateObject("WScript.Shell")
    regKey = "HKCU\SOFTWARE\Microsoft\Office\" & _
             Application.Version & "\Word\Options\"

    'Get existing registry value, if any
    On Error Resume Next 'Ignore error if reg value does not exist
    regVal = wShell.RegRead(regKey & "DisableConvertPdfWarning")
    On Error GoTo 0      'Break on errors after this point

    wShell.regwrite regKey & "DisableConvertPdfWarning", newVal, "REG_DWORD"

    'Return original setting / registry value (0 if omitted)
    If Err.Number <> 0 Or regVal = 0 Then
        TogglePDFWarning = 0
    Else
        TogglePDFWarning = 1
    End If

End Function

