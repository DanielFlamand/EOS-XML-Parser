' ============================================================
' EOS_RefreshExcel.vbs
' Opent EOS_Import_Template.xlsm, voert de macro uit en sluit Excel
' ============================================================

Dim oExcel
Dim oWorkbook
Dim sBestand
Dim sResultaat

sBestand = "C:\Users\dflamand\Ultimoo Groep\Proces & Innovatie - IT support en development - IT support en development\PowerQuery\EOS XML aanlevering\EOS_Import_Template.xlsm"

On Error Resume Next

' Excel openen (onzichtbaar)
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = False
oExcel.DisplayAlerts = False

' Werkmap openen
Set oWorkbook = oExcel.Workbooks.Open(sBestand)

If Err.Number <> 0 Then
    ' Fout bij openen bestand
    Dim iFile
    iFile = FreeFile()
    Open "C:\Temp\eos_refresh_result.txt" For Output As #iFile
    Print #iFile, "FOUT: Kan bestand niet openen - " & Err.Description
    Close #iFile
    oExcel.Quit
    Set oExcel = Nothing
    WScript.Quit 1
End If

On Error Resume Next

' Macro uitvoeren
oExcel.Run "VerversEnOpslaan"

If Err.Number <> 0 Then
    ' Fout bij uitvoeren macro
    Dim iFile2
    iFile2 = FreeFile()
    Open "C:\Temp\eos_refresh_result.txt" For Output As #iFile2
    Print #iFile2, "FOUT: Macro mislukt - " & Err.Description
    Close #iFile2
End If

' Excel netjes sluiten
oWorkbook.Close False
oExcel.Quit

Set oWorkbook = Nothing
Set oExcel = Nothing

WScript.Quit 0
