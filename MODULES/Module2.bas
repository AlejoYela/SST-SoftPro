Attribute VB_Name = "Module2"
Sub BotónExportarPDF()
    Dim folderPath As String
    Dim patientFolder As String

    ' Ruta de la carpeta principal "CERTIFICADOS"
    folderPath = ActiveWorkbook.Path & "\CERTIFICADOS\"

    ' Verificar si la carpeta principal existe, si no, crearla
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    ' Nombre de la subcarpeta con el número de documento del paciente
    patientFolder = folderPath & ActiveSheet.Range("F11").Value & "\"

    ' Verificar si la subcarpeta del paciente existe, si no, crearla
    If Dir(patientFolder, vbDirectory) = "" Then
        MkDir patientFolder
    End If

    ' Exportar como PDF en la subcarpeta del paciente
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=patientFolder & "Certificado.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

    MsgBox ("Certificado exportado correctamente en la carpeta " & patientFolder)
End Sub
    
