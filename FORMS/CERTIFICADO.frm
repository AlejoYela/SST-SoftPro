VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CERTIFICADO 
   Caption         =   "Generar certificado médico laboral"
   ClientHeight    =   11940
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   22820
   OleObjectBlob   =   "CERTIFICADO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "CERTIFICADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BotonGenerarCertificdado_Click()
    
    Advertencia = MsgBox("¿Desea generar un certificado médico laboral? Recuerde que la historia clínica debe estar actualizada", vbYesNo + vbQuestion, "Confirmar")
    
    If Advertencia = vbYes Then
        Dim linea, linea2, linea3 As Integer
        Dim fila, fila2, fila3 As Object
        
        codigo = Me.ListaPacientes.List(ListaPacientes.ListIndex, 0)
        
        Set fila = Sheets("BASE DE DATOS 2024").Range("A:A").Find(codigo, lookat:=xlWhole)
        Set fila2 = Sheets("TABLA CERTIFICADOS").Range("A:A").Find(codigo, lookat:=xlWhole)
        Set fila3 = Sheets("TABLA HC").Range("B:B").Find(codigo, lookat:=xlWhole)
        
        linea = fila.Row
        linea2 = fila2.Row
        linea3 = fila3.Row
        
        Sheets("CERTIFICADO").Range("C5").Value = Sheets("TABLA CERTIFICADOS").Range("I" & linea2).Value & ", Colombia"
        Sheets("CERTIFICADO").Range("I5").Value = Sheets("TABLA CERTIFICADOS").Range("C" & linea2).Value
        Sheets("CERTIFICADO").Range("C6").Value = Me.FechaEmision.Value
        Sheets("TABLA CERTIFICADOS").Range("F" & linea2).Value = Me.FechaEmision.Value
        
        If Me.CertificadoPara.Value = "Ingreso" Then
            Sheets("CERTIFICADO").Range("D7").Value = "X"
            Sheets("TABLA CERTIFICADOS").Range("J" & linea2).Value = "Ingreso"
        ElseIf Me.CertificadoPara.Value = "Egreso" Then
            Sheets("CERTIFICADO").Range("F7").Value = "X"
            Sheets("TABLA CERTIFICADOS").Range("J" & linea2).Value = "Egreso"
        ElseIf Me.CertificadoPara.Value = "Periódico" Then
            Sheets("CERTIFICADO").Range("H7").Value = "X"
            Sheets("TABLA CERTIFICADOS").Range("J" & linea2).Value = "Periódico"
        End If
        
        Sheets("CERTIFICADO").Range("B10").Value = Sheets("BASE DE DATOS 2024").Range("B" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("C" & linea).Value
        Sheets("CERTIFICADO").Range("F10").Value = Sheets("BASE DE DATOS 2024").Range("D" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("E" & linea).Value
        Sheets("CERTIFICADO").Range("C11").Value = Sheets("BASE DE DATOS 2024").Range("G" & linea).Value
        Sheets("CERTIFICADO").Range("F11").Value = Sheets("BASE DE DATOS 2024").Range("H" & linea).Value
        Sheets("CERTIFICADO").Range("H11").Value = Sheets("BASE DE DATOS 2024").Range("J" & linea).Value & ", " & Sheets("BASE DE DATOS 2024").Range("I" & linea).Value
        Sheets("CERTIFICADO").Range("B13").Value = Sheets("BASE DE DATOS 2024").Range("Q" & linea).Value
        Sheets("CERTIFICADO").Range("C12").Value = Sheets("BASE DE DATOS 2024").Range("N" & linea).Value
        Sheets("CERTIFICADO").Range("H11").Value = Sheets("BASE DE DATOS 2024").Range("J" & linea).Value
        Sheets("CERTIFICADO").Range("G12").Value = Sheets("BASE DE DATOS 2024").Range("M" & linea).Value & " (" & Sheets("BASE DE DATOS 2024").Range("L" & linea).Value & "), " & Sheets("BASE DE DATOS 2024").Range("N" & linea).Value
        Sheets("CERTIFICADO").Range("F13").Value = Sheets("TABLA CERTIFICADOS").Range("D" & linea2).Value
        Sheets("CERTIFICADO").Range("I13").Value = Sheets("BASE DE DATOS 2024").Range("T" & linea).Value
        
        Sheets("CERTIFICADO").Range("A17").Value = UCase(Sheets("TABLA CERTIFICADOS").Range("AP" & linea2).Value)
        
        If Sheets("CERTIFICADO").Range("A17").Value = "APTO" Then
            Sheets("CERTIFICADO").Range("A17").Interior.Color = RGB(198, 224, 180)
        ElseIf Sheets("CERTIFICADO").Range("A17").Value = "APTO CON RESTRICCIONES QUE NO INTERFIEREN CON SU TRANAJO NORMAL" Then
            Sheets("CERTIFICADO").Range("A17").Interior.Color = RGB(255, 230, 153)
        ElseIf Sheets("CERTIFICADO").Range("A17").Value = "APTO CON RESTRICCIONES QUE LIMITAN SU TRABAJO NORMAL" Then
            Sheets("CERTIFICADO").Range("A17").Interior.Color = RGB(248, 203, 173)
        ElseIf Sheets("CERTIFICADO").Range("A17").Value = "APLAZADO" Then
            Sheets("CERTIFICADO").Range("A17").Interior.Color = RGB(219, 219, 219)
        ElseIf Sheets("CERTIFICADO").Range("A17").Value = "NO APTO" Then
            Sheets("CERTIFICADO").Range("A17").Interior.Color = RGB(255, 177, 177)
        End If
        
        Sheets("CERTIFICADO").Range("C18").Value = Sheets("TABLA HC").Range("Q" & linea3).Value
        Sheets("CERTIFICADO").Range("F18").Value = Sheets("TABLA HC").Range("P" & linea3).Value
        Sheets("CERTIFICADO").Range("H18").Value = Sheets("TABLA CERTIFICADOS").Range("AQ" & linea2).Value
        Sheets("CERTIFICADO").Range("A21").Value = Sheets("TABLA CERTIFICADOS").Range("AS" & linea2).Value
        Sheets("CERTIFICADO").Range("A24").Value = Sheets("TABLA CERTIFICADOS").Range("AR" & linea2).Value
        Sheets("CERTIFICADO").Range("A32").Value = Me.RestriccionesML.Value
        Sheets("TABLA CERTIFICADOS").Range("AT" & linea2).Value = Me.RestriccionesML.Value
        
        Sheets("CERTIFICADO").Select
        
        Advertencia = MsgBox("Certificado generado, ¿Desea ir a la hoja CERTIFICADO para revisarlo y exportarlo?", vbYesNo + vbQuestion, "Confirmar")
        
        If Advertencia = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub BotonSalir_Click()
    Unload Me
End Sub

Private Sub BotonVolver_Click()
    Unload Me
    INICIO.Show
End Sub

Private Sub ListaPacientes_Click()
    codigo = Me.ListaPacientes.List(ListaPacientes.ListIndex, 0)
End Sub

Private Sub UserForm_Initialize()
    With Me
    .Width = Application.Width
    .Height = Application.Height
    End With
    
    Me.ListaPacientes.RowSource = "DATABASE"
    Me.ListaPacientes.ColumnCount = 8
    
    Sheets("OTROS").Select
    Me.CertificadoPara.List = Range("D1:D3").Value
End Sub
