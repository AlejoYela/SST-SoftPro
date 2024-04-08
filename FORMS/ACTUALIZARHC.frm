VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ACTUALIZARHC 
   Caption         =   "Pacientes actuales"
   ClientHeight    =   11676
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   22820
   OleObjectBlob   =   "ACTUALIZARHC.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ACTUALIZARHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub BuscarPacientes_Click()
    Sheets("BASE DE DATOS 2024").Select
    
    Numerodedatos = Range("A" & Rows.Count).End(xlUp).Row
    
    Me.ListaPacientes = Clear
    Me.ListaPacientes.RowSource = Clear
    
    Y = 0
    
    For fila = 3 To Numerodedatos
        ' Buscar en ambas columnas
        nombreA = ActiveSheet.Cells(fila, 1).Value
        nombreB = ActiveSheet.Cells(fila, 2).Value
        nombreC = ActiveSheet.Cells(fila, 3).Value
        nombreD = ActiveSheet.Cells(fila, 4).Value
        nombreE = ActiveSheet.Cells(fila, 5).Value
        nombreF = ActiveSheet.Cells(fila, 6).Value
        nombreG = ActiveSheet.Cells(fila, 7).Value
        nombreH = ActiveSheet.Cells(fila, 8).Value
        
        If UCase(nombreA) Like "*" & UCase(Me.BuscadorPaciente.Value) & "*" Or _
           UCase(nombreB) Like "*" & UCase(Me.BuscadorPaciente.Value) & "*" Or _
           UCase(nombreC) Like "*" & UCase(Me.BuscadorPaciente.Value) & "*" Or _
           UCase(nombreD) Like "*" & UCase(Me.BuscadorPaciente.Value) & "*" Or _
           UCase(nombreE) Like "*" & UCase(Me.BuscadorPaciente.Value) & "*" Or _
           UCase(nombreF) Like "*" & UCase(Me.BuscadorPaciente.Value) & "*" Or _
           UCase(nombreG) Like "*" & UCase(Me.BuscadorPaciente.Value) & "*" Or _
           UCase(nombreH) Like "*" & UCase(Me.BuscadorPaciente.Value) & "*" Then
            Me.ListaPacientes.AddItem
            Me.ListaPacientes.List(Y, 0) = nombreA
            Me.ListaPacientes.List(Y, 1) = nombreB
            Me.ListaPacientes.List(Y, 2) = nombreC
            Me.ListaPacientes.List(Y, 3) = nombreD
            Me.ListaPacientes.List(Y, 4) = nombreE
            Me.ListaPacientes.List(Y, 5) = nombreF
            Me.ListaPacientes.List(Y, 6) = nombreG
            Me.ListaPacientes.List(Y, 7) = nombreH
            
            Y = Y + 1
        End If
    Next
End Sub

Private Sub CommandButton1_Click()
    Unload Me
    HC.Show
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    CERTIFICADO.Show
End Sub

Private Sub CommandButton3_Click()
 Unload Me
End Sub

Private Sub CommandButton4_Click()
    Unload Me
    INICIO.Show
End Sub

Private Sub CommandButton5_Click()
    Unload Me
    NUEVOFOLIO.Show
End Sub

Private Sub ListaPacientes_Click()
    Me.Label6.Visible = True
    Me.Label7.Visible = True
    Me.NombresCompletos.Visible = True
    Me.Documento.Visible = True
    Sheets("OTROS").Select
    Range("G2").Value = Me.ListaPacientes.List(ListaPacientes.ListIndex, 0)
    
    Dim linea As Integer
    Dim fila As Object
    
    
    IdHC = Sheets("OTROS").Range("G2").Value
    
    Set fila = Sheets("BASE DE DATOS 2024").Range("A:A").Find(IdHC, lookat:=xlWhole)
    linea = fila.Row
    
    Me.NombresCompletos.Value = Sheets("BASE DE DATOS 2024").Range("B" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("C" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("D" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("E" & linea).Value
    Me.Documento.Value = Sheets("BASE DE DATOS 2024").Range("G" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("H" & linea).Value
End Sub

Private Sub UserForm_Initialize()
    
    Sheets("OTROS").Select
    Sheets("BASE DE DATOS 2024").Select
    
    With Me
        .Width = Application.Width
        .Height = Application.Height
    End With
    
    Me.ListaPacientes.RowSource = "DATABASE"
    Me.ListaPacientes.ColumnCount = 8
End Sub
