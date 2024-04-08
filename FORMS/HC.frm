VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HC 
   Caption         =   "Historia clínica electrónica"
   ClientHeight    =   12396
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   22820
   OleObjectBlob   =   "HC.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "HC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BotonModificar_Click()
    Advertencia = MsgBox("¿Está seguro de que desea modificar la historia clínica de este paciente?, solo podrá modificar datos personales", vbYesNo + vbQuestion, "Confirmar")
    
    If Advertencia = vbYes Then
        Dim control As control
        
        For Each control In Me.Controls
            If TypeName(control) = "TextBox" Or TypeName(control) = "ComboBox" Then
                control.Locked = False
                control.SpecialEffect = fmSpecialEffectSunken
                control.BackColor = RGB(255, 255, 255)
            End If
        Next control
        Me.CommandButton1.Enabled = True
    End If

End Sub

Private Sub BotonVerEscala_Click()
    IMC_FORM.Show
End Sub

Private Sub CommandButton1_Click()

    Dim linea As Integer
    Dim fila As Object
    
    
    IdHC = Sheets("OTROS").Range("G2").Value
    
    Set fila = Sheets("BASE DE DATOS 2024").Range("A:A").Find(IdHC, lookat:=xlWhole)
    linea = fila.Row
    
    Sheets("BASE DE DATOS 2024").Select

    Range("G" & linea).Value = Me.TipoDocumento.Value
    Range("H" & linea).Value = Me.NumeroDocumento.Value
    Range("K" & linea).Value = Me.FechaExpedicion.Value
    Range("N" & linea).Value = Me.FechaNacimiento.Value
    Range("P" & linea).Value = Me.Edad.Value
    Range("R" & linea).Value = Me.OcupacionActual.Value
    Range("S" & linea).Value = Me.Direccion.Value
    Range("T" & linea).Value = Me.Telefono.Value
    Range("X" & linea).Value = Me.ZonaResidencia.Value
    Range("Y" & linea).Value = Me.GrupoAE.Value
    Range("Z" & linea).Value = Me.ARL.Value
    Range("AA" & linea).Value = Me.Pensiones.Value
    Range("AB" & linea).Value = Me.Aseguradora.Value
    Range("AC" & linea).Value = Me.TipoVinculacion.Value
    Range("AD" & linea).Value = Me.NombreAcudiente.Value
    Range("AE" & linea).Value = Me.Parentezco.Value
    Range("AF" & linea).Value = Me.TelefonoAcudiente.Value

    
    Sheets("OTROS").Select
    Sheets("BASE DE DATOS 2024").Select
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    ACTUALIZARHC.Show
End Sub

Private Sub CommandButton3_Click()
 Unload Me
End Sub
Private Sub UserForm_Initialize()

    With Me
        .Width = Application.Width
        .Height = Application.Height
    End With
    
    Dim linea, linea2, linea3 As Integer
    Dim fila, fila2, fila3 As Object
    
    IdHC = Sheets("OTROS").Range("G2").Value
    
    Set fila = Sheets("BASE DE DATOS 2024").Range("A:A").Find(IdHC, lookat:=xlWhole)
    Set fila2 = Sheets("TABLA CERTIFICADOS").Range("A:A").Find(IdHC, lookat:=xlWhole)
    Set fila3 = Sheets("TABLA HC").Range("B:B").Find(IdHC, lookat:=xlWhole)
    
    linea = fila.Row
    linea2 = fila2.Row
    linea3 = fila3.Row
    
    Me.Imagen1.Picture = LoadPicture(Sheets("BASE DE DATOS 2024").Range("AG" & linea).Value)
    
    Me.PrimerNombre.Value = Sheets("BASE DE DATOS 2024").Range("B" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("C" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("D" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("E" & linea).Value
    Me.TipoDocumento.Value = Sheets("BASE DE DATOS 2024").Range("G" & linea).Value
    Me.NumeroDocumento.Value = Sheets("BASE DE DATOS 2024").Range("H" & linea).Value
    Me.FechaExpedicion.Value = Sheets("BASE DE DATOS 2024").Range("K" & linea).Value & " en " & Sheets("BASE DE DATOS 2024").Range("J" & linea).Value & ", " & Sheets("BASE DE DATOS 2024").Range("I" & linea).Value
    Me.LugarNacimiento.Value = Sheets("BASE DE DATOS 2024").Range("M" & linea).Value & ", " & Sheets("BASE DE DATOS 2024").Range("L" & linea).Value
    Me.FechaNacimiento.Value = Sheets("BASE DE DATOS 2024").Range("N" & linea).Value
    Me.Edad.Value = Sheets("BASE DE DATOS 2024").Range("P" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("O" & linea).Value
    Me.ARL.Value = Sheets("BASE DE DATOS 2024").Range("Z" & linea).Value
    Me.Pensiones.Value = Sheets("BASE DE DATOS 2024").Range("AA" & linea).Value
    Me.Aseguradora.Value = Sheets("BASE DE DATOS 2024").Range("AB" & linea).Value
    Me.TipoVinculacion.Value = Sheets("BASE DE DATOS 2024").Range("AC" & linea).Value
    Me.LugarResidencia.Value = Sheets("BASE DE DATOS 2024").Range("W" & linea).Value & ", " & Sheets("BASE DE DATOS 2024").Range("V" & linea).Value
    Me.ZonaResidencia.Value = Sheets("BASE DE DATOS 2024").Range("X" & linea).Value
    Me.Telefono.Value = Sheets("BASE DE DATOS 2024").Range("T" & linea).Value
    Me.Direccion.Value = Sheets("BASE DE DATOS 2024").Range("S" & linea).Value
    Me.GrupoAE.Value = Sheets("BASE DE DATOS 2024").Range("Y" & linea).Value
    Me.OcupacionActual.Value = Sheets("BASE DE DATOS 2024").Range("R" & linea).Value
    
    Me.CargoAOcupar.Value = Sheets("TABLA CERTIFICADOS").Range("D" & linea2).Value
    Me.Entidad.Value = Sheets("TABLA CERTIFICADOS").Range("E" & linea2).Value
    Me.FechaIngreso.Value = Sheets("TABLA CERTIFICADOS").Range("G" & linea2).Value
    
    Me.NombreAcudiente.Value = Sheets("BASE DE DATOS 2024").Range("AD" & linea).Value
    Me.Parentezco.Value = Sheets("BASE DE DATOS 2024").Range("AE" & linea).Value
    Me.TelefonoAcudiente.Value = Sheets("BASE DE DATOS 2024").Range("AF" & linea).Value
    
    Me.FechaAtencion.Value = Sheets("TABLA CERTIFICADOS").Range("H" & linea2).Value
    Me.LugarAtencion.Value = Sheets("TABLA CERTIFICADOS").Range("I" & linea2).Value
    Me.TipoConsulta.Value = Sheets("TABLA CERTIFICADOS").Range("J" & linea2).Value
    
    Me.AntFamiliares.Value = Sheets("TABLA HC").Range("E" & linea3).Value
    Me.AntPatologicos.Value = Sheets("TABLA HC").Range("F" & linea3).Value
    Me.AntFarmacologicos.Value = Sheets("TABLA HC").Range("G" & linea3).Value
    Me.AntQuirurgicos.Value = Sheets("TABLA HC").Range("H" & linea3).Value
    Me.AntTox.Value = Sheets("TABLA HC").Range("I" & linea3).Value
    Me.GinG.Value = Sheets("TABLA HC").Range("J" & linea3).Value
    Me.GinP.Value = Sheets("TABLA HC").Range("K" & linea3).Value
    Me.GinC.Value = Sheets("TABLA HC").Range("L" & linea3).Value
    Me.GinA.Value = Sheets("TABLA HC").Range("M" & linea3).Value
    Me.GinV.Value = Sheets("TABLA HC").Range("N" & linea3).Value
    Me.GinM.Value = Sheets("TABLA HC").Range("O" & linea3).Value
    
    Me.EmbarazoActual.Value = Sheets("TABLA CERTIFICADOS").Range("K" & linea2).Value
    
    Me.EnfCual.Value = Sheets("TABLA HC").Range("Q" & linea3).Value
    Me.AntCual.Value = Sheets("TABLA HC").Range("P" & linea3).Value
    Me.DiscCual.Value = Sheets("TABLA HC").Range("R" & linea3).Value
    
    Me.FactRiesgos.Value = Sheets("TABLA CERTIFICADOS").Range("L" & linea2).Value
    
    Me.TA.Value = Sheets("TABLA CERTIFICADOS").Range("M" & linea2).Value
    Me.Pulso.Value = Sheets("TABLA CERTIFICADOS").Range("N" & linea2).Value
    Me.FR.Value = Sheets("TABLA CERTIFICADOS").Range("O" & linea2).Value
    Me.Temp.Value = Sheets("TABLA CERTIFICADOS").Range("P" & linea2).Value
    Me.Spo2.Value = Sheets("TABLA CERTIFICADOS").Range("Q" & linea2).Value
    Me.Peso.Value = Sheets("TABLA CERTIFICADOS").Range("R" & linea2).Value
    Me.Talla.Value = Sheets("TABLA CERTIFICADOS").Range("S" & linea2).Value
    Me.Imc.Value = Sheets("TABLA CERTIFICADOS").Range("T" & linea2).Value
    Me.OB1.Value = Sheets("TABLA CERTIFICADOS").Range("U" & linea2).Value
    Me.OB2.Value = Sheets("TABLA CERTIFICADOS").Range("V" & linea2).Value
    Me.OB3.Value = Sheets("TABLA CERTIFICADOS").Range("W" & linea2).Value
    Me.OB4.Value = Sheets("TABLA CERTIFICADOS").Range("X" & linea2).Value
    Me.OB5.Value = Sheets("TABLA CERTIFICADOS").Range("Y" & linea2).Value
    Me.OB6.Value = Sheets("TABLA CERTIFICADOS").Range("Z" & linea2).Value
    Me.OB7.Value = Sheets("TABLA CERTIFICADOS").Range("AA" & linea2).Value
    Me.OB8.Value = Sheets("TABLA CERTIFICADOS").Range("AB" & linea2).Value
    Me.OB9.Value = Sheets("TABLA CERTIFICADOS").Range("AC" & linea2).Value
    Me.OB10.Value = Sheets("TABLA CERTIFICADOS").Range("AD" & linea2).Value
    Me.OB11.Value = Sheets("TABLA CERTIFICADOS").Range("AE" & linea2).Value
    Me.OB12.Value = Sheets("TABLA CERTIFICADOS").Range("AF" & linea2).Value
    Me.OB13.Value = Sheets("TABLA CERTIFICADOS").Range("AG" & linea2).Value
    Me.OB14.Value = Sheets("TABLA CERTIFICADOS").Range("AH" & linea2).Value
    Me.OB15.Value = Sheets("TABLA CERTIFICADOS").Range("AI" & linea2).Value
    Me.OB16.Value = Sheets("TABLA CERTIFICADOS").Range("AJ" & linea2).Value
    Me.OB17.Value = Sheets("TABLA CERTIFICADOS").Range("AK" & linea2).Value
    Me.OB18.Value = Sheets("TABLA CERTIFICADOS").Range("AL" & linea2).Value
    
    Me.ProcedimientosRealizados.Value = Sheets("TABLA CERTIFICADOS").Range("AS" & linea2).Value
    
    Me.Diag1.Value = Sheets("TABLA CERTIFICADOS").Range("AM" & linea2).Value
    Me.Diag2.Value = Sheets("TABLA CERTIFICADOS").Range("AN" & linea2).Value
    Me.Diag3.Value = Sheets("TABLA CERTIFICADOS").Range("AO" & linea2).Value
    
    Me.ConceptoML.Value = Sheets("TABLA CERTIFICADOS").Range("AP" & linea2).Value
    Me.DiagnosticoLaboral.Value = Sheets("TABLA CERTIFICADOS").Range("AQ" & linea2).Value
    Me.RecomMedicas.Value = Sheets("TABLA CERTIFICADOS").Range("AR" & linea2).Value
    Me.Restricciones.Value = Sheets("TABLA CERTIFICADOS").Range("AT" & linea2).Value
End Sub
