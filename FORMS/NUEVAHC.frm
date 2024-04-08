VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NUEVAHC 
   Caption         =   "Generar nueva historia clínica"
   ClientHeight    =   11592
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   22800
   OleObjectBlob   =   "NUEVAHC.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "NUEVAHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AntSi_Click()
    Me.AntCual.Enabled = True
    Me.AntCual.Visible = True
    Me.Label69.Visible = True
End Sub

Private Sub BotonAsignar_Click()
    ' Verifica si se ha seleccionado una fila en la ListaCIE10
    If ListaCIE10.ListIndex = -1 Then
        MsgBox "Por favor, seleccione un diagnóstico antes de asignarlo.", vbExclamation
        Exit Sub
    End If

    ' Obtiene el valor correspondiente
    Dim valor As Variant
    valor = Me.ListaCIE10.List(ListaCIE10.ListIndex, 1)

    ' Asigna el valor al primer control Diag vacío que encuentre
    If Me.Diag1.Value = "" Then
        Me.Diag1.Value = valor
    ElseIf Me.Diag2.Value = "" Then
        Me.Diag2.Value = valor
    ElseIf Me.Diag3.Value = "" Then
        Me.Diag3.Value = valor
    End If
End Sub


Private Sub BotonGenerarCertificdado_Click()
    Unload Me
    CERTIFICADO.Show
End Sub

Private Sub BotonGuardar_Click()
    Sheets("BASE DE DATOS 2024").Select
    
    Dim ultimaFilaP As Long
    
    ' Encontrar la última fila ocupada en la columna A
    ultimaFilaP = Sheets("BASE DE DATOS 2024").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Insertar una nueva fila después de la última fila ocupada
    Sheets("BASE DE DATOS 2024").Rows(ultimaFilaP + 1).Insert Shift:=xlDown

    Range("A" & ultimaFilaP + 1).EntireRow.Insert
    
    If ultimaFilaP = 2 Then
        Range("A" & ultimaFilaP + 1).Value = 1
    ElseIf ultimaFilaP > 2 Then
        Range("A" & ultimaFilaP + 1).Value = Range("A" & ultimaFilaP).Value
    End If
    
    Range("AG" & ultimaFilaP + 1).Value = Me.RutaImg.Value
    Range("B" & ultimaFilaP + 1).Value = Me.PrimerNombre.Value
    Range("C" & ultimaFilaP + 1).Value = Me.SegundoNombre.Value
    Range("D" & ultimaFilaP + 1).Value = Me.PrimerApellido.Value
    Range("E" & ultimaFilaP + 1).Value = Me.SegundoApellido.Value
    Range("F" & ultimaFilaP + 1).Value = Me.EstadoCivil.Value
    Range("G" & ultimaFilaP + 1).Value = Me.TipoDocumento.Value
    Range("H" & ultimaFilaP + 1).Value = Me.NDocumento.Value
    Range("I" & ultimaFilaP + 1).Value = Me.DepExpedicion.Value
    Range("J" & ultimaFilaP + 1).Value = Me.MunExpedicion.Value
    Range("K" & ultimaFilaP + 1).Value = Me.FechaExpedicion.Value
    Range("L" & ultimaFilaP + 1).Value = Me.DeptoNacimiento.Value
    Range("M" & ultimaFilaP + 1).Value = Me.MunNacimiento.Value
    Range("N" & ultimaFilaP + 1).Value = Me.FechaNacimiento.Value
    Range("O" & ultimaFilaP + 1).Value = Me.UnidadMedida.Value
    Range("P" & ultimaFilaP + 1).Value = Me.Edad.Value
    
    If Me.SexoF.Value = True Then
        Range("Q" & ultimaFilaP + 1).Value = "F"
    ElseIf Me.SexoM.Value = True Then
        Range("Q" & ultimaFilaP + 1).Value = "M"
    End If
    
    Range("R" & ultimaFilaP + 1).Value = Me.OcupacionActual.Value
    Range("S" & ultimaFilaP + 1).Value = Me.Direccion.Value
    Range("T" & ultimaFilaP + 1).Value = Me.Telefono.Value
    Range("U" & ultimaFilaP + 1).Value = Me.Rh.Value
    Range("V" & ultimaFilaP + 1).Value = Me.DeptoResidencia.Value
    Range("W" & ultimaFilaP + 1).Value = Me.MunResidencia.Value
    
    If Me.Rural.Value = True Then
        Range("X" & ultimaFilaP + 1).Value = "Rural"
    ElseIf Me.Urbano.Value = True Then
        Range("X" & ultimaFilaP + 1).Value = "Urbana"
    End If
    
    Range("Y" & ultimaFilaP + 1).Value = Me.GrupoAE.Value
    Range("Z" & ultimaFilaP + 1).Value = Me.ARL.Value
    Range("AA" & ultimaFilaP + 1).Value = Me.Pensiones.Value
    Range("AB" & ultimaFilaP + 1).Value = Me.Aseguradora.Value
    
    If Me.Contributivo.Value = True Then
        Range("AC" & ultimaFilaP + 1).Value = "Contributivo"
    ElseIf Me.Subsidiado.Value = True Then
        Range("AC" & ultimaFilaP + 1).Value = "Subsidiado"
    End If
    
    Range("AD" & ultimaFilaP + 1).Value = Me.NombreAcudiente.Value
    Range("AE" & ultimaFilaP + 1).Value = Me.Parentezco.Value
    Range("AF" & ultimaFilaP + 1).Value = Me.TelefonoAcudiente.Value
    
    Dim contador As Integer
    contador = 1
    Numerodedatos = Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 3 To Numerodedatos
        Range("A" & i).Value = contador
        contador = contador + 1
    Next
    
    Sheets("TABLA CERTIFICADOS").Select
    
    Dim ultimaFilaC As Long
    
    ' Encontrar la última fila ocupada en la columna A
    ultimaFilaC = Sheets("TABLA CERTIFICADOS").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Insertar una nueva fila después de la última fila ocupada
    Sheets("TABLA CERTIFICADOS").Rows(ultimaFilaC + 1).Insert Shift:=xlDown
    
    Range("A" & ultimaFilaC + 1).Value = Sheets("BASE DE DATOS 2024").Range("A" & ultimaFilaP + 1).Value
    
    Range("C" & ultimaFilaC + 1).Value = "CSST-" & Year(Date) & "-00" & Range("B1").Value
    
    Range("D" & ultimaFilaC + 1).Value = Me.CargoAOcupar.Value
    Range("E" & ultimaFilaC + 1).Value = Me.Entidad.Value
    Range("G" & ultimaFilaC + 1).Value = Me.FechaIngreso.Value
    Range("H" & ultimaFilaC + 1).Value = Me.FechaAtencion.Value
    Range("I" & ultimaFilaC + 1).Value = Me.MunAtencion.Value & ", " & Me.DeptoAtencion.Value
    Range("J" & ultimaFilaC + 1).Value = Me.TipoConsulta.Value
    
    If Me.EmbSi.Value = True Then
        Range("K" & ultimaFilaC + 1).Value = "Si"
    ElseIf Me.EmbNo.Value = True Then
        Range("K" & ultimaFilaC + 1).Value = "No"
    End If
    
    Range("L" & ultimaFilaC + 1).Value = Me.FactRiesgos.Value
    
    Range("M" & ultimaFilaC + 1).Value = Me.TA.Value
    Range("N" & ultimaFilaC + 1).Value = Me.Pulso.Value
    Range("O" & ultimaFilaC + 1).Value = Me.FR.Value
    Range("P" & ultimaFilaC + 1).Value = Me.Temp.Value
    Range("Q" & ultimaFilaC + 1).Value = Me.Spo2.Value
    Range("R" & ultimaFilaC + 1).Value = Me.Peso.Value
    Range("S" & ultimaFilaC + 1).Value = Me.Talla.Value
    Range("T" & ultimaFilaC + 1).Value = Me.Imc.Value
    
    Range("U" & ultimaFilaC + 1).Value = Me.OB1.Value
    Range("V" & ultimaFilaC + 1).Value = Me.OB2.Value
    Range("W" & ultimaFilaC + 1).Value = Me.OB3.Value
    Range("X" & ultimaFilaC + 1).Value = Me.OB4.Value
    Range("Y" & ultimaFilaC + 1).Value = Me.OB5.Value
    Range("Z" & ultimaFilaC + 1).Value = Me.OB6.Value
    Range("AA" & ultimaFilaC + 1).Value = Me.OB7.Value
    Range("AB" & ultimaFilaC + 1).Value = Me.OB8.Value
    Range("AC" & ultimaFilaC + 1).Value = Me.OB9.Value
    Range("AD" & ultimaFilaC + 1).Value = Me.OB10.Value
    Range("AE" & ultimaFilaC + 1).Value = Me.OB11.Value
    Range("AF" & ultimaFilaC + 1).Value = Me.OB12.Value
    Range("AG" & ultimaFilaC + 1).Value = Me.OB13.Value
    Range("AH" & ultimaFilaC + 1).Value = Me.OB14.Value
    Range("AI" & ultimaFilaC + 1).Value = Me.OB15.Value
    Range("AJ" & ultimaFilaC + 1).Value = Me.OB16.Value
    Range("AK" & ultimaFilaC + 1).Value = Me.OB17.Value
    Range("AL" & ultimaFilaC + 1).Value = Me.OB18.Value
    
    Range("AM" & ultimaFilaC + 1).Value = Me.Diag1.Value
    Range("AN" & ultimaFilaC + 1).Value = Me.Diag2.Value
    Range("AO" & ultimaFilaC + 1).Value = Me.Diag3.Value
    
    Range("AP" & ultimaFilaC + 1).Value = Me.ConceptoML.Value
    
    Range("AQ" & ultimaFilaC + 1).Value = Me.DiagnosticoLaboral.Value
    
    Sheets("TABLA HC").Select
    
    Dim j As Integer
    Dim result As String
    
    For j = 1 To 9 ' Supongo que los controles se llaman Rec1, Rec2, ..., Rec9
        If Me.Controls("Rec" & i).Value = True Then
            If result <> "" Then
                result = result & vbCrLf
            End If
            result = result & Me.Controls("Rec" & j).Caption ' Utiliza Caption para obtener el texto del checkbox
        End If
    Next j
    
    
    Sheets("TABLA CERTIFICADOS").Select
    
    Range("AR" & ultimaFilaC + 1).Value = result
    
    If Me.RecOtro.Value <> "" Then
        Range("AR" & ultimaFilaC + 1).Value = Range("AR" & ultimaFilaC + 1).Value & vbCrLf & Me.RecOtro.Value
    End If
    
    
    Range("AS" & ultimaFilaC + 1).Value = Me.ProcedimientosRealizados.Value
    
    Sheets("TABLA HC").Select
    
   Dim ultimaFilaH As Long

    ' Encontrar la última fila ocupada en cualquier columna
    ultimaFilaH = Sheets("TABLA HC").Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Insertar una nueva fila después de la última fila ocupada
    Sheets("TABLA HC").Rows(ultimaFilaH + 1).Insert Shift:=xlDown
    
    Range("C" & ultimaFilaH + 1).Value = 1
    Range("B" & ultimaFilaH + 1).Value = Sheets("BASE DE DATOS 2024").Range("A" & ultimaFilaP + 1).Value
    Range("D" & ultimaFilaH + 1).Value = Date


    Range("E" & ultimaFilaH + 1).Value = Me.AntFamiliares.Value
    Range("F" & ultimaFilaH + 1).Value = Me.AntPatologicos.Value
    Range("G" & ultimaFilaH + 1).Value = Me.AntFarmacologicos.Value
    Range("H" & ultimaFilaH + 1).Value = Me.AntQuirurgicos.Value
    Range("I" & ultimaFilaH + 1).Value = Me.AntTox.Value
    Range("J" & ultimaFilaH + 1).Value = Me.GinG.Value
    Range("K" & ultimaFilaH + 1).Value = Me.GinP.Value
    Range("L" & ultimaFilaH + 1).Value = Me.GinC.Value
    Range("M" & ultimaFilaH + 1).Value = Me.GinA.Value
    Range("N" & ultimaFilaH + 1).Value = Me.GinV.Value
    Range("O" & ultimaFilaH + 1).Value = Me.GinM.Value
    
    
    If Me.AntSi.Value = True Then
        Range("P" & ultimaFilaH + 1).Value = Me.AntCual.Value
    ElseIf Me.AntNo.Value = True Then
        Range("P" & ultimaFilaH + 1).Value = "Negativo"
    End If
    
    If Me.EnfSi.Value = True Then
        Range("Q" & ultimaFilaH + 1).Value = Me.EnfCual.Value
    ElseIf Me.EnfNo.Value = True Then
        Range("Q" & ultimaFilaH + 1).Value = "Negativo"
    End If
    
    If Me.DiscSi.Value = True Then
        Range("R" & ultimaFilaH + 1).Value = Me.DiscCual.Value
    ElseIf Me.DiscNo.Value = True Then
        Range("R" & ultimaFilaH + 1).Value = "Negativo"
    End If
    
    MsgBox ("Información del paciente guardada correctamente.")
    Advertencia = MsgBox("¿Ver base de datos?", vbYesNo + vbQuestion, "Confirmar")
    
    If Advertencia = vbYes Then
        Sheets("BASE DE DATOS 2024").Select
        Unload Me
    End If
    
End Sub

Private Sub BotonLimpiar_Click()
    Me.Diag1.Value = ""
    Me.Diag2.Value = ""
    Me.Diag3.Value = ""
End Sub

Private Sub BotonLimpiar2_Click()
    Me.DiagnosticoLaboral.Value = ""
End Sub

Private Sub BotonSalir_Click()
    Unload Me
End Sub

Private Sub BotonSeleccionarImg_Click()
    Set explorar_archivo = Application.FileDialog(msoFileDialogFilePicker)
    explorar_archivo.Title = "Seleccionar foto del equipo"
    explorar_archivo.AllowMultiSelect = False
    explorar_archivo.Show
    
    ruta_imagen = explorar_archivo.SelectedItems(1)
    Me.RutaImg.Value = ruta_imagen
    Me.Imagen1.Picture = LoadPicture(ruta_imagen)
    
End Sub

Private Sub BotonVerEscala_Click()
 IMC_FORM.Show
End Sub

Private Sub BotonVolver_Click()
    Unload Me
    INICIO.Show
End Sub

Private Sub ActualizarMunicipios(comboDepto As MSForms.ComboBox, comboMun As MSForms.ComboBox)
    
    ' Limpiar la lista actual
    comboMun.Clear
    
    ' Obtener el departamento seleccionado
    Dim deptoSeleccionado As String
    deptoSeleccionado = comboDepto.Value
    
    ' Agregar lógica para cargar los municipios según el departamento seleccionado
    If comboDepto.Value = "Amazonas" Then
    comboMun.List = Range("E2:E12").Value
    ElseIf comboDepto.Value = "Antioquia" Then
        comboMun.List = Range("E13:E137").Value
    ElseIf comboDepto.Value = "Arauca" Then
        comboMun.List = Range("E138:E144").Value
    ElseIf comboDepto.Value = "Atlántico" Then
        comboMun.List = Range("E147:E169").Value
    ElseIf comboDepto.Value = "Bogotá D.C." Then
        comboMun.Clear
        comboMun.AddItem "Bogotá D.C."
    ElseIf comboDepto.Value = "Bolívar" Then
        comboMun.List = Range("E171:E215").Value
    ElseIf comboDepto.Value = "Boyacá" Then
        comboMun.List = Range("E216:E338").Value
    ElseIf comboDepto.Value = "Caldas" Then
        comboMun.List = Range("E339:E365").Value
    ElseIf comboDepto.Value = "Caquetá" Then
        comboMun.List = Range("E366:E381").Value
    ElseIf comboDepto.Value = "Casanare" Then
        comboMun.List = Range("E382:E400").Value
    ElseIf comboDepto.Value = "Cauca" Then
        comboMun.List = Range("E401:E441").Value
    ElseIf comboDepto.Value = "Cesar" Then
        comboMun.List = Range("E442:E466").Value
    ElseIf comboDepto.Value = "Chocó" Then
        comboMun.List = Range("E467:E497").Value
    ElseIf comboDepto.Value = "Córdoba" Then
        comboMun.List = Range("E498:E525").Value
    ElseIf comboDepto.Value = "Cundinamarca" Then
        comboMun.List = Range("E526:E641").Value
    ElseIf comboDepto.Value = "Guainía" Then
        comboMun.List = Range("E642:E650").Value
    ElseIf comboDepto.Value = "Guaviare" Then
        comboMun.List = Range("E651:E654").Value
    ElseIf comboDepto.Value = "Huila" Then
        comboMun.List = Range("E655:E691").Value
    ElseIf comboDepto.Value = "La Guajira" Then
        comboMun.List = Range("E692:E706").Value
    ElseIf comboDepto.Value = "Magdalena" Then
        comboMun.List = Range("E707:E736").Value
    ElseIf comboDepto.Value = "Meta" Then
        comboMun.List = Range("E737:E762").Value
    ElseIf comboDepto.Value = "Nariño" Then
        comboMun.List = Range("E766:E829").Value
    ElseIf comboDepto.Value = "Norte de Santander" Then
        comboMun.List = Range("E830:E869").Value
    ElseIf comboDepto.Value = "Putumayo" Then
        comboMun.List = Range("E870:E882").Value
    ElseIf comboDepto.Value = "Quindío" Then
        comboMun.List = Range("E883:E894").Value
    ElseIf comboDepto.Value = "Risaralda" Then
        comboMun.List = Range("E895:E908").Value
    ElseIf comboDepto.Value = "Archipiélago de San Andrés" Then
        comboMun.Clear
        comboMun.AddItem "San Andrés y Providencia"
    ElseIf comboDepto.Value = "Santander" Then
        comboMun.List = Range("E909:E995").Value
    ElseIf comboDepto.Value = "Sucre" Then
        comboMun.List = Range("E996:E1021").Value
    ElseIf comboDepto.Value = "Tolima" Then
        comboMun.List = Range("E1022:E1068").Value
    ElseIf comboDepto.Value = "Valle del Cauca" Then
        comboMun.List = Range("E1069:E1110").Value
    ElseIf comboDepto.Value = "Vaupés" Then
        comboMun.List = Range("E1111:E1116").Value
    ElseIf comboDepto.Value = "Vichada" Then
        comboMun.List = Range("E1117:E1120").Value
    End If
End Sub

Private Sub BuscarCIE10_Click()
    Sheets("CIE10").Select
    Me.ListaCIE10 = Clear
    Me.ListaCIE10.RowSource = Clear
    
    Y = 0
    
    For fila = 7 To 12430
        ' Buscar en ambas columnas C y D
        nombreC = ActiveSheet.Cells(fila, 3).Value
        nombreD = ActiveSheet.Cells(fila, 4).Value
        
        If UCase(nombreC) Like "*" & UCase(Me.BuscadorDiag.Value) & "*" Or _
           UCase(nombreD) Like "*" & UCase(Me.BuscadorDiag.Value) & "*" Then
            Me.ListaCIE10.AddItem
            Me.ListaCIE10.List(Y, 0) = nombreC
            Me.ListaCIE10.List(Y, 1) = nombreD
            Y = Y + 1
        End If
    Next
End Sub


Private Sub BuscarLabs_Click()
    Sheets("ENFERMEDADES LABORALES").Select
    Me.ListaDiag = Clear
    Me.ListaDiag.RowSource = Clear
    
    Y = 0
    
    For fila = 5 To 352
        ' Buscar en ambas columnas C y D
        nombreA = ActiveSheet.Cells(fila, 1).Value
        nombreB = ActiveSheet.Cells(fila, 2).Value
        
        If UCase(nombreA) Like "*" & UCase(Me.BuscadorLab.Value) & "*" Or _
           UCase(nombreB) Like "*" & UCase(Me.BuscadorLab.Value) & "*" Then
            Me.ListaDiag.AddItem
            Me.ListaDiag.List(Y, 0) = nombreA
            Me.ListaDiag.List(Y, 1) = nombreB
            Y = Y + 1
        End If
    Next
End Sub


Private Sub DepExpedicion_Change()
    Sheets("TABLA REGIONES").Select
    ActualizarMunicipios Me.DepExpedicion, Me.MunExpedicion
End Sub

Private Sub DeptoAtencion_Change()
    Sheets("TABLA REGIONES").Select
    ActualizarMunicipios Me.DeptoAtencion, Me.MunAtencion
End Sub

Private Sub DeptoNacimiento_Change()
    Sheets("TABLA REGIONES").Select
    ActualizarMunicipios Me.DeptoNacimiento, Me.MunNacimiento
End Sub

Private Sub DeptoResidencia_Change()
    Sheets("TABLA REGIONES").Select
    ActualizarMunicipios Me.DeptoResidencia, Me.MunResidencia
End Sub

Private Sub DiscSi_Click()
    Me.DiscCual.Enabled = True
    Me.DiscCual.Visible = True
    Me.Label74.Visible = True
End Sub

Private Sub EnfSi_Click()
    Me.EnfCual.Enabled = True
    Me.EnfCual.Visible = True
    Me.Label71.Visible = True
End Sub

Private Sub FechaAtencion_AfterUpdate()
    ' Verificar el formato de la fecha
    If Not IsDate(Me.FechaAtencion.Value) And Me.FechaAtencion.Value <> "" Then
        MsgBox "Formato de fecha no válido. Utilice DD/MM/AAAA.", vbExclamation
        Me.FechaAtencion.Value = ""
        Exit Sub
    End If
End Sub

Private Sub FechaExpedicion_AfterUpdate()
    ' Verificar el formato de la fecha
    If Not IsDate(Me.FechaExpedicion.Value) And Me.FechaExpedicion.Value <> "" Then
        MsgBox "Formato de fecha no válido. Utilice DD/MM/AAAA.", vbExclamation
        Me.FechaExpedicion.Value = ""
        Exit Sub
    End If
End Sub


Private Sub FechaIngreso_AfterUpdate()
    ' Verificar el formato de la fecha
    If Not IsDate(Me.FechaIngreso.Value) And Me.FechaIngreso.Value <> "" Then
        MsgBox "Formato de fecha no válido. Utilice DD/MM/AAAA.", vbExclamation
        Me.FechaIngreso.Value = ""
        Exit Sub
    End If
End Sub

Private Sub FechaNacimiento_AfterUpdate()
    ' Verificar el formato de la fecha
    If Not IsDate(Me.FechaNacimiento.Value) And Me.FechaNacimiento.Value <> "" Then
        MsgBox "Formato de fecha no válido. Utilice DD/MM/AAAA.", vbExclamation
        Me.FechaNacimiento.Value = ""
        Exit Sub
    End If
    
    ' Convertir la fecha al formato deseado (DD/MM/AAAA)
    Me.FechaNacimiento.Value = Format(Me.FechaNacimiento.Value, "DD/MM/YYYY")
    
    ' Calcular la edad
    Me.Edad.Value = CalcularEdad(Me.FechaNacimiento.Value)
End Sub

Private Function CalcularEdad(FechaNacimiento As Date) As Integer
    ' Calcular la edad en años
    Dim Edad As Integer
    Edad = DateDiff("yyyy", FechaNacimiento, Date)
    
    ' Ajustar si aún no ha llegado el cumpleaños este año
    If Date < DateSerial(Year(Date), Month(FechaNacimiento), Day(FechaNacimiento)) Then
        Edad = Edad - 1
    End If
    
    ' Devolver la edad calculada
    CalcularEdad = Edad
End Function



Private Sub ListaCIE10_Click()
    codigo = Me.ListaCIE10.List(ListaCIE10.ListIndex, 0)
End Sub



Private Sub ListaDiag_Click()
    valorb = Me.ListaDiag.List(ListaDiag.ListIndex, 1)
    Me.DiagnosticoLaboral.Value = valorb
End Sub

Private Sub Talla_Change()
    Dim Talla As Double
    Dim Peso As Double
    Dim Imc As Double

    ' Obtener los valores de Talla y Peso
    Talla = Val(Me.Talla.Value)
    Peso = Val(Me.Peso.Value)

    ' Verificar si los valores son numéricos y mayores que cero
    If IsNumeric(Talla) And IsNumeric(Peso) And Talla > 0 And Peso > 0 Then
        ' Calcular el IMC y redondear a 5 decimales
        Imc = Round(Peso / ((Talla / 100) ^ 2), 5)
        Me.Imc.Value = Imc
    Else
        Me.Imc.Value = ""
    End If
    
    ' Cambiar el color de fondo según el valor del IMC
    If IsNumeric(Imc) Then
        If Imc < 18.5 Then
            Me.Imc.BackColor = RGB(243, 224, 98)
        ElseIf Imc >= 18.5 And Imc < 25 Then
            Me.Imc.BackColor = RGB(106, 224, 113)
        ElseIf Imc >= 25 And Imc < 30 Then
            Me.Imc.BackColor = RGB(244, 157, 93)
        ElseIf Imc >= 30 Then
            Me.Imc.BackColor = RGB(242, 104, 96)
        End If
    Else
        Me.Imc.BackColor = RGB(255, 255, 255)
    End If
    
End Sub


Private Sub UserForm_Initialize()
    With Me
    .Width = Application.Width
    .Height = Application.Height
    End With
    
    Me.OB1.Value = "NORMOCEFALO, CUERO CABELLUDO NORMOIMPLANTADO, SIN DEFORMIDADES NO SE OBSERVAN  LESIONES NI CICATRICES."
    Me.OB2.Value = "PUPILAS ISOCORICAS NORMORREACTIVAS A LA LUZ, CONJUNTIVAS ROSADAS SIN LESIONES, ESCLERAS SIN LESIONES, ANICTÉRICAS, AGUDEZA VISUAL OI / OD /"
    Me.OB3.Value = "PABELLONES AURICULARES TAMAÑO Y FORMA NORMALES, NORMOIMPLANTADOS. CONDUCTOS AUDITIVOS EXTERNOS PERMEABLES"
    Me.OB4.Value = "PERMEABLE, NO DESVIACIÓN DEL TABIQUE NASAL, NO RINORREA NI ESTIGMAS DE EPISTAXIS"
    Me.OB5.Value = "LABIOS ROSADOS HÚMEDOS, NO MALFORMACIONES CONGENITAS, COMISURAS LABIALES SIN LESIONES MUCOSA HUMEDA ENCIAS NO LESIONES, ESTRUCTURAS DENTOMAXILARES, LENGUA SIN ALTERACION, PALADAR DURO BLANQUECINO SIN LESION, UVULA SIN LESION, PALADAR  BLANDO ROSADO SIN LESIÓN."
    Me.OB6.Value = "FARINGE Y AMIGDALAS SIN LESIÓN"
    Me.OB7.Value = "CENTRAL NO MASAS NI ADENOPATIAS, NO INGURGITACIÓN YUGULAR, NO SOPLOS CAROTIDEOS."
    Me.OB8.Value = "SIMETRICO NORMOEXPLANSIBLE, NO SE OBSERVAN SIGNOS DE TIRAJE INTER NI SUBCOSTAL. EN MUJERES DESCRIBIR CARACTERISITICAS DE LAS MAMAS"
    Me.OB9.Value = "RITMICO, NO SE AUSCULTAN SOPLOS"
    Me.OB10.Value = "CAMPOS PULMONARES BIEN VENTILADOS, MURMULLO VESICULAR CONSERVADO, SIN PRESENCIA DE RUIDOS AGREGADOS"
    Me.OB11.Value = "BLANDO DEPRESIBLE, NO MASAS, NI MEGALIAS, PERISTALTISMO POSITIVO, NORMAL EN INTENSIDAD Y FRECUENCIA, NO HERNIAS, NO SIGNOS DE IRRITACION PERITONEAL, NO SOPLOS PERIUMBILICALES."
    Me.OB12.Value = "SIN ALTERACIÓN"
    Me.OB13.Value = "PUNTOS PIELOURETERALES ANTERIORES NEGATIVOS, PUÑO PERCUSIÓN NEGATIVA. GENITALES NORMOCONFIGURADOS."
    Me.OB14.Value = "MÓVILES, SIMÉTRICAS, NO EDEMAS SENSIBILIDAD CONSERVADA, NO ALTERACION MUSCULOESQUELETICA, PULSOS PERIFÉRICOS PRESENTES."
    Me.OB15.Value = "MÓVILES, SIMÉTRICAS, NO EDEMAS SENSIBILIDAD CONSERVADA, NO ALTERACION MUSCULOESQUELETICA, REFLEJOS ROTULIANOS Y AQUILEANOS CONSERVADOS, PULSOS FEMOREALES PEDIOS POPLITEOS PALPABLES, NO SIGNOS DE INSUFICIENCIA VENOSA O ARTERIAL."
    Me.OB16.Value = "SIMETRICA, SIN DESVIACIÓN, NO DOLOR A LA PALPACIÓN, NO CONTRACTURA"
    Me.OB17.Value = "NORMOELASTICA NORMOTENSA NORMOTERMICA, HIDRATADA (HACER ENFASIS EN ESTADO DE HIDRATACIÓN), NO LESIONES, NO DESCAMACIONES, LLENADO CAPILAR MENOR A 2 SEGUNDOS"
    Me.OB18.Value = "SIN DEFICIT APARENTE, UBICADA EN TIEMPO, LUGAR Y PERSONA, NO HAY DEFICIT MOTOR NI SENSITIVO, NO SIGNOS MENINGEOS NI DE FOCALIZACION, LENGUAJE ESPONTANEO MARCHA NORMAL, TROFISMO MUSCULAR, TONO MUSCULAR NORMAL     DISMINUIDO. EXAMEN MENTAL: CONCIENTE, ORIENTADO."
    
    Sheets("OTROS").Select
    Me.EstadoCivil.List = Range("B1:B6").Value
    Me.UnidadMedida.List = Range("C1:C3").Value
    Me.TipoDocumento.List = Range("A1:A6").Value
    Me.TipoConsulta.List = Range("D1:D3").Value
    Me.ConceptoML.List = Range("F1:F5").Value
    
    Sheets("TABLA REGIONES").Select
    Me.DeptoNacimiento.List = Range("G3:G35").Value
    Me.DepExpedicion.List = Range("G3:G35").Value
    Me.DeptoAtencion.List = Range("G3:G35").Value
    Me.DeptoResidencia.List = Range("G3:G35").Value
    
    Sheets("ARL").Select
    Me.ARL.List = Range("D3:D12").Value
    
    Sheets("PENSIONES").Select
    Me.Pensiones.List = Range("D11:D130").Value
    
    Sheets("ASEGURADORAS").Select
    Me.Aseguradora.List = Range("B2:B30").Value
    
    Sheets("POBLACIONES ESPECIALES").Select
    Me.GrupoAE.List = Range("A2:A18").Value
    
  
    
    Me.ListaCIE10.RowSource = "TABLACIE10"
    Me.ListaCIE10.ColumnCount = 2
    
    Me.ListaDiag.RowSource = "ENFERMEDADES_LABORALES"
    Me.ListaDiag.ColumnCount = 2
     
End Sub
