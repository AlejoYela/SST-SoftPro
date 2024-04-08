VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NUEVOFOLIO 
   Caption         =   "Agregar nuevo folio a Historia Clínica existente"
   ClientHeight    =   11592
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   22800
   OleObjectBlob   =   "NUEVOFOLIO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "NUEVOFOLIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub BotonGuardar_Click()
    Dim i As Integer
    Dim result As String
    Dim folio As Integer
    
    folio = 1
    valor = Me.ListaPacientes.List(ListaPacientes.ListIndex, 0)

    Sheets("TABLA HC").Select
    
    Numerodedatos = Range("B" & Rows.Count).End(xlUp).Row
    
    For fila = 2 To Numerodedatos
        ' Buscar en ambas columnas
        nombreB = ActiveSheet.Cells(fila, 2).Value
        
        If nombreB Like valor Then
            ' Incrementar el contador de coincidencias
            folio = folio + 1
        End If
    Next

    
    For i = 1 To 9 ' Supongo que los controles se llaman Rec1, Rec2, ..., Rec9
        If Me.Controls("Rec" & i).Value = True Then
            If result <> "" Then
                result = result & vbCrLf
            End If
            result = result & Me.Controls("Rec" & i).Caption ' Utiliza Caption para obtener el texto del checkbox
        End If
    Next i
    
    If Me.RecOtro.Value <> "" Then
        Sheets("OTROS").Range("H2").Value = Sheets("OTROS").Range("H2").Value & vbCrLf & Me.RecOtro.Value
    End If
    
    
    Sheets("OTROS").Range("I2").Value = Me.ProcedimientosRealizados.Value

    Dim ultimaFilaH As Long

    ' Encontrar la última fila ocupada en cualquier columna
    ultimaFilaH = Sheets("TABLA HC").Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Insertar una nueva fila después de la última fila ocupada
    Sheets("TABLA HC").Rows(ultimaFilaH + 1).Insert Shift:=xlDown
    
    Range("B" & ultimaFilaH + 1).Value = Me.ListaPacientes.List(ListaPacientes.ListIndex, 0)
    Range("D" & ultimaFilaH + 1).Value = Date

    Sheets("OTROS").Range("H2").Value = result
    
    Range("C" & ultimaFilaH + 1).Value = folio

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
    Advertencia = MsgBox("¿Ver base de datos de historias clínicas?", vbYesNo + vbQuestion, "Confirmar")
    
    If Advertencia = vbYes Then
        Sheets("TABLA HC").Select
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

Private Sub BotonVerEscala_Click()
    IMC_FORM.Show
End Sub

Private Sub BotonVolver_Click()
    Unload Me
    ACTUALIZARHC.Show
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

Private Sub DeptoAtencion_Change()
    Sheets("TABLA REGIONES").Select
    ActualizarMunicipios Me.DeptoAtencion, Me.MunAtencion
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
    If Not IsDate(Me.FechaAtencion.Value) And Me.FechaAtencion.Value <> "" Then
        MsgBox "Formato de fecha no válido. Utilice DD/MM/AAAA.", vbExclamation
        Me.FechaAtencion.Value = ""
        Exit Sub
    End If
End Sub

Private Sub ListaCIE10_Click()
    codigo = Me.ListaCIE10.List(ListaCIE10.ListIndex, 0)
End Sub

Private Sub ListaDiag_Click()
    valorb = Me.ListaDiag.List(ListaDiag.ListIndex, 1)
    Me.DiagnosticoLaboral.Value = valorb
End Sub

Private Sub ListaPacientes_Click()
    Me.Label6.Visible = True
    Me.Label7.Visible = True
    Me.NombresCompletos.Visible = True
    Me.Documento.Visible = True
    
    Dim linea As Integer
    Dim fila As Object
    
    
    IdHC = Me.ListaPacientes.List(ListaPacientes.ListIndex, 0)
    
    Set fila = Sheets("BASE DE DATOS 2024").Range("A:A").Find(IdHC, lookat:=xlWhole)
    linea = fila.Row
    
    Me.NombresCompletos.Value = Sheets("BASE DE DATOS 2024").Range("B" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("C" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("D" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("E" & linea).Value
    Me.Documento.Value = Sheets("BASE DE DATOS 2024").Range("G" & linea).Value & " " & Sheets("BASE DE DATOS 2024").Range("H" & linea).Value
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
    
    Me.ListaPacientes.RowSource = "DATABASE"
    Me.ListaPacientes.ColumnCount = 8
    
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
    Me.TipoConsulta.List = Range("D1:D3").Value
    Me.ConceptoML.List = Range("F1:F5").Value
    
    Sheets("TABLA REGIONES").Select
    Me.DeptoAtencion.List = Range("G3:G35").Value
    
    
    Me.ListaCIE10.RowSource = "TABLACIE10"
    Me.ListaCIE10.ColumnCount = 2
    
    Me.ListaDiag.RowSource = "ENFERMEDADES_LABORALES"
    Me.ListaDiag.ColumnCount = 2
End Sub
