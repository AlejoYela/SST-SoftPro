VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} INICIO 
   Caption         =   "Inicio - SST SoftPro "
   ClientHeight    =   5736
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   8820.001
   OleObjectBlob   =   "INICIO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "INICIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---- Botones del Formulario ----

Private Sub BotonActualizarHC_Click()
    ' Manejador de eventos para el clic en el bot�n "Actualizar Historia Cl�nica"
    Unload Me   ' Descarga el formulario actual
    ACTUALIZARHC.Show   ' Muestra el formulario de actualizaci�n de historias cl�nicas
End Sub

Private Sub BotonGenerarCertificado_Click()
    ' Manejador de eventos para el clic en el bot�n "Generar Certificado"
    Unload Me   ' Descarga el formulario actual
    CERTIFICADO.Show   ' Muestra el formulario de generaci�n de certificados m�dicos
End Sub

Private Sub BotonNuevoPaciente_Click()
    ' Manejador de eventos para el clic en el bot�n "Nuevo Paciente"
    Unload Me   ' Descarga el formulario actual
    NUEVAHC.Show   ' Muestra el formulario para registrar a un nuevo paciente
End Sub

Private Sub BotonSalir_Click()
    ' Manejador de eventos para el clic en el bot�n "Salir"
    Unload Me   ' Descarga el formulario actual
    ' El formulario se cierra sin mostrar otro formulario (ya que no hay llamada a Show)
End Sub

' ---- Im�genes y MouseMove ----

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Manejador de eventos para el movimiento del mouse sobre el Frame1
    Me.Image2.SpecialEffect = fmSpecialEffectFlat   ' Efecto visual en la imagen2
    Me.Image4.SpecialEffect = fmSpecialEffectFlat   ' Efecto visual en la imagen4
End Sub

Private Sub Image2_Click()
    ' Manejador de eventos para el clic en la imagen2
    Unload Me   ' Descarga el formulario actual
    NUEVAHC.Show   ' Muestra el formulario para registrar a un nuevo paciente
End Sub

Private Sub Image2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Manejador de eventos para el movimiento del mouse sobre la imagen2
    Me.Image2.SpecialEffect = fmSpecialEffectEtched   ' Efecto visual en la imagen2
End Sub

Private Sub Image3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Manejador de eventos para el movimiento del mouse sobre la imagen3
    Me.Image3.SpecialEffect = fmSpecialEffectEtched   ' Efecto visual en la imagen3
End Sub

Private Sub Image4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Manejador de eventos para el movimiento del mouse sobre la imagen4
    Me.Image4.SpecialEffect = fmSpecialEffectEtched   ' Efecto visual en la imagen4
End Sub

Private Sub Image4_Click()
    ' Manejador de eventos para el clic en la imagen4
    Unload Me   ' Descarga el formulario actual
    ACTUALIZARHC.Show   ' Muestra el formulario de actualizaci�n de historias cl�nicas
End Sub

Private Sub UserForm_Click()
    ' Manejador de eventos para el clic en cualquier lugar del formulario (fuera de los controles)
    ' Puedes agregar c�digo aqu� si necesitas realizar alguna acci�n cuando se hace clic en el formulario
End Sub

