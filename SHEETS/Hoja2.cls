VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim contador As Integer
    Dim Numerodedatos As Long

    Application.EnableEvents = False ' Desactivar eventos para evitar bucles infinitos
    
    contador = 1
    Numerodedatos = Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 3 To Numerodedatos
        Range("A" & i).Value = contador
        contador = contador + 1
    Next
    
    Application.EnableEvents = True ' Volver a activar eventos
End Sub

