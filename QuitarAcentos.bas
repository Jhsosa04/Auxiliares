Attribute VB_Name = "QuitarAcentos"
Option Explicit

' ============================================================
'  BOT QUITAR ACENTOS
'  Reemplaza letras acentuadas por su version sin acento:
'  a e i o u u n  (y sus mayusculas)
'  Ej.: Diaz con tilde -> Diaz | Nunez con enie -> Nunez
'  Alcance: la seleccion (si hay varias celdas seleccionadas)
'           o toda la hoja activa (si solo hay una celda).
'  Solo toca texto escrito, NO modifica formulas.
' ============================================================

Public Sub QuitarAcentos()
    Dim ws As Worksheet
    Dim zona As Range, celda As Range
    Dim conAcento As Variant, sinAcento As Variant
    Dim original As String, limpio As String
    Dim i As Long, cambiadas As Long

    ' Letras con acento (codigos Unicode para evitar problemas de codificacion)
    conAcento = Array(ChrW(225), ChrW(233), ChrW(237), ChrW(243), ChrW(250), ChrW(252), ChrW(241), _
                      ChrW(193), ChrW(201), ChrW(205), ChrW(211), ChrW(218), ChrW(220), ChrW(209))
    '                  a         e          i          o          u          u(dieresis) n
    '                  A         E          I          O          U          U(dieresis) N
    sinAcento = Array("a", "e", "i", "o", "u", "u", "n", _
                      "A", "E", "I", "O", "U", "U", "N")

    Set ws = ActiveSheet

    ' Definir zona de trabajo: solo celdas con texto (sin formulas)
    On Error Resume Next
    If TypeName(Selection) = "Range" And Selection.Cells.CountLarge > 1 Then
        Set zona = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    Else
        Set zona = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
    End If
    On Error GoTo 0

    If zona Is Nothing Then
        MsgBox "No hay celdas con texto para revisar.", vbExclamation, "Quitar acentos"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each celda In zona.Cells
        original = CStr(celda.Value)
        limpio = original
        For i = LBound(conAcento) To UBound(conAcento)
            If InStr(1, limpio, conAcento(i), vbBinaryCompare) > 0 Then
                limpio = Replace(limpio, conAcento(i), sinAcento(i), , , vbBinaryCompare)
            End If
        Next i
        If limpio <> original Then
            celda.Value = limpio
            cambiadas = cambiadas + 1
        End If
    Next celda

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Listo." & vbCrLf & "Celdas corregidas: " & cambiadas & vbCrLf & _
           "Celdas revisadas: " & zona.Cells.CountLarge, vbInformation, "Quitar acentos"
End Sub
