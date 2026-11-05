Attribute VB_Name = "MarcarCheques"
Option Explicit

' ============================================================
'  BOT CHECK LIST - ARCHIVAR BHD
'  Escribe o escanea el cheque -> lo busca -> lo marca en
'  amarillo -> limpia la casilla y espera el siguiente.
'  Para terminar: deja la casilla vacia o pulsa Cancelar.
' ============================================================

Private Const COLOR_MARCA As Long = 65535   ' amarillo (RGB 255,255,0)
Private Const COL_CODIGO As Long = 1        ' columna A = codigo BHD...

Public Sub MarcarCheques()
    Dim ws As Worksheet
    Dim celda As Range, objetivo As Range
    Dim codigo As String, estado As String
    Dim marcados As Long, coincidencias As Long

    Set ws = ActiveSheet
    estado = "Listo para empezar."

    Do
        codigo = InputBox( _
            "Introduce o escanea el cheque y pulsa Enter." & vbCrLf & _
            "Vacio o Cancelar = terminar." & vbCrLf & vbCrLf & _
            ">> " & estado, _
            "Check List BHD  |  Marcados en esta sesion: " & marcados)

        codigo = UCase$(Trim$(codigo))
        If codigo = "" Then Exit Do

        Set objetivo = Nothing

        ' 1) Busqueda exacta en columnas A y B
        Set celda = ws.Range("A:B").Find(What:=codigo, LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)

        If Not celda Is Nothing Then
            Set objetivo = ws.Cells(celda.Row, COL_CODIGO)
        Else
            ' 2) Busqueda por terminacion en columna A (ej.: escribir solo "0487")
            coincidencias = Application.WorksheetFunction.CountIf( _
                            ws.Columns(COL_CODIGO), "*" & codigo)
            If coincidencias = 1 Then
                Set celda = ws.Columns(COL_CODIGO).Find(What:="*" & codigo, _
                            LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                If Not celda Is Nothing Then Set objetivo = celda
            ElseIf coincidencias > 1 Then
                Beep
                estado = "AMBIGUO: '" & codigo & "' coincide con " & coincidencias & _
                         " cheques. Escribe mas digitos."
            End If
        End If

        If Not objetivo Is Nothing Then
            ' Llevar la vista al cheque encontrado
            ActiveWindow.ScrollRow = Application.Max(1, objetivo.Row - 8)
            objetivo.Select

            If objetivo.Interior.Color = COLOR_MARCA Then
                Beep
                estado = "YA ESTABA MARCADO: " & objetivo.Value & " (fila " & objetivo.Row & ")"
            Else
                objetivo.Interior.Color = COLOR_MARCA
                marcados = marcados + 1
                estado = "OK: " & objetivo.Value & " marcado (fila " & objetivo.Row & ")"
            End If
        ElseIf coincidencias <= 1 Then
            Beep
            estado = "NO ENCONTRADO: " & codigo
        End If

        coincidencias = 0
    Loop

    MsgBox "Sesion terminada." & vbCrLf & "Cheques marcados: " & marcados, _
           vbInformation, "Check List BHD"
End Sub
