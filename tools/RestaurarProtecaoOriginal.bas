Attribute VB_Name = "ProtecaoPGFLOW"
Option Explicit

' Execucao manual pelo menu Desenvolvedor > Macros. Nenhum evento ou botao.
Public Sub RestaurarProtecaoOriginal()
    On Error GoTo Falhou
    Dim total As Long
    total = PGFLOWRestaurarProtecao()
    MsgBox CStr(total) & " aba(s) com a protecao original restaurada." & vbCrLf & _
        "Abas originalmente livres continuam livres.", vbInformation, "PGFLOW"
    Exit Sub
Falhou:
    MsgBox Err.Description, vbExclamation, "PGFLOW"
End Sub

' Funcao separada para permitir verificacao automatica sem caixas de dialogo.
Public Function PGFLOWRestaurarProtecao(Optional ByVal ModoVerificacao As Boolean = False) As Long
    Dim partes As Object, doc As Object, no As Object, faixa As Object
    Dim aba As Worksheet, nome As String
    Dim tela As Boolean, eventos As Boolean, total As Long
    tela = Application.ScreenUpdating
    eventos = Application.EnableEvents
    On Error GoTo Falhou
    Set partes = ThisWorkbook.CustomXMLParts.SelectByNamespace("urn:pgflow:protecao:v1")
    If partes.Count <> 1 Then Err.Raise vbObjectError + 2932, , "Mapa original de protecao ausente ou duplicado."
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    If Not doc.LoadXML(partes.Item(1).XML) Then Err.Raise vbObjectError + 2933, , "Mapa original de protecao invalido."
    doc.SetProperty "SelectionNamespaces", "xmlns:p='urn:pgflow:protecao:v1'"
    ' Confere todas as abas antes de restaurar qualquer uma.
    For Each no In doc.SelectNodes("/p:protection/p:sheet")
        Set aba = ThisWorkbook.Worksheets(CStr(no.getAttribute("name")))
    Next no
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For Each no In doc.SelectNodes("/p:protection/p:sheet")
        Set aba = ThisWorkbook.Worksheets(CStr(no.getAttribute("name")))
        nome = aba.Name
        ' Nunca desprotege uma aba nem amplia o conjunto protegido do molde.
        If Not (aba.ProtectContents Or aba.ProtectDrawingObjects Or aba.ProtectScenarios) Then
            aba.Cells.Locked = Flag(no, "defaultLocked", True)
            aba.Cells.FormulaHidden = Flag(no, "defaultHidden", False)
            For Each faixa In no.SelectNodes("p:range")
                aba.Range(CStr(faixa.getAttribute("address"))).Locked = Flag(faixa, "locked", True)
                aba.Range(CStr(faixa.getAttribute("address"))).FormulaHidden = Flag(faixa, "hidden", False)
            Next faixa
            ' Cada mesclagem e aplicada inteira, sem alterar a mesclagem ou seu conteudo.
            For Each faixa In no.SelectNodes("p:merge")
                aba.Range(CStr(faixa.getAttribute("address"))).Locked = Flag(faixa, "locked", True)
                aba.Range(CStr(faixa.getAttribute("address"))).FormulaHidden = Flag(faixa, "hidden", False)
            Next faixa
            aba.Protect Password:="PGFLOW2026", _
                DrawingObjects:=Flag(no, "objects", False), Contents:=Flag(no, "sheet", True), _
                Scenarios:=Flag(no, "scenarios", False), UserInterfaceOnly:=False, _
                AllowFormattingCells:=Not Flag(no, "formatCells", True), _
                AllowFormattingColumns:=Not Flag(no, "formatColumns", True), _
                AllowFormattingRows:=Not Flag(no, "formatRows", True), _
                AllowInsertingColumns:=Not Flag(no, "insertColumns", True), _
                AllowInsertingRows:=Not Flag(no, "insertRows", True), _
                AllowInsertingHyperlinks:=Not Flag(no, "insertHyperlinks", True), _
                AllowDeletingColumns:=Not Flag(no, "deleteColumns", True), _
                AllowDeletingRows:=Not Flag(no, "deleteRows", True), _
                AllowSorting:=Not Flag(no, "sort", True), _
                AllowFiltering:=Not Flag(no, "autoFilter", True), _
                AllowUsingPivotTables:=Not Flag(no, "pivotTables", True)
            If Flag(no, "selectUnlockedCells", False) Then
                aba.EnableSelection = xlNoSelection
            ElseIf Flag(no, "selectLockedCells", False) Then
                aba.EnableSelection = xlUnlockedCells
            Else
                aba.EnableSelection = xlNoRestrictions
            End If
            total = total + 1
        End If
    Next no
    Application.ScreenUpdating = tela
    Application.EnableEvents = eventos
    PGFLOWRestaurarProtecao = total
    Exit Function
Falhou:
    Application.ScreenUpdating = tela
    Application.EnableEvents = eventos
    If ModoVerificacao Then
        Application.StatusBar = "Falha ao restaurar " & nome & ": " & Err.Description
        PGFLOWRestaurarProtecao = -1
        Exit Function
    End If
    Err.Raise vbObjectError + 2934, "RestaurarProtecaoOriginal", _
        "Falha ao restaurar " & nome & ": " & Err.Description
End Function

Private Function Flag(ByVal no As Object, ByVal chave As String, ByVal padrao As Boolean) As Boolean
    Dim valor As Variant
    valor = no.getAttribute(chave)
    If IsNull(valor) Or IsEmpty(valor) Then
        Flag = padrao
    Else
        Flag = (CStr(valor) = "1" Or LCase$(CStr(valor)) = "true")
    End If
End Function
