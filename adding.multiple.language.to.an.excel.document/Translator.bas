Attribute VB_Name = "Translator"
Sub TranslateActiveCell()
Attribute TranslateActiveCell.VB_ProcData.VB_Invoke_Func = "t\n14"
    Call Translate(ActiveCell)
End Sub

Function Translate(cell As range)
Attribute Translate.VB_ProcData.VB_Invoke_Func = "t\n14"
    Dim ws As Worksheet
    
    ' Creating translation sheet if not already done.
    Set ws = InitTranslator()
    
    If Not cell.HasFormula And Not cell.Value = Empty And Not IsNumeric(cell.Value) Then
        Dim existingTranslation As range
        Dim nextAvailableCell As range
        
        Set existingTranslation = ws.range("A:A").Find(What:=cell.Value, LookAt:=xlWhole, MatchCase:=True)
        
        If Not existingTranslation Is Nothing Then
            cell.Formula = GetTranslationFormula(existingTranslation)
        Else
            Set nextAvailableCell = ws.range("A:A").Find(What:=Empty, LookAt:=xlWhole, MatchCase:=True)
            
            ' Set original value in translation tab
            nextAvailableCell.Value = cell.Value
            ' Set the formula to reference the value in the translation tab
            cell.Formula = GetTranslationFormula(nextAvailableCell)
        End If
    End If
    
    EndTranslator
End Function

Function GetSheetName() As String
    GetSheetName = "translation"
End Function

Function GetTranslationFormula(rng As range) As String
    GetTranslationFormula = "=IF(lang=""english"", " & GetCellFullAddress(rng) & ", " & GetCellFullAddress(rng.Offset(0, 1)) & ")"
End Function

Function InitTranslator() As Worksheet
    Dim ws As Worksheet
    
    For Each Sheet In Worksheets
        If GetSheetName() = Sheet.Name Then
            Set ws = Sheet
        End If
    Next Sheet
    
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
        Set InitTranslator = ws
        Exit Function
    End If
    
    Set ws = Worksheets.Add()
    ws.Name = GetSheetName()
    ws.range("A1").Value = "english"
    ws.range("B1").Value = "french"
    'ws.range("C1").Value = "spanish"
    
    Set InitTranslator = ws
End Function

Sub EndTranslator()
    Worksheets(GetSheetName()).Visible = xlSheetHidden
End Sub

Function GetCellFullAddress(cell As range) As String
    Dim worksheetName As String
    
    worksheetName = cell.Parent.Name
    
    ' Adding quote to ensure valid name when space are used.
    GetCellFullAddress = "'" & worksheetName & "'!" & cell.Address(External:=False)
End Function
