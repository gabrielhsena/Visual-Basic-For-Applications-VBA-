Mod_Framework
'---------------------------------------------------------------------------------------
' Module    : Sub_Framework
' Author    : GABRIEL HERNANDES SENA
' Purpose   : PROCESSOS GERAIS
'---------------------------------------------------------------------------------------

Dim Filter As String, Title As String, msg As String, NomeArquivo As String, TXTFile As String, OpenFileType As String
Dim i As Integer, FilterIndex As Integer, TmpInteger As Integer
Dim Filename As Variant
Dim book1 As Workbook, book2 As Workbook
Dim FlagErro As Boolean
Global var_DataAnalisada As String, TotalDias As Integer, xCell

'Macro para abrir vários arquivos
Sub OpenFiles()
    Dim Filter As String, Title As String, msg As String
    Dim i As Integer, FilterIndex As Integer
    Dim Filename As Variant
    'Caminho da Planilha Ativa
    WBPath = ActiveWorkbook.Path

    'Seta a variável book1 como a planilha ativa
    Set book1 = ActiveWorkbook

    ' Filtros
    Filter = "Excel Files (*.xls),*.xls," & _
             "Excel Files Binary (*.xlsb),*.xlsb," & _
             "Text Files (*.txt),*.txt," & _
             "All Files (*.*),*.*"
    ' Filtro para *.*
    FilterIndex = 4
    ' "Setar" o caption do OpenFileDialog
    Title = "Abra o Arquivo " & NomeArquivo & "..."

    With Application

        ' Reset Drive/Path
        If Left(.DefaultFilePath, 1) <> "\" Then
            ' Seleciona o drive inicial
            ChDrive (Left(CStr(Environ("USERPROFILE")), 1))
            ChDir (CStr(Environ("USERPROFILE")))
        Else
            ChDrive (Left(WBPath, 1))
            ChDir (WBPath)
        End If
        ' Nomes dos arquivos = arquivos selecionados
        Filename = .GetOpenFilename(Filter, FilterIndex, Title, , True)

    End With
    ' Sair quando cancela
    If Not IsArray(Filename) Then
        MsgBox "Nenhum arquivo foi escolhido."
        Exit Sub
    End If
    ' Abre Arquivo
    ' Abre Arquivo
    For i = LBound(Filename) To UBound(Filename)
        If ((OpenFileType = "TXT")) Then
            TXTFile = Filename(i)
            LerArquivo
        Else
            'Abre o arquivo da vez
            Workbooks.Open Filename(i)
            'Rotina para copiar as worksheets
            CopiarWorksheets
            'Rotina para fechar as outras workbooks abertas
            FechaWB
        End If
    Next i

End Sub


'Macro para Ler Arquivo TXT
Sub LerArquivo()

    Open TXTFile For Input As #1
    'adciona uma planilha nova no final
    Sheets.Add after:=Sheets(Sheets.Count)
    'Nome da planila é Arquivo_Texto_Número
    Worksheets(Sheets.Count).Name = "Arquivo_Texto"
    'seleciona a nova planilha
    Worksheets(Sheets.Count).Select
    'Seleciona a 1° célula e começa a imprimir o arquivo txt
    Range("A1").Select

    Do While Not EOF(1)
        Line Input #1, Lin
        If Lin = "" Then
            'Do nothing
        ElseIf InStr(1, Lin, "===") > 0 Then
            Lin = Replace(Lin, "=", "*")
            ActiveCell.Value = Lin
            ActiveCell.Offset(1, 0).Select
        Else
            ActiveCell.Value = Lin
            ActiveCell.Offset(1, 0).Select
        End If
    Loop
    Close #1

End Sub

Sub UpperCase()
'Converte em maiusculo e tira os espaços desnecessários
    For Each xCell In Selection
        xCell.Value = Trim(UCase(xCell.Value))
    Next xCell
End Sub

Sub TrimCells()
'Converte em maiusculo e tira os espaços desnecessários
    For Each xCell In Selection
        xCell.Value = Trim(xCell.Value)
    Next xCell
End Sub

Sub DateConv()
    For Each xCell In Selection
        If IsNumeric(xCell.Value) Then GoTo NextXcell:
        xCell.Value = CDbl(DateValue(xCell.Value))
NextXcell:
        xCell.NumberFormat = "dd/mm/yyyy"
    Next xCell
End Sub

Sub TimeConv()
    For Each xCell In Selection
        If IsNumeric(xCell.Value) Then GoTo NextXcell:
        xCell.Value = CDbl(TimeValue(xCell.Value))
NextXcell:
        xCell.NumberFormat = "hh:mm:ss"
    Next xCell
End Sub

Sub Convert_text_to_number()
    BetterPerformace True
    
    For Each xCell In Selection
        If IsNumeric(xCell) Then xCell.Value = 1 * xCell.Value
    Next xCell
    
    BetterPerformace False
End Sub

Sub DeletaWS()

    WSCounter = Sheets.Count

    For i = WSCounter To 1 Step -1
        WSName = Worksheets(i).Name
        If ((WSName <> "Settings")) Then
            Worksheets(i).Delete
        End If
    Next i
End Sub

Sub UnhideRowsAndColumns()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ws = ActiveWorkbook.Sheets.Count

    For i = 5 To ws
        Worksheets(i).Select
        Worksheets(i).Unprotect Password:="2014"
        ActiveWindow.DisplayHeadings = True
        ActiveSheet.Cells.EntireRow.Hidden = False
        ActiveSheet.Cells.EntireColumn.Hidden = False

    Next i

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.CutCopyMode = False

End Sub

Sub CopiarWorksheets()
    Application.ScreenUpdating = False

    If SheetExists(book1.Name, "Gráficos Pausas") = True Then book1.Sheets("Gráficos Pausas").Delete

    Set book2 = ActiveWorkbook
    i = 0
    For Each ws In Worksheets
        'Pega o nome da planilha na Workbook que está sendo aberta
        strPlan = ws.Name
        'Copia a planilha para a planilha que irá gerar o CSV
        book2.Worksheets(strPlan).Copy after:=book1.Worksheets(book1.Sheets.Count)
        'Renomeia a planilha importada
        book1.Worksheets(book1.Sheets.Count).Name = "TEMP" & i
        i = i + 1
    Next ws
End Sub

Sub FechaWB()

'Conta a qtd de Workbook
    WBcounter = Workbooks.Count
    'fecha as workbooks
    Do
        Workbooks(WBcounter).Close
        'Diminui contador
        WBcontador = WBcounter - 1
        WBcounter = WBcontador
    Loop Until WBcounter = 1

End Sub

Sub ImportBOOK()

    OpenFileType = "BOOK"
    Application.StatusBar = "Abrindo Arquivo..."
    'Abre o OpenFileDialog para abrir o arquivo e o processa
    OpenFiles

    'Se erro sai do processo
    If FlagErro = True Then
        'Deleta planilhas não desejadas
        DeletaWS
        Exit Sub
    End If

    ClearStatusBar

End Sub

Sub ClearStatusBar()
'Limpando a status bar
    Application.StatusBar = False
End Sub

Sub ConvertXLS2XLSX()

'Variável do arquivo original
    Dim sFileName As String
    'Variável do novo arquivo
    Dim sNewFileName As String
    'Variável da planilha
    Dim wkbk As Workbook

    'Variável da Constante .xlsx
    Const sEXT As String = ".xlsx"

    'GetOpenFilename com o filtro em xls
    sFileName = Application.GetOpenFilename("Excel Files (*.xls), *.xls", , "Selecione o arquivo do GSP")

    'Se cancelou a abertura do arquivo, então sai do procedimento
    If sFileName = "False" Then
        MsgBox "Cancel was clicked. Operation aborted.", vbOKOnly + vbExclamation
        Exit Sub
    End If

    'Desativa a atualiszação da tela
    Application.ScreenUpdating = False
    'Desabilita os cálculos automáticos
    Application.Calculation = xlCalculationManual

    'Cria o nome do novo arquivo
    sNewFileName = Left(sFileName, Len(sFileName) - 4) & sEXT

    'Abre o arquivo originaç
    Set wkbk = Workbooks.Open(Filename:=sFileName)
    'Salva como xls
    wkbk.SaveAs Filename:=sNewFileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    'fecha o arquivo
    wkbk.Close True
    'Termina a variável
    Set wkbk = Nothing
    'abre o novo arquivo
    Set wkbk = Workbooks.Open(Filename:=sNewFileName)
    'e o fecha novamente
    wkbk.Close True
    'deleta o arquivo original
    Kill sFileName

    'Reseta os calculos
    Application.Calculation = xlCalculationAutomatic
    'Reseta a atualização de tela
    Application.ScreenUpdating = True

End Sub

Sub Total_de_Dias()

'Recebe o ano da variável
    varAno = Year(var_DataAnalisada)
    'Recebe o dia da variável
    varDia = Month(var_DataAnalisada)
    'Recebe o mes da variável
    varMes = Day(var_DataAnalisada)

    'Data serial da data selecionada
    DataSelecionada = DateSerial(varAno, varDia, varMes)

    'Total de dias do mês
    TotalDias = 32 - Day(DataSelecionada - Day(DataSelecionada) + 32)

End Sub

Sub BetterProcess_i()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

End Sub

Sub BetterProcess_f()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    Application.StatusBar = False

End Sub

Sub BetterPerformace(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    Application.DisplayAlerts = Not (isOn)
    ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub

Sub SubProcesso()
    BetterPerformace True

    'código

    BetterPerformace False

End Sub

Sub Remove_Links()

    Dim ExternalLinks As Variant
    Dim wb As Workbook
    Dim x As Long
    Set wb = ActiveWorkbook

    Dim xConnect As Object
    For Each xConnect In ActiveWorkbook.Connections
        If xConnect.Name <> "ThisWorkbookDataModel" Then xConnect.Delete
    Next xConnect

    ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)

    If IsArrayEmpty(ExternalLinks) = False Then
        For x = 1 To UBound(ExternalLinks)
            wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
        Next x

    End If

End Sub
                        
Sub ReplaceAccentedChar()
Const sFm As String = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
Const sTo As String = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
Dim i As Long, ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    For i = 1 To Len(sFm)
        ws.Cells.Replace Mid(sFm, i, 1), Mid(sTo, i, 1), LookAt:=xlPart, MatchCase:=True
    Next i
Next ws
End Sub       


Sub HTML()


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Dim RC, CC, iRow, iCol


RC = ActiveSheet.UsedRange.Rows.Count
CC = ActiveSheet.UsedRange.Columns.Count

Col_Header = 3

For iRow = 4 To RC


Cells(iRow, CC + 1).Value = "<table>"

For iCol = 2 To CC

Cells(iRow, CC + 1).Value = Cells(iRow, CC + 1).Value & "<tr><td>" & Cells(Col_Header, iCol).Value & "</td>"
Cells(iRow, CC + 1).Value = Cells(iRow, CC + 1).Value & "<td>" & Cells(iRow, iCol).Value & "</td></tr>"

Next iCol

Cells(iRow, CC + 1).Value = Cells(iRow, CC + 1).Value & "</table>"

Next iRow

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

