Mod_Functions
'---------------------------------------------------------------------------------------
' Module    : Sub_Functions
' Author    : GABRIEL HERNANDES SENA
' Purpose   : FUNÇÕES GERAIS
'---------------------------------------------------------------------------------------

'Função MsgBox com Timeout
'Função MsgBox com Timeout
#If VB7 Then
Private Declare PtrSafe Function MsgBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As VbMsgBoxStyle, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
#Else
Private Declare Function MsgBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As VbMsgBoxStyle, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
#End If


Sub Function_Time()

    Dim TotalTime As Date
    STime = Now()
    

    ETime = Now()
    TotalTime = ETime - STime
    Resumo = "Processo Concluído. Tempo Total: " & TotalTime & vbCrLf & "Arquivo criado com sucesso: " & vbCrLf & TargetFile
    MsgBoxTimeout 0, Resumo, "Conluído", vbInformation, 0, 5000

    Application.StatusBar = False

End Sub





'Função para determinar a existência de um arquivo.
Public Function FileExists(ByVal Filename As String) As Boolean
'Verifica se o caminho não é vazio ou se não há erros
    If (Dir$(Filename) <> "") Or (Err.Number <> 0) Then
        'Caso esteja ok o arquivo existe
        FileExists = True
    Else
        'Caso contrário o arquivo não existe
        FileExists = False
    End If

End Function
'Função para determinar a existência de um caminho.
Public Function FolderExists(ByVal pathname As String) As Boolean
'Verifica se o caminho existe ou se não há erros.
    If (Dir$(pathname & "\*.*", vbNormal + vbDirectory) <> "") Or (Err.Number <> 0) Then
        'Caso esteja ok o caminho existe
        FolderExists = True
    Else
        'caso contrário o caminho não existe
        FolderExists = False
    End If

End Function
'Função para determinar a existência de uma planilha.
Function SheetExists(WorkbookName, SheetName As String) As Boolean
'variável de nome recebe vazio
    SN = ""
    'variável que conta a quantidade de planilhas
    SC = Workbooks(WorkbookName).Sheets.Count
    'variável de contagem i
    i = 1
    'laço enquanto para verificação da existencia da planilha
    'enquanto SN for diferente da planilha procurada e i for menor que SC
    While (SN <> SheetName) And (i <= SC)
        'Se a planilha ativa tiver o nome procurado
        If Workbooks(WorkbookName).Sheets(i).Name = SheetName Then
            'armazena o nome da planilha na variável SN
            SN = Workbooks(WorkbookName).Sheets(i).Name
        End If
        'soma mais um na variável i
        i = i + 1
    Wend
    'Se SN não for vazio então a planilhaa existe
    If SN <> "" Then
        SheetExists = True
    Else
        SheetExists = False
    End If

End Function
'Função para determinar a extensão de um arquivo.
Public Function GetFileExt(ByVal Filename As String) As String

    pos = 0

    i = Int(Len(Filename))
    While (i <> 0) And (pos = 0)
        C = Mid(Filename, i, 1)
        If C = "." Then
            pos = i + 1
        End If
        i = i - 1
    Wend

    If pos <> 0 Then
        GetFileExt = Mid(Filename, pos, (Len(Filename) + 1 - pos))
    Else
        GetFileExt = ""
    End If

End Function
'Função para determinar o nome de um arquivo.
Public Function GetFileName(ByVal Filename As String) As String

    pos = 0

    i = Int(Len(Filename))
    While (i <> 0) And (pos = 0)
        C = Mid(Filename, i, 1)
        If C = "." Then
            pos = i - 1
        End If
        i = i - 1
    Wend

    If pos <> 0 Then
        GetFileName = Left(Filename, pos)
    Else
        GetFileName = Filename
    End If

End Function
'Função para determinar o número da semana.
Public Function WeekNumber(AnyDate As Date) As Integer

    Dim FirstDay, ThisDate, ThisYear, DayOfWeek As Integer

    AnyDate = Int(AnyDate)
    ThisYear = Year(AnyDate)
    ThisDate = DateSerial(ThisYear, Month(AnyDate), Day(AnyDate))
    FirstDay = DateSerial(ThisYear, 1, 1)
    DayOfWeek = Weekday(FirstDay, vbMonday)
    WeekNumber = ((ThisDate - FirstDay + DayOfWeek - 4) / 7) + 1

End Function
'Função para determinar o número de uma linha.
Public Function Find_Row(ByVal WhatFind As String, ByVal WhereColumn As Integer) As Long

'Seleciona a coluna aonde irá procurar a linha
    Columns(WhereColumn).Select
    'Tenta procurar a linha
    Set LinFound = Selection.Find(What:=WhatFind, after:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    'Se não achar retorna 0
    If LinFound Is Nothing Then
        Find_Row = 0
        'Se achar retorna o número da linha
    Else
        Find_Row = LinFound.Row
    End If

End Function
'Função para determinar o número de uma coluna.
Public Function Find_Col(ByVal WhatFind As String, ByVal WhereRow As Integer) As Long

'Seleciona a linha aonde irá procurar a coluna
    Rows(WhereRow).Select
    'Tenta procurar a coluna
    Set ColFound = Selection.Find(What:=WhatFind, after:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    'Se não achar retorna 0
    If ColFound Is Nothing Then
        Find_Col = 0
        'Se achar retorna o número da coluna
    Else
        Find_Col = ColFound.Column
    End If

End Function
'Função para determinar o número de uma linha.
Public Function Count_Rows(ByVal WorksheetName As String) As Long
'Pega o Nome da Worksheet atual
    Set AWS = ActiveSheet
    'Seleciona a workhsheet que se quer o nome
    Worksheets(WorksheetName).Select
    'Conta a quantidade de linhas
    Count_Rows = ActiveSheet.UsedRange.Rows.Count
    'Seleciona a planilha atual
    AWS.Select
End Function

Public Function Count_Columns(ByVal WorksheetName As String) As Long
'Pega o Nome da Worksheet atual
    Set AWS = ActiveSheet
    'Seleciona a workhsheet que se quer o nome
    Worksheets(WorksheetName).Select
    'Conta a quantidade de linhas
    Count_Columns = ActiveSheet.UsedRange.Columns.Count
    'Seleciona a planilha atual
    AWS.Select
End Function


'Função que conta a quantidade de caracteres da linha do arquivo texto
Public Function CharCount(strString As String) As String
'Delaração de variáveis
    Dim intLenOfString As Integer
    Dim intCounter As Integer
    Dim intNumOfSemiColoms As Integer

    'Se a linha não possui nenhum caractere, então a função retorna 0 e fecha
    If Len(strString) = 0 Then
        CharCount = 0
        Exit Function
    End If

    'conta o tamanho da string
    intLenOfString = Len(strString)

    'Conta quanto semicolom possui na linha, se tiver, há um incremento na variável que armazena o número de ;
    For intCounter = 1 To intLenOfString
        Select Case Mid$(strString, intCounter, 1)
            'Caso ;, incremento de +1
        Case ";"
            intNumOfSemiColoms = intNumOfSemiColoms + 1
        End Select
    Next
    'armazena o número de ;
    CharCount = intNumOfSemiColoms
End Function

Public Function RegexTester(TestStr As String, RegExStr As String) As Boolean
    Dim txt As String
    txt = TestStr
    'Cria objeto para fazer a validação do REGEX
    With CreateObject("VBScript.RegExp")
        .Pattern = RegExStr
        'Testa a string no pattern
        If .test(txt) Then
            'Se verdadeiro retorna verdadeiro
            RegexTester = True
        Else
            RegexTester = False
        End If
    End With

End Function

'---------------------------------------------------------------------------------------
' Procedure : SetTimeFormat
' Author    : Gabriel Hernandes Sena
' Purpose   : Função para transformar segundos no formato hora
'---------------------------------------------------------------------------------------
'
Public Function SetTimeFormat(ByVal TimeValue As Double)

'Recebe o valor em segundos, divide por 3600 e pega a parte inteira
    hours = Fix(TimeValue / 3600)
    'Recebe o valor em segundos, divide por 60, pega a parte inteira e subtrai as horas
    mins = Fix(TimeValue / 60) - (hours * 60)
    'Recebe o valor em segundos, subtrai os minutos e as horas
    secs = Fix(TimeValue) - (mins * 60) - (hours * 3600)

    'Se o valor de segundos for menor do que 10, então coloca 0 na frente
    If secs < 10 Then secs = "0" & secs
    'Se o valor de minutos for menor do que 10, então coloca 0 na frente
    If mins < 10 Then mins = "0" & mins
    'Se o valor de horas for menor do que 10, então coloca 0 na frente
    If hours < 10 Then hours = "0" & hours
    'Constroi a string de hora
    SetTimeFormat = hours & ":" & mins & ":" & secs
End Function

Public Function QuantidadeDiasMes(var_DataAnalisada As Double)

'Recebe o ano da variável
    varAno = Year(var_DataAnalisada)
    'Recebe o dia da variável
    varDia = Month(var_DataAnalisada)
    'Recebe o mes da variável
    varMes = Day(var_DataAnalisada)

    'Data serial da data selecionada
    DataSelecionada = DateSerial(varAno, varDia, varMes)

    'Total de dias do mês
    QuantidadeDiasMes = 32 - Day(DataSelecionada - Day(DataSelecionada) + 32)

End Function

Public Function UniqueValues(ByVal OrigArray As Variant) As Variant

    Dim vAns() As Variant
    Dim lStartPoint As Long
    Dim lEndPoint As Long
    Dim lCtr As Long, lCount As Long
    Dim iCtr As Integer
    Dim col As New Collection
    Dim sIndex As String

    Dim vTest As Variant, vItem As Variant
    Dim iBadVarTypes(4) As Integer

    'Função não funciona se os elementos do array for dos seguintes tipo
    iBadVarTypes(0) = vbObject
    iBadVarTypes(1) = vbError
    iBadVarTypes(2) = vbDataObject
    iBadVarTypes(3) = vbUserDefinedType
    iBadVarTypes(4) = vbArray


    'Checa se o parametro é um array
    If Not IsArray(OrigArray) Then
        Err.Raise ERR_BP_NUMBER, , ERR_BAD_PARAMETER
        Exit Function
    End If

    'Pega os limites do array
    lStartPoint = LBound(OrigArray)
    lEndPoint = UBound(OrigArray)

    For lCtr = lStartPoint To lEndPoint
        vItem = OrigArray(lCtr)

        'Checa se o tipo de variável é aceitavel
        For iCtr = 0 To UBound(iBadVarTypes)

            If VarType(vItem) = iBadVarTypes(iCtr) Or _
               VarType(vItem) = iBadVarTypes(iCtr) + vbVariant Then

                Err.Raise ERR_BT_NUMBER, , ERR_BAD_TYPE
                Exit Function

            End If

        Next iCtr

        'Adiciona o elemento para a coleção, usando um indice
        'Se acontecer um erro, o elemento já existe

        sIndex = CStr(vItem)

        'primeiro elemento é adicionado automaticamente
        If lCtr = lStartPoint Then
            col.Add vItem, sIndex
            ReDim vAns(lStartPoint To lStartPoint) As Variant
            vAns(lStartPoint) = vItem
        Else
            On Error Resume Next
            col.Add vItem, sIndex

            If Err.Number = 0 Then
                lCount = UBound(vAns) + 1
                ReDim Preserve vAns(lStartPoint To lCount)
                vAns(lCount) = vItem
            End If
        End If
        Err.Clear
    Next lCtr

    UniqueValues = vAns

End Function

Public Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim LB As Long
    Dim UB As Long

    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I
        ' cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occasions, LBound is 0 and
        ' UBound is -1.
        ' To accommodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(Arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function
                                                
'Função que valida CPF
Public Function lfValidaCPF(ByVal lNumCPF As String) As Boolean
    Application.Volatile
    
    Dim lMultiplicador  As Integer
    Dim lDv1            As Integer
    Dim lDv2            As Integer
    
    lMultiplicador = 2
    
    'Realiza o preenchimento dos zeros á esquerda
    lNumCPF = String(11 - Len(lNumCPF), "0") & lNumCPF
    
    'Realiza o cálculo do dividendo para o dv1 e o dv2
    For i = 9 To 1 Step -1
        lDv1 = (Mid(lNumCPF, i, 1) * lMultiplicador) + lDv1
        
        lDv2 = (Mid(lNumCPF, i, 1) * (lMultiplicador + 1)) + lDv2
        
        lMultiplicador = lMultiplicador + 1
    Next
    
    'Realiza o cálculo para chegar no primeiro dígio
    lDv1 = lDv1 Mod 11
    
    If lDv1 >= 2 Then
        lDv1 = 11 - lDv1
    Else
        lDv1 = 0
    End If
    
    'Realiza o cálculo para chegar no segundo dígido
    lDv2 = lDv2 + (lDv1 * 2)
    
    lDv2 = lDv2 Mod 11
    
    If lDv2 >= 2 Then
        lDv2 = 11 - lDv2
    Else
        lDv2 = 0
    End If
    
    'Realiza a validação e retorna na função
    If Right(lNumCPF, 2) = CStr(lDv1) & CStr(lDv2) Then
        lfValidaCPF = True
    Else
        lfValidaCPF = False
    End If
End Function
                                                
