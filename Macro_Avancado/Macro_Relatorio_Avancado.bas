Option Explicit

Sub GerarRelatorioCompleto()
    On Error GoTo TratarErro

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim templatePath As String
    Dim pastaBase As String
    Dim pastaAno As String
    Dim anoEmissao As String
    Dim nomeArquivo As String
    Dim dataCarregamento As Date
    Dim ultimaLinha As Long, i As Long, j As Long
    Dim linhaSim As Long, linhaIDTF As Long, maisProximo As Long
    Dim idtfsEspecificos As Collection
    Dim idtfEspecificoAchado As Variant
    Dim dataEmissao, compartimentoCarga
    Dim produtoTransportado, dataDescarregamento, localEntrega
    Dim dataLimpeza, horaLimpeza, tipoLimpeza, responsavel
    Dim localArquivos As String ' Armazena caminhos dos arquivos salvos

    ' Caminhos genéricos
    pastaBase = "C:\User\Pasta\" ' ← Pasta de saída dos relatórios
    If Right(pastaBase, 1) <> "\" Then pastaBase = pastaBase & "\"
    templatePath = "C:\Users\User\Documents\Modelos\ModeloGenerico.dotx" ' ← Caminho do template Word

    ' Inicia aplicação do Word
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True

    ' Códigos genéricos de produto que indicam relevância
    Set idtfsEspecificos = New Collection
    idtfsEspecificos.Add 40342 ' Exemplo: Cloreto de Potássio
    idtfsEspecificos.Add 40285 ' Exemplo: Fertilizantes artificiais

    Set wb = ThisWorkbook

    ' Percorre todas as abas da planilha
    For Each ws In wb.Sheets
        ultimaLinha = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row ' Coluna G

        For i = 2 To ultimaLinha
            If LCase(ws.Cells(i, 7).Value) = "sim" Then
                linhaSim = i

                ' Busca o IDTF mais próximo acima da linha marcada
                linhaIDTF = 0
                maisProximo = 0
                For j = linhaSim - 1 To 1 Step -1
                    For Each idtfEspecificoAchado In idtfsEspecificos
                        If ws.Cells(j, 6).Value = idtfEspecificoAchado Then
                            If maisProximo = 0 Or j > maisProximo Then
                                maisProximo = j
                                linhaIDTF = j
                            End If
                        End If
                    Next
                Next j

                dataEmissao = ws.Cells(linhaSim, 2).Value
                compartimentoCarga = ws.Name

                ' Determina o ano
                If IsDate(dataEmissao) Then
                    anoEmissao = Year(dataEmissao)
                Else
                    anoEmissao = Year(Date)
                End If

                pastaAno = pastaBase & anoEmissao & "\"
                If Dir(pastaAno, vbDirectory) = "" Then MkDir pastaAno

                ' Cria documento Word a partir do modelo
                Set WordDoc = WordApp.Documents.Add(templatePath)

                ' Preenche Tabela 1 (dados da linha marcada como "sim")
                With WordDoc.Tables(1)
                    .Cell(1, 2).Range.Text = Format(dataEmissao, "dd/mm/yyyy")
                    .Cell(2, 2).Range.Text = compartimentoCarga
                End With

                If linhaIDTF = 0 Then
                    ' Sem produto específico encontrado
                    With WordDoc.Tables(2)
                        For j = 2 To 5
                            .Cell(j, 2).Range.Text = "Verificar o mês anterior"
                        Next j
                    End With
                    With WordDoc.Tables(3)
                        For j = 3 To 6
                            .Cell(j, 2).Range.Text = "Verificar o mês anterior"
                        Next j
                    End With
                    nomeArquivo = "Sem IDTF - " & Format(dataEmissao, "dd.mm.yyyy")
                Else
                    ' Coleta dados da linha do produto
                    dataCarregamento = ws.Cells(linhaIDTF, 2).Value
                    produtoTransportado = ws.Cells(linhaIDTF, 8).Value
                    dataDescarregamento = ws.Cells(linhaIDTF, 12).Value
                    localEntrega = ws.Cells(linhaIDTF, 13).Value
                    dataLimpeza = ws.Cells(linhaIDTF + 1, 17).Value
                    horaLimpeza = ws.Cells(linhaIDTF + 1, 18).Value
                    tipoLimpeza = ws.Cells(linhaIDTF + 1, 14).Value
                    responsavel = ws.Cells(linhaIDTF + 1, 19).Value

                    ' Preenche Tabela 2
                    With WordDoc.Tables(2)
                        .Cell(2, 2).Range.Text = Format(dataCarregamento, "dd/mm/yyyy")
                        .Cell(3, 2).Range.Text = produtoTransportado
                        .Cell(4, 2).Range.Text = Format(dataDescarregamento, "dd/mm/yyyy")
                        .Cell(5, 2).Range.Text = localEntrega
                    End With

                    ' Preenche Tabela 3
                    With WordDoc.Tables(3)
                        .Cell(3, 2).Range.Text = dataLimpeza
                        .Cell(4, 2).Range.Text = Format(horaLimpeza, "hh:mm")
                        .Cell(5, 2).Range.Text = tipoLimpeza
                        .Cell(6, 2).Range.Text = responsavel
                    End With

                    ' Gera nome do arquivo com produto + data
                    nomeArquivo = Replace(produtoTransportado, "/", "_")
                    nomeArquivo = Replace(nomeArquivo, "\", "_")
                    nomeArquivo = Replace(nomeArquivo, ":", "_")
                    nomeArquivo = Replace(nomeArquivo, "*", "_")
                    nomeArquivo = Replace(nomeArquivo, "?", "_")
                    nomeArquivo = Replace(nomeArquivo, """", "_")
                    nomeArquivo = Replace(nomeArquivo, "<", "_")
                    nomeArquivo = Replace(nomeArquivo, ">", "_")
                    nomeArquivo = Replace(nomeArquivo, "|", "_")
                    If nomeArquivo = "" Then nomeArquivo = "Documento"

                    nomeArquivo = nomeArquivo & " - " & Format(dataCarregamento, "dd.mm.yyyy")
                    ' Exemplo: "Cloreto de Potássio - 18.07.2025.docx"
                End If

                ' Salva e fecha documento
                WordDoc.SaveAs2 pastaAno & nomeArquivo & ".docx"
                WordDoc.Close

                ' Armazena caminho
                localArquivos = localArquivos & pastaAno & nomeArquivo & ".docx" & vbCrLf

                Exit For ' Só processa a primeira ocorrência "sim" por aba
            End If
        Next i
    Next ws

    WordApp.Quit
    Set WordApp = Nothing

    MsgBox "Relatórios gerados com sucesso nos seguintes locais:" & vbCrLf & vbCrLf & localArquivos, vbInformation, "Concluído"
    Exit Sub

TratarErro:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro na execução"
    If Not WordApp Is Nothing Then
        WordApp.Quit
        Set WordApp = Nothing
    End If
End Sub
