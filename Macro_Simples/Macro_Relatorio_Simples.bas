Option Explicit

Sub GerarDocumentoPlanilhaAtiva()
    On Error GoTo TratarErro

    Dim ws As Worksheet
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim templatePath As String
    Dim pastaBase As String
    Dim pastaAno As String
    Dim nomeArquivo As String
    Dim anoEmissao As String
    Dim dataCarregamento As Date
    Dim ultimaLinha As Long, i As Long, j As Long
    Dim linhaSim As Long, linhaIDTF As Long, maisProximo As Long
    Dim idtfsEspecificos As Collection
    Dim idtfEspecificoAchado As Variant
    Dim dataEmissao, compartimentoCarga
    Dim produtoTransportado, dataDescarregamento, localEntrega
    Dim dataLimpeza, horaLimpeza, tipoLimpeza, responsavel

    ' Pasta base onde os arquivos serão salvos
    pastaBase = "C:\User\Pasta\"
    If Right(pastaBase, 1) <> "\" Then pastaBase = pastaBase & "\"

    ' Caminho do template Word
    templatePath = "C:\Users\User\Documents\Modelos\ModeloRelatorio.dotx"

    ' Lista de códigos IDTF específicos
    Set idtfsEspecificos = New Collection
    idtfsEspecificos.Add 40342 ' Exemplo genérico: Cloreto de Potássio
    idtfsEspecificos.Add 40285 ' Exemplo genérico: Fertilizantes artificiais

    ' Inicia o Word
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True

    ' Planilha ativa
    Set ws = ActiveSheet
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

            ' Define ano da emissão
            If IsDate(dataEmissao) Then
                anoEmissao = Year(dataEmissao)
            Else
                anoEmissao = Year(Date)
            End If

            pastaAno = pastaBase & anoEmissao & "\"
            If Dir(pastaAno, vbDirectory) = "" Then MkDir pastaAno

            ' Cria novo documento Word baseado no modelo
            Set WordDoc = WordApp.Documents.Add(templatePath)

            ' Tabela 1
            With WordDoc.Tables(1)
                .Cell(1, 2).Range.Text = Format(dataEmissao, "dd/mm/yyyy")
                .Cell(2, 2).Range.Text = compartimentoCarga
            End With

            If linhaIDTF = 0 Then
                ' Sem produto específico
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
                ' Dados coletados
                dataCarregamento = ws.Cells(linhaIDTF, 2).Value
                produtoTransportado = ws.Cells(linhaIDTF, 8).Value
                dataDescarregamento = ws.Cells(linhaIDTF, 12).Value
                localEntrega = ws.Cells(linhaIDTF, 13).Value
                dataLimpeza = ws.Cells(linhaIDTF + 1, 17).Value
                horaLimpeza = ws.Cells(linhaIDTF + 1, 18).Value
                tipoLimpeza = ws.Cells(linhaIDTF + 1, 14).Value
                responsavel = ws.Cells(linhaIDTF + 1, 19).Value

                ' Tabela 2
                With WordDoc.Tables(2)
                    .Cell(2, 2).Range.Text = Format(dataCarregamento, "dd/mm/yyyy")
                    .Cell(3, 2).Range.Text = produtoTransportado
                    .Cell(4, 2).Range.Text = Format(dataDescarregamento, "dd/mm/yyyy")
                    .Cell(5, 2).Range.Text = localEntrega
                End With

                ' Tabela 3
                With WordDoc.Tables(3)
                    .Cell(3, 2).Range.Text = dataLimpeza
                    .Cell(4, 2).Range.Text = Format(horaLimpeza, "hh:mm")
                    .Cell(5, 2).Range.Text = tipoLimpeza
                    .Cell(6, 2).Range.Text = responsavel
                End With

                ' Nome do arquivo
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
                ' Exemplo genérico de como fica: "Cloreto de Potássio - 18.07.2025.docx"
            End If

            ' Salvar documento
            WordDoc.SaveAs2 pastaAno & nomeArquivo & ".docx"
            WordDoc.Close
        End If
    Next i

    WordApp.Quit
    Set WordApp = Nothing

    MsgBox "Documento(s) gerado(s) com sucesso na pasta do ano: " & pastaAno
    Exit Sub

TratarErro:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro na execução"
    If Not WordApp Is Nothing Then
        WordApp.Quit
        Set WordApp = Nothing
    End If
End Sub
