# 📝 Macro VBA para Geração de Documentos Word a partir da Planilha Ativa

## 📌 Descrição
Automação desenvolvida em VBA no Excel que varre a planilha ativa, identifica entradas marcadas com "sim" e gera documentos Word formatados com base em um template pré-definido. A lógica inclui coleta de dados operacionais, verificação por código de produto e organização de arquivos por período.

## ⚙️ Tecnologias Utilizadas
- **Excel VBA**
- **Microsoft Word (via automação COM)**
- **Manipulação de células e tabelas**
- **Criação de diretórios e nomenclatura dinâmica**

## 🚀 Funcionalidades
- Verifica entradas com valor "sim" na coluna de referência.
- Identifica produtos específicos por códigos (IDTFs) previamente parametrizados.
- Coleta dados como data de emissão, produto, responsável e local de entrega.
- Preenche automaticamente três tabelas em um modelo Word.
- Cria pastas por ano e salva os documentos com nomes dinâmicos:
  - Exemplo: `Cloreto de Potássio - 18.07.2025.docx`

## 📂 Como Usar
1. Estruture a planilha com os dados operacionais conforme esperado.
2. Ajuste os caminhos do `templatePath` e `pastaBase` conforme seu ambiente local.
3. Execute a macro `GerarDocumentoPlanilhaAtiva`.
4. Verifique a pasta de destino para os documentos gerados.

## 🧠 Abordagem Técnica
- Utiliza lógica de busca retroativa para encontrar o produto mais próximo da entrada marcada.
- Os códigos IDTF são parametrizados via `Collection` e podem ser modificados conforme o tipo de produto.
- Evita hardcode, permitindo fácil manutenção e escalabilidade da automação.
- Realiza tratamento completo de erros para garantir execução segura.

## ✨ Observações
> Este projeto foi adaptado para fins didáticos e demonstrativos. Os caminhos e dados utilizados são genéricos e não representam sistemas reais ou controlados por empresas.

## ✍️ Autor
**Kauan da Silva Terrão.**