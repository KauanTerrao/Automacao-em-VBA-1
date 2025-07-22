# 📄 Macro VBA — Relatório Automatizado Multiabás com Template Word

## 📌 Descrição
Automação em VBA que percorre todas as abas de uma planilha, verifica dados marcados como "sim", coleta informações logísticas e gera documentos Word automaticamente com base em um modelo estruturado. Os arquivos são salvos com nomes dinâmicos e organizados em pastas por ano.

## ⚙️ Tecnologias Utilizadas
- **Microsoft Excel VBA**
- **Microsoft Word (automação via COM)**
- **Estrutura de Collections para codificação parametrizada**
- **Lógica de busca reversa**
- **Criação de diretórios por ano**
- **Tratamento robusto de erros**

## 🚀 Funcionalidades
- Identifica células com valor `"sim"` em todas as abas da planilha.
- Busca o produto relevante (via código IDTF) acima da linha marcada.
- Coleta dados como produto, data de emissão, responsáveis e local de entrega.
- Preenche três tabelas em um documento Word baseado em template `.dotx`.
- Salva arquivos em subpastas por ano, com nomes formatados como:
- `Cloreto de Potássio - 18.07.2025.docx`
- Exibe mensagem com os caminhos gerados ao final da execução.

## 🧩 Parametrização de Produtos
A automação utiliza uma lista de códigos IDTF (Identificadores de Tipos de Fertilizantes), que representam produtos específicos de interesse. Esses códigos são definidos no início do código e podem ser adaptados conforme o cenário.

**VBA:**
Set idtfsEspecificos = New Collection
idtfsEspecificos.Add 40342 ' Cloreto de Potássio
idtfsEspecificos.Add 40285 ' Corretivos minerais do solo

## 📂 Como Usar
- Verifique se sua planilha está estruturada conforme esperado nas colunas padrão.
- Ajuste os caminhos do template Word (templatePath) e pasta de destino (pastaBase).
- Execute a macro GerarRelatorioCompleto no Excel.
- Os relatórios serão salvos por ano e listados ao final em uma mensagem de confirmação.

##🔧 Observações Técnicas
- Apenas a primeira ocorrência de "sim" por aba é processada.
- Se nenhum IDTF for encontrado, a macro insere "Verificar o mês anterior" nas tabelas.
- O modelo Word deve conter pelo menos 3 tabelas com estrutura compatível.

##❗ Aviso de Ética
Esta automação foi adaptada para fins didáticos e demonstrativos. Os dados utilizados são genéricos e nenhuma estrutura corporativa real foi incluída.

##✍️ Autor
**Kauan da Silva Terrão.**
