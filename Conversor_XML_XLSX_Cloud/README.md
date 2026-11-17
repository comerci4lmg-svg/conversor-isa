# Conversor XML para XLSX

## Como abrir no Windows
1. Instale Python 3.10 ou superior, com o Python Launcher (comando `py`).
2. Extraia todo o ZIP para uma pasta.
3. Dê dois cliques em `iniciar.cmd`.
4. Na primeira execução, é necessária internet para instalar as dependências. O aplicativo abrirá no navegador.
5. Envie os XMLs, confira o grupo de registros e a prévia, clique em **Converter para XLSX** e baixe.

Alternativa pelo terminal, dentro da pasta:
```
py -m pip install -r requirements.txt
py -m streamlit run streamlit_app.py
```

## Estrutura dos dados
Cada XML gera um XLSX com a aba Dados. É possível baixar todos juntos em ZIP.
O grupo sugerido prioriza elementos repetidos com campos internos. Confira a seleção: cada elemento do grupo vira uma linha, e apenas os dados dentro dele são exportados.
Campos aninhados usam caminhos separados por `/`; atributos usam `@`; elementos repetidos usam índices `[1]`, `[2]` etc. Namespaces permanecem nos nomes para evitar misturar campos diferentes.
Textos mistos após elementos aparecem em colunas `#tail`. Espaços nas bordas dos textos são removidos.
Todos os valores são texto, inclusive números e datas, para preservar códigos, zeros à esquerda e TRUE/FALSE. Conteúdo começando com `=` não é executado como fórmula.

XML é hierárquico: a seleção de registros é uma projeção em tabela, não uma cópia integral do documento. Para NF-e ou outros layouts específicos, selecione o grupo de itens desejado; dados de pais fora desse grupo não são repetidos automaticamente.

Limites: 50 MB por XML; 1.048.575 registros mais o cabeçalho; 16.384 colunas; 32.767 caracteres por célula. O processamento ocorre na memória do computador que executa o aplicativo. DTDs e entidades são rejeitados.

Use `exemplo.xml` para experimentar.

## XML de Excel (SpreadsheetML 2003)
Arquivos com raiz Workbook no namespace de planilhas Microsoft são reconhecidos automaticamente. A primeira linha vira cabeçalho; linhas e células com índices explícitos mantêm as posições vazias. Cada aba gera um XLSX separado. Cabeçalhos vazios ou repetidos recebem nomes únicos. A conversão preserva os valores como texto, incluindo o resultado armazenado de fórmulas; não preserva fórmulas, estilos ou mesclagens originais.
