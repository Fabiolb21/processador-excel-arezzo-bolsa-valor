# Processador de Planilhas Excel - Streamlit

Este programa foi desenvolvido para processar arquivos Excel, extrair e transformar dados conforme especificações, gerar linhas duplicadas baseadas em quantidade, e exportar planilhas separadas por pedido em formato .xls.

## Funcionalidades

O programa realiza as seguintes transformações nos dados:

1. **PEDIDO**: Recebe a coluna PEDIDO da planilha original
2. **STYLE NAME**: Recebe a coluna STYLE_NAME da planilha original
3. **MATERIAL-COLOR**: Recebe a coluna MATERIAL_COLOR da planilha original
4. **SKU**: Recebe a coluna SKU da planilha original
5. **MODELO**: Recebe a coluna REFERENCIA_LOJA da planilha original
6. **EAN**: Recebe a coluna EAN da planilha original
7. **TAM**: Recebe a coluna TAM da planilha original
8. **QUANT**: Recebe a coluna QUANT da planilha original
9. **QUANT EXTRA**: Recebe o valor referente a coluna QUANT + 5
10. **VALOR**: Recebe a coluna MSRP da planilha original, formatado como R$ no formato contábil brasileiro
11. **NUM DA ETQ**: Contém uma numeração crescente começando em 000001 com zeros à frente
12. **VALOR DO FILTRO**: Recebe em todas as linhas o valor 1
13. **PREFIXO DA EMP**: Recebe os 7 primeiros caracteres da coluna EAN
14. **ITEM DE REF**: Recebe do oitavo ao décimo segundo caractere da coluna EAN acrescido de um zero na frente
15. **SERIAL**: Coluna vazia

### Duplicação de Linhas

O programa gera linhas duplicadas com os respectivos valores da linha anterior de acordo com a quantidade da coluna **QUANT EXTRA**. Por exemplo, se QUANT EXTRA for 56, serão geradas 56 linhas idênticas com aqueles dados.

### Geração de Planilhas

O programa gera planilhas de Excel separadas conforme a coluna **PEDIDO**, ou seja, uma planilha para cada PEDIDO diferente. Todas as planilhas são empacotadas em um arquivo ZIP para download. As colunas são formatadas como texto.

## Requisitos

- Python 3.11 ou superior
- Bibliotecas listadas em `requirements.txt`

## Instalação Local

### 1. Instalar Python

Certifique-se de ter o Python 3.11 ou superior instalado em seu sistema. Você pode verificar a versão com:

```bash
python --version
```

### 2. Instalar Dependências

No diretório do projeto, execute:

```bash
pip install -r requirements.txt
```

Ou instale manualmente:

```bash
pip install streamlit pandas openpyxl
```

## Como Usar Localmente

### 1. Executar a Aplicação

No diretório onde está o arquivo `app.py`, execute:

```bash
streamlit run app.py
```

### 2. Acessar a Aplicação

O Streamlit abrirá automaticamente seu navegador padrão. Caso não abra, acesse manualmente:

```
http://localhost:8501
```

### 3. Fazer Upload do Arquivo

1. Clique no botão **"Browse files"**
2. Selecione seu arquivo Excel (.xlsx ou .xlsm)
3. Aguarde o processamento

### 4. Baixar os Resultados

Após o processamento, você verá:
- Uma amostra dos dados processados (primeiras 50 linhas)
- Informação sobre quantos pedidos únicos foram encontrados
- Um botão **"Baixar todas as planilhas (.zip com arquivos .xls)"**

Clique no botão para baixar o arquivo ZIP contendo todas as planilhas separadas por pedido.

## Estrutura dos Arquivos de Saída

O arquivo ZIP baixado conterá:
- `pedido_XXXXXXXXXX.xls` - Um arquivo para cada número de pedido único
- Cada arquivo contém todas as linhas expandidas daquele pedido específico
- Todas as colunas estão formatadas como texto

## Deploy na Nuvem (Streamlit Cloud)

### 1. Criar Repositório no GitHub

1. Crie um novo repositório no GitHub
2. Faça upload dos arquivos:
   - `app.py`
   - `requirements.txt`
   - `README.md` (opcional)

### 2. Deploy no Streamlit Cloud

1. Acesse [share.streamlit.io](https://share.streamlit.io)
2. Faça login com sua conta GitHub
3. Clique em **"New app"**
4. Selecione seu repositório, branch e arquivo principal (`app.py`)
5. Clique em **"Deploy"**

A aplicação estará disponível em uma URL pública fornecida pelo Streamlit Cloud.

## Observações Técnicas

- O programa processa arquivos de até 200MB por padrão (limite do Streamlit)
- A duplicação de linhas pode gerar arquivos grandes dependendo das quantidades
- Os arquivos são gerados em formato .xls (compatível com Excel) usando a engine openpyxl
- Todas as colunas são convertidas para texto para garantir compatibilidade

## Suporte

Para dúvidas ou problemas, consulte a documentação do Streamlit em [docs.streamlit.io](https://docs.streamlit.io)
