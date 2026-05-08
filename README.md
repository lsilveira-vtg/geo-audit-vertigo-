# GEO Audit Vertigo

Ferramenta web de auditoria SEO on-page construída com Python e Flask. Analisa title, meta description, H1/H2/H3, comprimento de URL e qualidade de conteúdo com insights de GEO para IAs.

## O que analisa

- **Title tag** — tamanho ideal (até 70 caracteres)
- **Meta description** — faixa ideal (140–160 caracteres)
- **H1** — presença e unicidade (exatamente 1 por página)
- **H2/H3** — estrutura hierárquica de conteúdo
- **URL** — comprimento e legibilidade
- **Conteúdo** — profundidade, dados quantitativos, linguagem de negócio, CTA
- **GEO** — otimização para IAs generativas (ChatGPT, Perplexity)
- **Palavras-chave** — extração automática por peso semântico

## Funcionalidades

- Cole uma ou várias URLs diretamente, ou faça upload de CSV/XLSX
- Dashboard interativo com cards de estatísticas e tabela filtável
- Histórico de auditorias salvo localmente
- Exportação dos resultados em XLSX com células coloridas por status
- Sugestões de pautas de conteúdo no posicionamento da Vertigo

## Como rodar

**Pré-requisito:** Python 3.9 ou superior instalado.

**1. Baixe o repositório**

```
git clone https://github.com/lsilveira-vtg/geo-audit-vertigo-.git
cd geo-audit-vertigo-
```

Ou clique em **Code → Download ZIP** e extraia a pasta.

**2. Instale as dependências**

```
pip install -r requirements.txt
```

**3. Inicie a aplicação**

```
python app.py
```

**4. Abra no navegador**

```
http://localhost:5050
```

## Stack

- **Backend:** Python + Flask
- **Parsing:** BeautifulSoup + lxml
- **Banco de dados:** SQLite
- **Export:** openpyxl
- **Frontend:** HTML, CSS e JavaScript
