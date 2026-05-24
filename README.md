<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=venom&height=260&color=0:020617,50:1e3a8a,100:16a34a&text=Sofascore%20Excel&fontSize=56&fontColor=ffffff&animation=fadeIn&fontAlignY=40"/>
</p>

<p align="center">
  ⚽ <strong>Automação para Coleta de Dados do Sofascore e Preenchimento de Planilha Excel</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Linguagem-Python-14532d?style=for-the-badge&logo=python&logoColor=white">
  <img src="https://img.shields.io/badge/Automação-Sofascore-14532d?style=for-the-badge">
  <img src="https://img.shields.io/badge/Status-Funcional-14532d?style=for-the-badge">
  <img src="https://img.shields.io/badge/Licença-MIT-14532d?style=for-the-badge">
</p>

<p align="center">
  👨‍🎓 <strong>Desenvolvido por:</strong> Vinícius Andrade Costa
</p>

---

## 📑 Sumário

- [Sobre o Projeto](#-sobre-o-projeto)
- [Arquitetura da Aplicação](#-arquitetura-da-aplicação)
- [Funcionalidades Implementadas](#️-funcionalidades-implementadas)
- [Funcionalidades Previstas](#-funcionalidades-previstas)
- [Principais Desafios](#️-principais-desafios)
- [Instalação e Execução](#-instalação-e-execução)
- [Como Usar](#-como-usar)
- [Tecnologias Utilizadas](#️-tecnologias-utilizadas)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Possíveis Erros e Soluções](#-possíveis-erros-e-soluções)
- [Observações Importantes](#-observações-importantes)

---

## ⚽ Sobre o Projeto

O **Sofascore → Excel** é uma aplicação em Python desenvolvida para automatizar a coleta de informações de partidas de futebol diretamente do **Sofascore** e registrar esses dados em uma planilha Excel.

O sistema permite capturar automaticamente a **URL da partida aberta no navegador Opera GX**, extrair o **ID do jogo**, consultar a API do Sofascore e preencher a planilha com informações como:

- 🕒 Horário da partida
- 🏟️ Estádio do jogo
- 📋 Linha disponível na planilha
- ✅ Registro automático dos dados coletados

A aplicação conta com uma interface gráfica simples e flutuante, facilitando o uso durante a navegação no Sofascore.

---

## 🧩 Arquitetura da Aplicação

O projeto foi desenvolvido com uma arquitetura simples, focada em automação local e integração com planilhas.

```text
┌─────────────────────────────────────────────┐
│                  Usuário                     │
│     abre o jogo no Sofascore (Opera GX)      │
└─────────────────────┬───────────────────────┘
                      ⬇
┌─────────────────────────────────────────────┐
│            Interface Tkinter                 │
│   botão de captura automática ou manual      │
└─────────────────────┬───────────────────────┘
                      ⬇
┌─────────────────────────────────────────────┐
│              Script Python                   │
│   extrai ID e consulta API do Sofascore      │
└─────────────────────┬───────────────────────┘
                      ⬇
┌─────────────────────────────────────────────┐
│             API Sofascore                    │
│       retorna dados oficiais da partida      │
└─────────────────────┬───────────────────────┘
                      ⬇
┌─────────────────────────────────────────────┐
│              Planilha Excel                  │
│   horário e estádio gravados automaticamente │
└─────────────────────────────────────────────┘
```

Essa estrutura permite um fluxo rápido para preencher bases de dados esportivas com menor risco de erro manual.

---

## 🏟️ Funcionalidades Implementadas

### 📋 Captura automática da URL

- Localiza a janela ativa do Opera GX
- Foca automaticamente no navegador
- Executa os comandos `Ctrl + L` e `Ctrl + C`
- Copia a URL da partida aberta no Sofascore
- Valida se o conteúdo copiado é uma URL válida

### 📎 Modo manual via clipboard

Além da captura automática, o sistema também permite o uso manual. Nesse modo, o usuário pode:

1. Abrir a partida no Sofascore
2. Copiar a URL manualmente
3. Clicar no botão de modo manual
4. Deixar o sistema consultar os dados e preencher a planilha

### 🔎 Extração do ID da partida

O sistema identifica o ID da partida a partir da URL do Sofascore. São tratados formatos como:

- `id:12345678`
- Links antigos contendo o número da partida no final da URL

### 🌐 Consulta à API do Sofascore

Após identificar o ID da partida, o sistema consulta diretamente o endpoint:

```http
GET https://api.sofascore.com/api/v1/event/{match_id}
```

A partir da resposta da API, são extraídos:

- Horário de início da partida
- Informações do estádio
- Nome do local do jogo, quando disponível

### 🕒 Conversão de horário

O horário da partida é obtido a partir do campo `startTimestamp` retornado pela API. O script converte o timestamp para o formato `HH:MM`.

**Exemplo:** `21:30`

### 🏟️ Coleta do estádio

O sistema busca o nome do estádio dentro das informações de `venue` retornadas pela API. A lógica prioriza:

1. Nome do estádio
2. Nome geral do local
3. Retorno vazio caso a API não envie a informação

### 🗄️ Escrita automática na planilha Excel

O sistema abre a planilha configurada no código e procura a primeira linha vazia da coluna de horário. Depois disso, grava automaticamente:

- Horário na coluna definida
- Estádio na coluna definida

**Configuração padrão:**

```python
COLUNA_HORARIO = "A"
COLUNA_ESTADIO = "F"
LINHA_INICIAL_BUSCA = 2
```

### 🖥️ Interface gráfica flutuante

A aplicação possui uma interface criada com Tkinter, apresentando:

- Botão de captura automática
- Botão de uso manual do clipboard
- Área de logs em tempo real
- Mensagens de sucesso e erro
- Indicação da linha preenchida na planilha

### ⚙️ Outras Funcionalidades

- ✅ Interface gráfica com janela sempre no topo
- ✅ Logs em tempo real dentro da aplicação
- ✅ Execução da automação em thread separada para evitar travamentos
- ✅ Tratamento de erros para URL inválida
- ✅ Tratamento de erro caso a planilha esteja aberta
- ✅ Tratamento de erro caso a aba configurada não exista
- ✅ Uso de headers personalizados na requisição
- ✅ Consulta com `curl_cffi` utilizando impersonação de navegador Chrome

---

## 🚧 Funcionalidades Previstas

Algumas melhorias podem ser implementadas em futuras versões:

- [ ] Selecionar a planilha diretamente pela interface
- [ ] Escolher a aba do Excel sem editar o código
- [ ] Registrar nome dos times automaticamente
- [ ] Registrar competição, rodada e data da partida
- [ ] Processar múltiplas URLs em lote
- [ ] Exportar logs para arquivo `.txt`
- [ ] Criar histórico de partidas já coletadas
- [ ] Evitar duplicidade de jogos na planilha
- [ ] Adicionar suporte para outros navegadores além do Opera GX

---

## ⚠️ Principais Desafios

Durante o desenvolvimento do projeto, alguns desafios técnicos foram considerados.

### 🔗 Integração com o Sofascore

O Sofascore pode limitar ou bloquear requisições dependendo do ambiente, headers ou IP utilizado. Por isso, o projeto utiliza:

- `curl_cffi`
- Headers personalizados
- Impersonação de navegador

### 📋 Automação do navegador

Para capturar a URL automaticamente, o sistema depende da janela do Opera GX estar aberta e com a partida ativa. Esse processo envolve:

- Localizar a janela correta
- Focar o navegador
- Simular atalhos de teclado
- Acessar o conteúdo copiado no clipboard

### 🗄️ Escrita segura no Excel

O sistema precisa garantir que a planilha:

- Exista no caminho configurado
- Tenha a aba correta
- Não esteja aberta no Excel no momento da escrita
- Possua linhas disponíveis para preenchimento

### 🧠 Tratamento de dados da API

Nem todas as partidas retornam os mesmos campos completos na API. Por isso, o sistema tenta buscar o estádio em mais de uma estrutura possível dentro do retorno da API.

---

## 🚀 Instalação e Execução

### 🔧 Pré-requisitos

Antes de iniciar, é necessário ter instalado:

- 🐍 **Python 3.10** ou superior
- 🌐 **Opera GX**
- 📊 **Microsoft Excel** ou programa compatível
- 📄 Planilha `.xlsx` configurada
- 🌍 Acesso à internet

### 📦 Instalação das dependências

No terminal, acesse a pasta do projeto e execute:

```bash
pip install curl_cffi openpyxl pyperclip pygetwindow keyboard
```

### ⚙️ Configuração da planilha

Antes de executar o sistema, edite as configurações no início do arquivo `automacao_sofascore.py`:

```python
CAMINHO_EXCEL = r"C:\Users\costa\Downloads\libertadores_brasileiros_serie_a_2021_2025.xlsx"
NOME_ABA = "Libertadores 2021-2025"
COLUNA_HORARIO = "A"
COLUNA_ESTADIO = "F"
LINHA_INICIAL_BUSCA = 2
```

#### Campos configuráveis

| Campo | Descrição |
|-------|-----------|
| `CAMINHO_EXCEL` | Caminho completo da planilha Excel |
| `NOME_ABA` | Nome da aba onde os dados serão gravados |
| `COLUNA_HORARIO` | Coluna onde será registrado o horário |
| `COLUNA_ESTADIO` | Coluna onde será registrado o estádio |
| `LINHA_INICIAL_BUSCA` | Primeira linha onde o sistema começará a procurar espaço vazio |

### ▶️ Como executar

No terminal, execute:

```bash
python automacao_sofascore.py
```

Após executar, será aberta a janela da aplicação.

---

## 📖 Como Usar

### 📋 Modo automático

Para usar o modo automático:

1. Abra o **Opera GX**
2. Acesse a página da partida no **Sofascore**
3. Deixe a aba da partida ativa
4. Clique no botão **📋 Copiar este jogo (automático)**

O sistema irá:

- Focar a janela do Opera
- Copiar a URL
- Consultar a API do Sofascore
- Extrair horário e estádio
- Gravar os dados na planilha

### 📎 Modo manual

Para usar o modo manual:

1. Abra a partida no **Sofascore**
2. Pressione `Ctrl + L`
3. Pressione `Ctrl + C`
4. Volte para a aplicação
5. Clique em **📎 Usar URL do clipboard (manual)**

> 💡 Esse modo é útil caso o sistema não consiga focar automaticamente a janela do Opera GX.

### ✅ Exemplo de fluxo de uso

```text
1. Abrir a partida no Sofascore
2. Clicar em "Copiar este jogo"
3. O sistema captura a URL
4. O ID da partida é extraído
5. A API do Sofascore é consultada
6. Horário e estádio são encontrados
7. Os dados são gravados na planilha
8. O sistema informa a linha preenchida
```

---

## 🛠️ Tecnologias Utilizadas

<table>
  <tr>
    <td><strong>Linguagem</strong></td>
    <td>Python</td>
  </tr>
  <tr>
    <td><strong>Interface Gráfica</strong></td>
    <td>Tkinter</td>
  </tr>
  <tr>
    <td><strong>Automação</strong></td>
    <td>PyGetWindow · Keyboard · Pyperclip</td>
  </tr>
  <tr>
    <td><strong>Requisições HTTP</strong></td>
    <td>curl_cffi</td>
  </tr>
  <tr>
    <td><strong>Manipulação de Planilhas</strong></td>
    <td>OpenPyXL</td>
  </tr>
  <tr>
    <td><strong>Dados</strong></td>
    <td>API do Sofascore</td>
  </tr>
</table>

---

## 📁 Estrutura do Projeto

```text
sofascore-excel/
├── automacao_sofascore.py    # Script principal da aplicação
└── README.md                  # Documentação do projeto
```

---

## 🧪 Possíveis Erros e Soluções

### ❌ Erro: Janela do Opera não encontrada

> Verifique se o Opera GX está aberto.

**Solução:**

- Abra o Opera GX
- Deixe a partida do Sofascore aberta
- Tente novamente

### ❌ Erro: Clipboard vazio

> O sistema não conseguiu copiar ou ler a URL.

**Solução:**

- Use o modo manual
- Copie a URL com `Ctrl + L` e `Ctrl + C`
- Clique no botão verde da aplicação

### ❌ Erro: Planilha não encontrada

> O caminho configurado no código está incorreto.

**Solução:**

- Confira o valor de `CAMINHO_EXCEL`
- Verifique se o arquivo existe no local informado

### ❌ Erro: Não consegui salvar o Excel

> A planilha provavelmente está aberta.

**Solução:**

- Feche o arquivo Excel
- Clique novamente no botão da aplicação

### ❌ Erro: Aba não existe

> O nome da aba configurada no código não corresponde ao nome da aba real da planilha.

**Solução:**

- Abra a planilha
- Confira o nome exato da aba
- Atualize a variável `NOME_ABA`

---

## 📌 Observações Importantes

> ⚠️ Este projeto foi desenvolvido para **uso local e acadêmico**.

A aplicação depende da disponibilidade da API do Sofascore e do funcionamento correto do navegador utilizado.

Recomenda-se:

- Respeitar os termos de uso das plataformas consultadas
- Evitar requisições excessivas
- Utilizar a automação apenas de forma moderada

---
