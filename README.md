# 🤖 Validação NF - Automação de Auditoria Fiscal

Este projeto é uma solução de **RPA e Engenharia de Dados** desenvolvida em Python para automatizar a triagem, conciliação e auditoria fiscal de Notas Fiscais Eletrônicas de Serviço (NFS-e) e de Material (NF-e) do ERP TOTVS contra as bases governamentais e portais como o Quirius.

O projeto reduz de horas para minutos a carga diária da equipe Fiscal, eliminando processos repetitivos e mitigando passivos tributários através da identificação automática de retenções indevidas ou não realizadas.

---

## 🎯 Objetivo do Projeto

Atuar como um motor de **"Malha Fina" inteligente**. O pipeline cruza notas lançadas no ERP (TOTVS - Programas REB405 e OFB013) com os respectivos XMLs hospedados na base da Receita Federal e da Quirius. Quando o XML não está disponível, a ferramenta aciona um **Motor de Regras Fiscais Contingencial**, validando o código do serviço (COD ITEM) contra uma planilha de retenção oficial parametrizada pela equipe.

---

## 🛠️ Tecnologias e Bibliotecas Utilizadas

- **Linguagem Core:** Python 3.10+
- **Web Scraping & RPA:** Selenium (ChromeDriver Headless)
- **Processamento de Dados:** Pandas, OpenPyXL
- **Comunicações de Rede:** Requests, urllib3
- **Certificados Digitais (A1):** Cryptography (Conversão segura de PFX para PEM)
- **Notificações & Azure:** Microsoft Graph API (Envio automatizado via Outlook corporativo)
- **Gerenciamento de Ambientes:** python-dotenv, Pydantic-Settings

---

## 🏗️ Estrutura do Pipeline

A execução é coordenada pelo script principal em lote (`run_validacao.bat`). A arquitetura é dividida em cinco etapas principais:

### 1. Quirius Fetcher (`quirius_fetcher.py`)

Utiliza o Selenium para acessar silenciosamente o portal web do fornecedor Quirius, aplicando contornos avançados de rede (manipulação de iframe, contorno do erro de cookie `ERR_TOO_MANY_REDIRECTS` e gerenciamento de loops/timeout). Ele recorta a extração por blocos de 30 dias para evitar o travamento do site legado e faz o download unificado para a pasta temporária.

### 2. Sincronizador Nacional (`nfse_fetcher.py`)

Conecta-se diretamente ao Webservice oficial do Ambiente de Dados Nacional (ADN) da Receita Federal usando certificados A1 em formato `.pem`. Extrai, via gzip/base64 e retroativamente por NSU (Número Sequencial Único), os XMLs de serviço emitidos contra a empresa e os armazena em `.csv` na rede corporativa.

### 3. Motor de Auditoria de Materiais (`materials_etl.py`)

Cruza de forma inteligente os arquivos `.xlsx` extraídos do Quirius com as exportações `OFB013` do TOTVS. Utiliza o **Caçador Inteligente de Chaves** para contornar o problema de *Header Shift* (deslocamento de colunas) clássico das extrações TOTVS, garantindo que as chaves eletrônicas de 44 dígitos sejam localizadas independentemente da formatação.

### 4. Motor de Auditoria de Serviços (`services_etl.py`)

O núcleo da malha fina. Consolida:

- XMLs governamentais obtidos no passo anterior;
- Lançamentos brutos do ERP (`REB405`);
- Informações de optantes pelo Simples Nacional da base paralela. Caso o fornecedor seja desconhecido, consulta a **BrasilAPI** para verificar em tempo real o status tributário;
- **Motor Contingencial de Regras**, que aplica as condições definidas pelos analistas na planilha de regras (via `COD ITEM`) para determinar se impostos (IR, CS, INSS e ISS) foram retidos incorretamente ou deixaram de ser retidos.

Ao final, gera planilhas executivas (dashboards) utilizando `openpyxl`.

### 5. Notificador Global (`email_sender.py`)

Captura as planilhas consolidadas, monta um e-mail estruturado e conecta-se à infraestrutura Azure utilizando a Microsoft Graph API via OAuth2 (Client Credentials), enviando automaticamente as listas de notas divergentes para a equipe responsável.

---

# 🚀 Como Configurar e Rodar o Projeto

## Pré-requisitos

- Python 3.10 ou superior;
- Google Chrome atualizado juntamente com o ChromeDriver compatível;
- Arquivo `.env` configurado na raiz do projeto;
- Certificados digitais A1 (`.pfx`) disponíveis na pasta de dados.

## Configuração do arquivo `.env`

Crie um arquivo chamado `.env` na raiz do projeto:

```env
# Acesso ao Quirius
QUIRIUS_USER=seu_usuario
QUIRIUS_PASS=sua_senha
QUIRIUS_URL=http://10.0.10.57:600

# Azure / Graph API
EMAIL_REMETENTE=robo.fiscal@suaempresa.com.br
EMAIL_DESTINATARIO=analistas@suaempresa.com.br
TENANT_ID=id_do_azure_tenant
CLIENT_ID=id_do_app
CLIENT_SECRET=chave_secreta

# Certificados A1
EMPRESA_1_CNPJ=00000000000000
EMPRESA_1_CERT_NOME=empresa1.pfx
EMPRESA_1_CERT_SENHA=senha123
```

## Inicializando a Automação

No Windows, basta executar:

```cmd
run_validacao.bat
```

O script criará automaticamente o ambiente virtual, instalará as dependências e executará todo o pipeline.

Para executar apenas um módulo específico:

```bash
python src/nfse_fetcher.py
```

---

## 🛠️ Tratamento de Erros e Contingências

- **Deslocamento de Colunas TOTVS:** localização inteligente das colunas pela semântica ou comprimento da string (Caçador Inteligente de Chaves de 44 dígitos), evitando falhas quando o ERP altera a formatação das exportações.
- **Loop de Redirecionamento do Quirius:** limpeza automática de cookies via Selenium, nova autenticação e tratamento de respostas HTML inválidas.
- **Integração Segura com Microsoft Graph:** substituição do SMTP tradicional por autenticação OAuth2 utilizando Azure Application Permissions.
- **Conversão de Certificados:** leitura direta de arquivos `.pfx` e conversão transparente para `.pem` utilizando a biblioteca `cryptography`, compatível com o `requests`.