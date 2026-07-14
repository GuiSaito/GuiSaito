# News Digest

Pipeline pessoal de curadoria de notícias que coleta artigos de múltiplas fontes, aplica filtros de qualidade e deduplicação, gera análises com LLM por meio de personas especializadas e entrega um digest diário por e-mail.

Às sextas-feiras, o projeto também produz um **Weekly Intel** separado, sintetizando os principais sinais da semana e conectando os temas monitorados.

## Objetivo

O projeto foi criado para reduzir o tempo gasto na leitura de feeds dispersos e centralizar informações relevantes em um único digest.

O fluxo automatiza:

- coleta de notícias via RSS e NewsAPI;
- limpeza e filtragem dos artigos;
- deduplicação dentro e entre tópicos e dias;
- diversificação de fontes;
- análise com LLM usando personas por área;
- montagem de e-mail HTML navegável;
- envio pelo Outlook desktop;
- armazenamento local de digests, logs e histórico semanal.

## Tópicos monitorados

O pipeline organiza as notícias em oito áreas:

1. Geopolítica e Economia Global
2. Finanças
3. Tecnologia, Consultoria e Mercado
4. Inteligência Artificial
5. Agro
6. Cibersegurança
7. Saúde e Biotech
8. Carreira de Dados

Cada tópico possui uma persona e instruções próprias de análise. A configuração de tópicos, feeds, consultas e palavras-chave fica externalizada em `topics.json`.

## Como funciona

O projeto segue um pipeline ETL leve:

### 1. Extract

- coleta RSS com `feedparser`;
- consulta à NewsAPI em português, inglês, espanhol e alemão;
- execução paralela da coleta com `ThreadPoolExecutor`.

### 2. Transform

- remoção de HTML dos resumos;
- descarte de artigos com conteúdo insuficiente;
- deduplicação por URL;
- deduplicação semântica por similaridade de títulos;
- deduplicação entre tópicos;
- deduplicação entre dias da mesma semana;
- filtragem por palavras-chave quando aplicável;
- diversificação de fontes por round-robin;
- sumarização paralela com LLM e instruções específicas por tópico.

### 3. Load

- criação de e-mail HTML com índice clicável;
- envio pelo Outlook desktop via `win32com`;
- salvamento de cópia HTML local;
- geração de log CSV mensal;
- arquivamento de contexto semanal.

## Pipeline diário

Em uma execução típica, o script:

1. valida as variáveis de ambiente necessárias;
2. define a janela de coleta;
3. coleta artigos de todos os tópicos em paralelo;
4. limpa, filtra e deduplica o conteúdo;
5. registra as URLs processadas na semana;
6. ignora chamadas ao LLM para tópicos vazios;
7. gera análises por tópico com rate limiting e retry;
8. verifica itens previamente marcados para acompanhamento;
9. monta e salva o digest em HTML;
10. envia o e-mail pelo Outlook;
11. registra o resultado em log.

Às sextas-feiras, o fluxo também gera o **Weekly Intel**, com:

- Quick Brief da semana;
- sinais fracos e assuntos para monitorar;
- mesa redonda entre as personas;
- pergunta da semana.

## Principais decisões técnicas

### Qualidade da curadoria

- deduplicação semântica por título com similaridade de Jaccard;
- deduplicação cross-topic e cross-day;
- preferência pelo artigo com resumo mais completo;
- diversificação por round-robin, evitando concentração excessiva de uma única fonte;
- descarte de conteúdos curtos demais para análise;
- supressão de tópicos sem artigos após os filtros.

### Uso de LLM

- mensagens de sistema e usuário separadas;
- instrução explícita para usar apenas os artigos fornecidos;
- prompts especializados por tópico;
- formato estruturado por evento;
- atribuição natural das fontes;
- seção final de pontos para acompanhamento;
- rate limiter para controlar rajadas de requisições;
- retry com backoff exponencial.

### E-mail e automação

- integração com Outlook desktop, evitando autenticação SMTP separada;
- conversão de Markdown para HTML;
- índice clicável por tópico;
- cópias HTML organizadas por ano;
- execução automatizada em dias úteis pelo Agendador de Tarefas do Windows.

## Estrutura esperada de arquivos

O repositório público contém apenas o código e a documentação do projeto. Arquivos locais com credenciais, entradas reais, logs e saídas não devem ser versionados.

```text
.
├── news_digest_v1.1.py
├── README.md
├── PROJECT_HISTORY.md
├── requirements-global.txt
└── .gitignore
```

Na instalação local, o script também utiliza arquivos externos, como:

```text
Inputs/
├── .env
├── topics.json
├── urls_semana.json
├── semana_atual.json
└── para_monitorar.md
```

> Os arquivos reais de entrada e credenciais não fazem parte do repositório público.

## Configuração

Crie um arquivo `.env` local com as variáveis necessárias:

```dotenv
GROQ_API_KEY=sua_chave_groq
NEWSAPI_KEY=sua_chave_newsapi
EMAIL_TO=seu_email_de_destino
```

Também é necessário criar um `topics.json` com os tópicos, feeds, consultas e palavras-chave desejados.

Nunca publique o `.env`, chaves de API, endereços pessoais ou arquivos de entrada reais.

## Dependências

O arquivo `requirements-global.txt` representa o ambiente Python compartilhado utilizado em scripts pessoais. Nem todas as bibliotecas listadas nele são necessariamente exigidas por este projeto.

Entre as dependências mencionadas na implementação estão:

- `feedparser`
- `requests`
- `python-dotenv`
- `groq`
- `pywin32`
- `markdown`

Para instalar o ambiente global atual:

```bash
python -m pip install -r requirements-global.txt
```

Como melhoria futura, o projeto poderá manter um `requirements.txt` contendo somente as dependências efetivamente utilizadas pelo News Digest.

## Execução

Com as dependências e os arquivos locais configurados:

```bash
python news_digest_v1.1.py
```

O envio por `win32com` pressupõe o Outlook desktop disponível e configurado no Windows.

## Segurança e privacidade

Este repositório não deve conter:

- arquivos `.env`;
- chaves de API ou tokens;
- endereços de e-mail pessoais;
- caminhos absolutos do computador do autor;
- inputs e outputs reais;
- logs locais;
- documentos ou dados corporativos;
- estudos ou artefatos vinculados a clientes.

Exemplo de `.gitignore` recomendado:

```gitignore
.env
.env.*
!.env.example

Inputs/
Relatorios/
logs/
outputs/
*.log

__pycache__/
*.py[cod]
.venv/
venv/

.vscode/
.idea/
.DS_Store
Thumbs.db
desktop.ini
```

## Histórico técnico

As decisões, correções e evoluções do projeto são mantidas em `PROJECT_HISTORY.md`.

Entre as melhorias documentadas estão:

- expansão do contexto enviado ao LLM;
- filtro mínimo de conteúdo;
- deduplicação semântica e entre dias;
- separação entre instruções de sistema e conteúdo dos artigos;
- proteção explícita contra alucinação;
- rate limiting para chamadas ao LLM;
- diversificação de fontes;
- acompanhamento de temas da semana anterior;
- acumulação semanal para o Weekly Intel;
- redução do contexto semanal para respeitar limites da API.

## Limitações conhecidas

- os planos gratuitos das APIs impõem limites de requisições, tokens e atualização;
- alguns feeds RSS não fornecem timestamp confiável;
- o LLM analisa apenas o conteúdo entregue pelo pipeline;
- a deduplicação entre tópicos depende da ordem definida no arquivo de configuração;
- o Weekly Intel depende da qualidade e disponibilidade dos resumos diários;
- a execução de e-mail depende do Outlook desktop no Windows.

## Próximas melhorias

- criar `requirements.txt` específico do projeto;
- disponibilizar exemplos sanitizados de `.env` e `topics.json`;
- substituir caminhos absolutos por configuração portátil;
- adicionar testes para filtros e deduplicação;
- separar configurações, regras de negócio e envio de e-mail em módulos menores;
- documentar a instalação automatizada no Agendador de Tarefas.

## Status

Projeto pessoal funcional, mantido de forma incremental e utilizado para automatizar a curadoria recorrente de notícias.

## Autor

Guilherme Saito — [GitHub](https://github.com/GuiSaito)
