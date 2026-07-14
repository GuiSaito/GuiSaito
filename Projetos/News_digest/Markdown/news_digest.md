# News Digest — Script de Curadoria de Notícias

**Projeto:** Pessoal  
**Data:** 2026-06-15  
**Versão:** 1.1

## Objetivo
Coletar automaticamente notícias relevantes por tópico de interesse, sumarizar com IA via personas especializadas, e enviar um digest diário por e-mail. Às sextas, um segundo e-mail separado ("Weekly Intel") sintetiza a semana com mesa redonda entre as personas. Reduzir o tempo gasto em leitura de feeds e centralizar informação relevante num único lugar.

## Tópicos monitorados (8) — ordem no e-mail
| # | Tópico | Persona | Filtro keywords |
|---|--------|---------|----------------|
| 1 | Geopolítica & Economia Global | Maquiavel | `false` |
| 2 | Finanças | Midas | `true` (PT/EN/ES/DE) |
| 3 | Tech, Consultoria & Mercado | Edison | `false` |
| 4 | Inteligência Artificial | Alan | `false` |
| 5 | Agro | Deméter | `true` (PT/EN/ES/DE) |
| 6 | Cibersegurança | Loki | `false` |
| 7 | Saúde & Biotech | Asclépio | `false` |
| 8 | Carreira de Dados | Florence | `false` |

> Ordem: macro → setor → carreira. Somente Agro e Finanças têm filtro ativo pois usam RSS de fontes generalistas.

## Dados
- **Fonte:** RSS feeds + NewsAPI (4 idiomas: PT, EN, ES, DE)
- **Período:** Últimas 48h (72h às segundas para cobrir fim de semana); execução de **segunda a sexta às 9:15** via Task Scheduler
- **Granularidade:** Por artigo — título, fonte, resumo (truncado a 700 chars), link
- **Tratamentos aplicados:** Strip HTML, deduplicação intra-tópico por URL, deduplicação semântica por título (Jaccard ≥ 70%), deduplicação cross-tópico, deduplicação cross-day (via `Inputs/urls_semana.json`), filtro de artigos com resumo < 80 chars, diversificação de fontes por round-robin, filtragem por word boundary regex (quando `usar_keywords: true`), sumarização via LLM
- **Configuração de tópicos:** externalizada em `Inputs/topics.json`

## Fundamentação Teórica
O script segue o padrão de **newsletter automatizada** via pipeline ETL leve:

1. **Extract** — RSS feeds (`feedparser`) + NewsAPI por idioma (`requests`), em paralelo via `ThreadPoolExecutor`
2. **Transform** — Strip HTML, deduplicação, filtragem por keyword com `re.search` (word boundary), sumarização via LLM com persona especializada
3. **Load** — E-mail HTML via Outlook desktop (`win32com`) + cópia HTML local + log CSV mensal

O LLM recebe até 20 artigos por tópico (após todos os filtros de qualidade) e usa um `prompt_override` específico por área, gerando análise com voz especializada em vez de resumo genérico.

## Implementação

### Arquivos principais
- **Script:** `news_digest_v1.1.py`
- **Config de tópicos:** `Inputs/topics.json`
- **Credenciais:** `Inputs/.env` (`GROQ_API_KEY`, `NEWSAPI_KEY`, `EMAIL_TO`)
- **Logs:** `Relatorios/log news script/log_YYYY-MM.csv`
- **Digests HTML:** `Relatorios/digests/YYYY/digest_YYYY-MM-DD.html`
- **URLs da semana:** `Inputs/urls_semana.json` (cross-day dedup; resetado toda sexta)
- **Monitoramento:** `Inputs/para_monitorar.md` (itens da semana anterior para acompanhar)

### Pipeline diário
1. Validar variáveis de ambiente (aborta com mensagem clara se ausente)
2. Calcular `HORAS_ATRAS` (72h segunda, 48h demais dias)
3. Coletar todos os tópicos em paralelo (8 workers) — RSS + NewsAPI
4. Filtrar artigos com resumo < 80 chars (sem conteúdo útil)
5. Deduplicar intra-tópico por URL
6. Deduplicar semanticamente por título (Jaccard ≥ 70%, mantém o mais completo)
7. Diversificar fontes por round-robin (garante diversidade nas primeiras posições)
8. Deduplicar cross-tópico (artigo vai para o primeiro tópico que o capturou)
9. Deduplicar cross-day (remove URLs já cobertas nos dias anteriores da semana)
10. Registrar URLs desta execução em `urls_semana.json`
11. Pular chamada LLM para tópicos com 0 artigos após dedup
12. Sumarizar tópicos com artigos em paralelo (4 workers + rate limiter 15 req/min)
13. Verificar se itens de `para_monitorar.md` têm cobertura hoje — gerar seção `🔭 Acompanhamento`
14. Montar e-mail HTML com índice clicável + seção de acompanhamento + seções por tópico
15. Salvar cópia HTML local (`Relatorios/digests/YYYY/`)
16. Enviar via Outlook desktop (`win32com`)
17. Salvar log CSV mensal
18. *(Sextas apenas)* Gerar e enviar Weekly Intel separado + arquivar semana + resetar acumuladores

### Pipeline semanal (sextas)
- Após o digest diário, passa os 8 resumos das personas ao LLM
- LLM gera: **Quick Brief** (tema da semana, sinais fracos, para monitorar) + **Mesa Redonda** (diálogo casual entre as 8 personas) + **Pergunta da semana**
- Enviado como e-mail separado com subject: `"Weekly Intel — [Tema gerado pelo LLM]"`

### Flags de controle no topo do script
```python
USAR_NEWSAPI = True/False   # False durante testes para poupar cota (100 req/dia)
AGRUPAR_POR_SUBTEMA = True  # agrupa por sub-tema quando >= 3 artigos
HORAS_ATRAS                 # calculado automaticamente (72h segunda, 48h demais)
```

## Decisões técnicas

### LLM e API
- **Groq (gratuito) ao invés de OpenAI** — quota da OpenAI esgotou; Groq oferece Llama 3.3 70B grátis com free tier generoso (~500k tokens/dia) e SDK compatível com OpenAI
- **Rate limiter explícito (`_RateLimiter`, 15 req/min)** — `ThreadPoolExecutor` com 4 workers pode disparar bursts que excedem o limite de tokens/min do Groq; classe usa `threading.Lock` para garantir intervalo mínimo de 4s entre chamadas LLM; trata a causa raiz em vez de depender apenas do retry
- **4 workers paralelos para LLM** — balanceia paralelismo com respeito ao rate limiter; coleta usa 8 workers (sem restrição de tokens)
- **Retry com backoff exponencial** — 3 tentativas (2s → 4s → 8s) antes de desistir; fallback para erros que escapem do rate limiter
- **`temperature=0.3` para resumos, `0.5` para Weekly Intel** — resumos precisam de consistência; mesa redonda precisa de alguma variabilidade mas menos do que antes (era 0.7) para reduzir alucinações
- **`max_tokens=700` resumos, `1200` Weekly Intel** — aumento de 600→700 para acomodar o novo formato estruturado com título em negrito por evento

### Coleta e filtragem
- **Word boundary (`\b`) nas keywords** — filtragem por substring simples gerava falsos positivos (ex: `\bIA\b` em modo case-insensitive batia no verbo "ia"); regex com `\b` exige palavra isolada
- **`usar_keywords: false`** em 6 tópicos cujos RSS já são temáticos (Geopolítica, Tech, IA, Carreira, Ciber, Saúde) — filtro só faz sentido em feeds generalistas (Agro e Finanças)
- **Strip HTML antes de filtrar e enviar ao LLM** — feeds RSS frequentemente incluem tags `<p>`, `<img>` etc. no campo summary; removidas via `re.sub`
- **Truncar resumos a 700 chars (era 400)** — 400 chars cortava muito cedo em artigos ricos; 700 chars aumenta o contexto sem impacto relevante no custo de tokens
- **Filtro de conteúdo mínimo (< 80 chars descartado)** — muitos artigos de NewsAPI têm `description=None` ou só o título repetido; sem conteúdo útil não contribuem para a análise e poluem o prompt do LLM
- **Deduplicação semântica por título (Jaccard ≥ 70%)** — artigos sobre o mesmo evento têm títulos muito similares mas URLs distintas; dedup por URL não os captura; similaridade de palavras-chave resolve; mantém o artigo com resumo mais longo (mais completo)
- **Diversificação por round-robin ao invés de cap fixo por fonte** — cap fixo (ex: 3 por fonte) descartava artigos genuinamente distintos de fontes prolíficas (Reuters com 8 artigos sobre eventos diferentes); round-robin intercala 1 de cada fonte por turno, garantindo diversidade nas primeiras posições sem desperdiçar artigos válidos
- **Deduplicação cross-day (`urls_semana.json`)** — sem isso o mesmo artigo aparecia no digest de terça e quarta se continuasse no top do RSS; histórico semanal de URLs impede repetição; resetado toda sexta com o acumulador de resumos
- **Deduplicação cross-tópico** — artigos que aparecem em múltiplos tópicos (ex: "Bayer lança IA para agro") ficam apenas no primeiro tópico que os capturou (ordem do JSON)
- **`HORAS_ATRAS` automático às segundas** — fim de semana tem menos publicações; 72h garante cobertura adequada sem configuração manual
- **Pular LLM para tópicos vazios** — após todas as deduplicações um tópico pode chegar com 0 artigos; chamar o LLM resultava em mensagem genérica "Nenhuma notícia..." e desperdício de cota; agora a seção é suprimida do email e a chamada é pulada

### LLM — Qualidade de output
- **System message separado do user message em `sumarizar()`** — misturar instruções de persona com os artigos no mesmo `user` message reduzia o peso das instruções; separar em `system`/`user` melhora a aderência às regras de formato
- **Anti-alucinação explícita no system message** — instrução: "Use APENAS as informações dos artigos fornecidos. Não adicione conhecimento geral."; sem isso o LLM complementava lacunas com informação inventada
- **Proibição de enchimentos** — frases como "é importante notar", "vale ressaltar" foram banidas explicitamente no system message; reduz ruído e aumenta densidade informacional
- **Formato estruturado com título por evento** — output esperado: `**[Título curto]** — 2-3 frases.` por evento distinto; mais escaneável do que bullets genéricos
- **Curadoria explícita solicitada ao LLM** — instrução: "Para múltiplos artigos sobre o mesmo evento, use apenas o mais completo."; evita repetição mesmo quando a dedup semântica não eliminou 100%
- **Atribuição de fonte natural** — instrução: "Mencione as fontes naturalmente (ex: 'Segundo o Reuters...')."; aumenta credibilidade e rastreabilidade
- **`🔎 Para ficar de olho:`** — seção obrigatória ao final de cada tópico; força o LLM a identificar o ponto de atenção mais relevante em vez de terminar no último bullet
- **Anti-alucinação no Weekly Intel** — sinais fracos e itens de monitoramento devem referenciar artigos específicos da semana; temperatura reduzida de 0.7 → 0.5 para menos especulação

### E-mail — Ciclo de inteligência
- **Seção `🔭 Acompanhamento`** — no início do digest, antes das seções por tópico; aparece quando algum item de `para_monitorar.md` tem artigos correspondentes hoje (matching por palavras-chave, Jaccard ≥ 30%, mín. 2 hits); fecha o loop entre o que foi monitorado na semana anterior e o que saiu hoje

### E-mail e formatação
- **win32com ao invés de SMTP** — conta usa MFA com Microsoft Authenticator; win32com aciona o Outlook desktop diretamente, sem senha
- **`<div>` ao invés de `<p>` no wrapper do resumo** — `markdown.markdown()` já retorna `<p>` internamente; wrapper `<p>` gerava HTML inválido (`<p><p>...</p></p>`) que Outlook renderizava errado
- **Pré-processamento de listas Markdown** — a biblioteca `markdown` exige linha em branco antes de itens `*`/`-` para convertê-los em `<ul><li>`; sem isso saem como asteriscos literais; resolvido com `re.sub` antes de cada chamada `markdown.markdown()`
- **Índice clicável no topo** — âncoras HTML para cada seção; útil para navegar direto ao tópico de interesse
- **Cópias HTML salvas por ano** — `Relatorios/digests/YYYY/digest_YYYY-MM-DD.html`; histórico consultável localmente

### Personas e Weekly Intel
- **`prompt_override` por tópico** — cada área tem uma persona com ângulo específico de análise; substitui o prompt genérico de resumo de notícias
- **Nomes das personas baseados em figuras históricas/mitológicas** — identidade memorável e associação intuitiva com a área
- **Conexões da Mesa Redonda emergem do conteúdo, não de template** — prompt não fixa quais tópicos influenciam quais; o LLM descobre as conexões reais daquela semana

### topics.json
- **Externalizado do script** — adicionar/remover tópicos, feeds e keywords sem alterar código
- **Ordem dos tópicos define ordem no e-mail** — Python dicts mantêm ordem de inserção (3.7+)
- **queries NewsAPI bilíngues** — termos concretos que aparecem em títulos reais (ex: `"Trump OR NATO OR tariffs"`) em vez de frases compostas raramente usadas
- **Keywords multilíngue em Agro e Finanças** — PT + EN + ES + DE para capturar artigos internacionais que passam pelo NewsAPI

## Fontes RSS ativas por tópico

| Tópico | Fonte | URL RSS |
|--------|-------|---------|
| Geopolítica | BBC World | `https://feeds.bbci.co.uk/news/world/rss.xml` |
| Geopolítica | NYT World | `https://rss.nytimes.com/services/xml/rss/nyt/World.xml` |
| Geopolítica | Reuters World | `https://feeds.reuters.com/reuters/worldNews` |
| Geopolítica | Al Jazeera | `https://www.aljazeera.com/xml/rss/all.xml` |
| Finanças | InfoMoney | `https://www.infomoney.com.br/feed/` |
| Finanças | Exame | `https://exame.com/feed/` |
| Finanças | Agência Brasil Economia | `https://agenciabrasil.ebc.com.br/economia/feed/atom` |
| Tech & Consultoria | Ars Technica | `https://feeds.arstechnica.com/arstechnica/index` |
| Tech & Consultoria | Reuters Technology | `https://feeds.reuters.com/reuters/technologyNews` |
| Tech & Consultoria | TechCrunch | `https://techcrunch.com/feed/` |
| Tech & Consultoria | Harvard Business Review | `https://feeds.hbr.org/harvardbusiness` |
| Inteligência Artificial | MIT Tech Review | `https://www.technologyreview.com/feed/` |
| Inteligência Artificial | The Verge AI | `https://www.theverge.com/ai-artificial-intelligence/rss/index.xml` |
| Inteligência Artificial | VentureBeat AI | `https://venturebeat.com/category/ai/feed/` |
| Inteligência Artificial | AI News | `https://artificialintelligence-news.com/feed/` |
| Agro | Canal Rural | `https://www.canalrural.com.br/feed/` |
| Agro | Notícias Agrícolas | `https://www.noticiasagricolas.com.br/rss/noticias.xml` |
| Agro | Reuters Business | `https://feeds.reuters.com/reuters/businessNews` |
| Cibersegurança | The Hacker News | `https://feeds.feedburner.com/TheHackersNews` |
| Cibersegurança | BleepingComputer | `https://www.bleepingcomputer.com/feed/` |
| Saúde & Biotech | STAT News | `https://www.statnews.com/feed/` |
| Saúde & Biotech | Fierce Biotech | `https://www.fiercebiotech.com/rss/xml` |
| Carreira de Dados | KDnuggets | `https://www.kdnuggets.com/feed` |
| Carreira de Dados | DataHackers (PT-BR) | `https://medium.com/data-hackers/feed` |
| Carreira de Dados | Analytics Vidhya | `https://www.analyticsvidhya.com/feed/` |
| Carreira de Dados | Towards Data Science | `https://towardsdatascience.com/feed` |
| Carreira de Dados | Real Python | `https://realpython.com/atom.xml` |
| Carreira de Dados | Fast.ai | `https://www.fast.ai/atom.xml` |

## Resultados
- Script funcional e estável desde 2026-06-15; rodando automaticamente desde 2026-06-18
- E-mail HTML entregue com formatação correta (markdown convertido para HTML, listas renderizadas corretamente)
- Log CSV gerado em `Relatorios/log news script/`
- Cópias HTML em `Relatorios/digests/YYYY/`
- Weekly Intel (sexta) confirmado funcionando em 2026-06-20
- Agendamento ativo via Task Scheduler — **segunda a sexta às 9:15**
  - Launcher: `C:\Users\guilherme.saito\task\news_digest.bat` (usa `python.exe` — `pythonw.exe` causava crash silencioso com `sys.stdout=None`)
  - Configurações: `StartWhenAvailable`, sem restrição de bateria, retry automático 2× a cada 10 min, limite de 30 min
  - Popup de falha: `notify_failure.ps1` exibe alerta com botões Retry/Cancel se o script falhar
- Weekly Intel agora usa acumulador semanal (`Inputs/semana_atual.json`) com resumos de todos os dias — não apenas sexta
- Assuntos monitorados da semana anterior injetados como contexto no prompt (`Inputs/para_monitorar.md`)
- Histórico semanal arquivado em `Relatorios/semanas/semana_YYYY-MM-DD.json`
- **2026-06-22 — melhorias de qualidade (sessão 1):** MAX_ARTICLES_PER_TOPIC 10→20, truncamento 400→700, filtro < 80 chars, dedup semântica, cross-day dedup, system message separado, anti-alucinação, formato estruturado, temperature Weekly Intel 0.7→0.5
- **2026-06-22 — melhorias de qualidade (sessão 2):** rate limiter Groq 15 req/min, diversificação por round-robin, seção `🔭 Acompanhamento` (feedback para_monitorar), supressão de tópicos vazios, footer versão corrigida
- **2026-06-26 — fix Weekly Intel (Error 413):** prompt do brief semanal atingiu 15.815 tokens (limite = 12.000 TPM free tier) com 5 dias × 8 personas sem truncamento; corrigido com truncamento de 400 chars por entrada diária antes de concatenar o bloco de resumos

## Limitações
- NewsAPI (plano gratuito): 100 req/dia — 8 tópicos × 4 idiomas = 32 req/execução; limite para ~3 execuções/dia
- NewsAPI free tier tem delay de ~24h; por isso `HORAS_ATRAS = 48`
- O LLM resume o que o pipeline entrega — não faz curadoria própria
- Groq free tier: ~500k tokens/dia, 6k tokens/min; retry automático resolve picos
- RSS sem timestamp são incluídos por segurança (podem trazer artigos antigos)
- Deduplicação cross-tópico usa ordem do JSON — artigo "multi-tópico" sempre vai para o tópico de maior prioridade (posição mais alta no arquivo)
- Weekly Intel depende da qualidade dos 8 resumos; se um tópico falhar, a mesa redonda fica incompleta

## Referências
- [feedparser docs](https://feedparser.readthedocs.io/)
- [NewsAPI docs](https://newsapi.org/docs)
- [Groq API docs](https://console.groq.com/docs)
- [python-dotenv](https://pypi.org/project/python-dotenv/)
- [pywin32 / win32com](https://pypi.org/project/pywin32/)
- [markdown lib](https://python-markdown.github.io/)
