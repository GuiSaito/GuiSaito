"""
News Digest — Coleta, sumariza e envia notícias por e-mail.
Script: news_digest_v1.1.py
"""

import sys
import os

# Forçar UTF-8 no stdout (Task Scheduler pode usar cp1252 por padrão)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import re
import csv
import json
import time
import threading
import feedparser
import requests
import win32com.client
import markdown  # type: ignore
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv(Path(__file__).parent / "Inputs" / ".env")

# ---------------------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------------------

GROQ_API_KEY = os.getenv("GROQ_API_KEY")
NEWSAPI_KEY  = os.getenv("NEWSAPI_KEY")
EMAIL_TO       = os.getenv("EMAIL_TO")   # destinatário(s) — separe por ponto-e-vírgula para múltiplos

MAX_ARTICLES_PER_TOPIC = 20   # artigos enviados para o LLM por tópico
GROQ_MODEL = "llama-3.3-70b-versatile"  # DEPRECADO em 16/08/2026 — trocar para qwen/qwen3.6-27b ou similar
_dia_semana = datetime.now().weekday()  # 0=segunda, 4=sexta, 6=domingo
HORAS_ATRAS = 72 if _dia_semana == 0 else 48  # 72h às segundas (cobre fim de semana)
AGRUPAR_POR_SUBTEMA = True    # LLM agrupa artigos por sub-tema antes de resumir
USAR_NEWSAPI = True           # False durante testes para poupar cota (100 req/dia no free tier)

HTML_DIR = Path(__file__).parent / "Relatorios" / "digests"
LOG_DIR  = Path(__file__).parent / "Relatorios" / "log news script"

# Acumulador semanal e rastreamento
SEMANA_FILE         = Path(__file__).parent / "Inputs" / "semana_atual.json"
PARA_MONITORAR_FILE = Path(__file__).parent / "Inputs" / "para_monitorar.md"
URLS_SEMANA_FILE    = Path(__file__).parent / "Inputs" / "urls_semana.json"
SEMANAS_DIR         = Path(__file__).parent / "Relatorios" / "semanas"

# ---------------------------------------------------------------------------
# Tópicos: carregados de Inputs/topics.json
# ---------------------------------------------------------------------------

TOPICS: dict = json.loads(
    (Path(__file__).parent / "Inputs" / "topics.json").read_text(encoding="utf-8-sig")
)

# ---------------------------------------------------------------------------
# Log de execução
# ---------------------------------------------------------------------------

def salvar_log(digest: dict, status_email: str, erros: list[str]) -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    mes_atual = datetime.now().strftime("%Y-%m")
    log_path = LOG_DIR / f"log_{mes_atual}.csv"
    campos = ["data_hora", "topico", "artigos_coletados", "status_email", "erros"]
    escrever_header = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        if escrever_header:
            writer.writeheader()
        data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        erros_str = " | ".join(erros) if erros else ""
        for topico, dados in digest.items():
            writer.writerow({
                "data_hora": data_hora,
                "topico": topico,
                "artigos_coletados": len(dados["artigos"]),
                "status_email": status_email,
                "erros": erros_str,
            })
    print(f"[Log] Salvo em {log_path}")


# ---------------------------------------------------------------------------
# Coleta de artigos
# ---------------------------------------------------------------------------

def _dentro_janela(entry) -> bool:
    """Verifica se o artigo foi publicado dentro da janela de coleta."""
    cutoff = datetime.now(timezone.utc) - timedelta(hours=HORAS_ATRAS)
    for attr in ("published_parsed", "updated_parsed"):
        ts = getattr(entry, attr, None)
        if ts:
            pub = datetime(*ts[:6], tzinfo=timezone.utc)
            return pub >= cutoff
    return True  # sem timestamp → inclui por segurança


def coletar_rss(urls: list[str]) -> list[dict]:
    artigos = []
    for url in urls:
        try:
            feed = feedparser.parse(url)
            for entry in feed.entries:
                if not _dentro_janela(entry):
                    continue
                artigos.append({
                    "titulo": _strip_html(entry.get("title", "")),
                    "resumo": _strip_html(entry.get("summary", entry.get("description", "")))[:700],
                    "url": entry.get("link", ""),
                    "fonte": feed.feed.get("title", url),
                })
        except Exception as e:
            print(f"[RSS] Erro ao coletar {url}: {e}")
    return artigos


NEWSAPI_LANGUAGES = ["pt", "en", "es", "de"]


def coletar_newsapi(query: str) -> list[dict]:
    if not NEWSAPI_KEY:
        return []
    from_date = (datetime.now(timezone.utc) - timedelta(hours=HORAS_ATRAS)).strftime("%Y-%m-%dT%H:%M:%SZ")
    artigos = []
    for lang in NEWSAPI_LANGUAGES:
        try:
            resp = requests.get(
                "https://newsapi.org/v2/everything",
                params={
                    "q": query,
                    "from": from_date,
                    "sortBy": "publishedAt",
                    "pageSize": 20,
                    "language": lang,
                    "apiKey": NEWSAPI_KEY,
                },
                timeout=15,
            )
            resp.raise_for_status()
            for art in resp.json().get("articles", []):
                artigos.append({
                    "titulo": _strip_html(art.get("title", "")),
                    "resumo": _strip_html(art.get("description", ""))[:700],
                    "url": art.get("url", ""),
                    "fonte": art.get("source", {}).get("name", "NewsAPI"),
                })
        except Exception as e:
            print(f"[NewsAPI] Erro na query '{query}' (lang={lang}): {e}")
    return artigos


def _strip_html(texto: str) -> str:
    """Remove tags HTML de strings vindas de feeds RSS."""
    return re.sub(r"<[^>]+>", "", texto or "").strip()


def deduplicar(artigos: list[dict]) -> list[dict]:
    vistos = set()
    unicos = []
    for art in artigos:
        url = art["url"].strip().rstrip("/")
        if url and url not in vistos:
            vistos.add(url)
            unicos.append(art)
    return unicos


def _titulo_palavras(titulo: str) -> set[str]:
    """Palavras relevantes do título para comparação semântica (ignora stopwords curtas)."""
    stopwords = {"de", "do", "da", "dos", "das", "em", "no", "na", "o", "a", "e", "é",
                 "um", "uma", "para", "com", "the", "of", "in", "to", "and", "a", "is"}
    return {w.lower() for w in re.findall(r'\w+', titulo) if len(w) > 2 and w.lower() not in stopwords}


def deduplicar_semantico(artigos: list[dict], limiar: float = 0.7) -> list[dict]:
    """Remove artigos sobre o mesmo evento por similaridade de título (Jaccard >= limiar)."""
    mantidos = []
    for art in artigos:
        palavras_art = _titulo_palavras(art["titulo"] or "")
        if not palavras_art:
            mantidos.append(art)
            continue
        duplicado = False
        for mantido in mantidos:
            palavras_m = _titulo_palavras(mantido["titulo"] or "")
            if not palavras_m:
                continue
            jaccard = len(palavras_art & palavras_m) / len(palavras_art | palavras_m)
            if jaccard >= limiar:
                # Mantém o artigo com resumo mais longo (mais completo)
                if len(art.get("resumo", "")) > len(mantido.get("resumo", "")):
                    mantidos.remove(mantido)
                    mantidos.append(art)
                duplicado = True
                break
        if not duplicado:
            mantidos.append(art)
    return mantidos


def filtrar_sem_conteudo(artigos: list[dict], min_chars: int = 80) -> list[dict]:
    """Remove artigos sem resumo útil (apenas título ou summary muito curto)."""
    return [a for a in artigos if len(a.get("resumo", "")) >= min_chars]


def _match_keyword(kw: str, texto: str) -> bool:
    """Match com word boundary para evitar falsos positivos por substring."""
    pattern = r'\b' + re.escape(kw) + r'\b'
    return bool(re.search(pattern, texto, re.IGNORECASE))


def coletar_topico(config: dict) -> list[dict]:
    artigos = coletar_rss(config["rss"])
    if USAR_NEWSAPI:
        artigos += coletar_newsapi(config.get("newsapi_query", ""))
    artigos = deduplicar(artigos)
    artigos = filtrar_sem_conteudo(artigos)
    artigos = deduplicar_semantico(artigos)

    if config.get("usar_keywords", True):
        kws = config.get("keywords", [])
        if kws:
            artigos = [
                a for a in artigos
                if any(_match_keyword(kw, (a["titulo"] or "") + " " + (a["resumo"] or "")) for kw in kws)
            ]
    artigos = _diversificar_fontes(artigos)
    return artigos[:MAX_ARTICLES_PER_TOPIC]


def _diversificar_fontes(artigos: list[dict]) -> list[dict]:
    """Round-robin por fonte: intercala artigos garantindo diversidade de fontes.
    Fontes com mais artigos preenchem os slots restantes depois que as menores se esgotam.
    Remove a necessidade de um cap fixo por fonte sem desperdiçar artigos distintos.
    """
    por_fonte: dict[str, list[dict]] = {}
    for art in artigos:
        por_fonte.setdefault(art.get("fonte", ""), []).append(art)

    resultado = []
    filas = list(por_fonte.values())
    while filas:
        filas_novas = []
        for fila in filas:
            if fila:
                resultado.append(fila.pop(0))
            if fila:
                filas_novas.append(fila)
        filas = filas_novas
    return resultado

# ---------------------------------------------------------------------------
# Sumarização com LLM
# ---------------------------------------------------------------------------

client = OpenAI(
    api_key=GROQ_API_KEY,
    base_url="https://api.groq.com/openai/v1",
)


class _RateLimiter:
    """Garante um intervalo mínimo entre chamadas à API do Groq (free tier)."""
    def __init__(self, calls_per_minute: int = 15):
        self._lock = threading.Lock()
        self._interval = 60.0 / calls_per_minute
        self._last: float = 0.0

    def wait(self):
        with self._lock:
            elapsed = time.time() - self._last
            if elapsed < self._interval:
                time.sleep(self._interval - elapsed)
            self._last = time.time()


_groq_rate_limiter = _RateLimiter(calls_per_minute=15)  # conservador p/ free tier


def sumarizar(topico: str, artigos: list[dict], config: dict | None = None) -> str:
    if not artigos:
        return f"Nenhuma notícia relevante encontrada nas últimas {HORAS_ATRAS}h."

    bloco = "\n\n".join(
        f"[{i+1}] {a['titulo']}\nFonte: {a['fonte']}\n{a['resumo']}"
        for i, a in enumerate(artigos)
    )

    prompt_override = (config or {}).get("prompt_override")
    eh_sexta = datetime.now().weekday() == 4

    # --- System message: identidade, anti-alucinação, formato ---
    system_msg = (
        "Você é um analista de inteligência de mercado sênior. "
        "Você recebe artigos jornalísticos e produz briefings executivos objetivos em português (PT-BR). "
        "REGRAS OBRIGATÓRIAS:\n"
        "1. Use APENAS as informações dos artigos fornecidos. Não adicione conhecimento geral ou contexto externo.\n"
        "2. Se uma informação não estiver nos artigos, não a mencione.\n"
        "3. Evite completamente frases de enchimento como 'é importante notar', 'vale ressaltar', "
        "'é fundamental destacar', 'cabe mencionar', 'destaca-se que'.\n"
        "4. Mencione as fontes naturalmente quando relevante (ex: 'Segundo o Reuters...').\n"
        "5. Identifique histórias distintas. Para múltiplos artigos sobre o mesmo evento, "
        "use apenas o mais completo — não repita o mesmo fato em bullets diferentes.\n"
        "6. Não repita o nome do tópico no início do resumo."
    )

    if prompt_override and eh_sexta:
        user_msg = (
            f"{prompt_override}\n\n"
            f"ATENÇÃO: Hoje é sexta-feira. Além do resumo habitual, adicione ao final uma seção "
            f"**Narrativa da Semana** com 2-3 bullets sobre: qual foi o tema dominante desta semana, "
            f"o que pode impactar a próxima semana.\n\n"
            f"Formato de cada notícia:\n"
            f"**[Título curto do evento]** — 2-3 frases diretas.\n\n"
            f"Finalize com:\n"
            f"🔎 **Para ficar de olho:** [1 ponto concreto de atenção]\n\n"
            f"ARTIGOS:\n{bloco}"
        )
    elif prompt_override:
        user_msg = (
            f"{prompt_override}\n\n"
            f"Formato de cada notícia:\n"
            f"**[Título curto do evento]** — 2-3 frases diretas.\n\n"
            f"Finalize com:\n"
            f"🔎 **Para ficar de olho:** [1 ponto concreto de atenção]\n\n"
            f"ARTIGOS:\n{bloco}"
        )
    elif eh_sexta:
        user_msg = (
            f"Abaixo estão artigos da semana sobre '{topico}'.\n"
            f"Faça um resumo semanal com duas partes:\n\n"
            f"**Destaques da Semana** — até 4 bullets, cada um no formato:\n"
            f"**[Título curto]** — 2-3 frases diretas.\n\n"
            f"**Narrativa da Semana** — 2 bullets sobre o tema dominante e o que pode impactar a próxima semana.\n\n"
            f"🔎 **Para ficar de olho:** [1 ponto concreto]\n\n"
            f"ARTIGOS:\n{bloco}"
        )
    else:
        usar_agrupamento = AGRUPAR_POR_SUBTEMA and len(artigos) >= 3
        if usar_agrupamento:
            user_msg = (
                f"Abaixo estão artigos recentes sobre '{topico}'.\n"
                f"Identifique as histórias distintas e agrupe por sub-tema (máximo 3 grupos). "
                f"Para cada grupo, escreva os destaques no formato:\n\n"
                f"**[Nome do Sub-tema]**\n"
                f"**[Título curto do evento]** — 2-3 frases diretas.\n"
                f"**[Próximo evento distinto]** — ...\n\n"
                f"Finalize com:\n"
                f"🔎 **Para ficar de olho:** [1 ponto concreto de atenção]\n\n"
                f"ARTIGOS:\n{bloco}"
            )
        else:
            user_msg = (
                f"Abaixo estão artigos recentes sobre '{topico}'.\n"
                f"Resuma os destaques no formato:\n\n"
                f"**[Título curto do evento]** — 2-3 frases diretas.\n"
                f"**[Próximo evento distinto]** — ...\n\n"
                f"Máximo 5 eventos. Finalize com:\n"
                f"🔎 **Para ficar de olho:** [1 ponto concreto de atenção]\n\n"
                f"ARTIGOS:\n{bloco}"
            )
    try:
        _groq_rate_limiter.wait()
        resp = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ],
            max_tokens=700,
            temperature=0.3,
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        # Retry com backoff exponencial (3 tentativas)
        for tentativa, espera in enumerate([2, 4, 8], start=1):
            try:
                print(f"[LLM] Retry {tentativa}/3 para '{topico}' após {espera}s...")
                time.sleep(espera)
                resp = client.chat.completions.create(
                    model=GROQ_MODEL,
                    messages=[
                        {"role": "system", "content": system_msg},
                        {"role": "user", "content": user_msg},
                    ],
                    max_tokens=700,
                    temperature=0.3,
                )
                return resp.choices[0].message.content.strip()
            except Exception:
                continue
        print(f"[LLM] Falhou após 3 tentativas para '{topico}': {e}")
        return "Erro ao gerar resumo."

# ---------------------------------------------------------------------------
# Montagem do e-mail HTML
# ---------------------------------------------------------------------------

def verificar_monitoramento(digest: dict) -> list[dict]:
    """
    Verifica quais itens de 'Para monitorar' têm artigos correspondentes nesta execução.
    Retorna lista de {item, artigos} para exibição no email.
    """
    if not PARA_MONITORAR_FILE.exists():
        return []
    try:
        conteudo = PARA_MONITORAR_FILE.read_text(encoding="utf-8")
    except Exception:
        return []

    bullets = [
        re.sub(r'^\s*[-•*]\s+', '', linha).strip()
        for linha in conteudo.splitlines()
        if re.match(r'^\s*[-•*]\s+', linha) and len(linha.strip()) > 5
    ]
    if not bullets:
        return []

    todos_artigos = [
        (topico, art)
        for topico, dados in digest.items()
        for art in dados.get("artigos", [])
    ]

    resultados = []
    for bullet in bullets:
        palavras_bullet = _titulo_palavras(bullet)
        if len(palavras_bullet) < 2:
            continue
        correspondentes = []
        for topico, art in todos_artigos:
            texto_art = (art.get("titulo", "") + " " + art.get("resumo", "")).lower()
            hits = sum(1 for p in palavras_bullet if p in texto_art)
            if hits >= 2 and hits / len(palavras_bullet) >= 0.3:
                correspondentes.append({**art, "topico": topico})
        if correspondentes:
            resultados.append({"item": bullet, "artigos": correspondentes[:3]})
    return resultados


def _secao_monitoramento_html(matches: list[dict]) -> str:
    if not matches:
        return ""
    itens_html = ""
    for m in matches:
        links = "".join(
            f'<li><a href="{a["url"]}" style="color:#555;">{a["titulo"]}</a> '
            f'<span style="color:#888;font-size:11px;">({a["fonte"]})</span></li>'
            for a in m["artigos"]
        )
        itens_html += (
            f'<div style="margin-bottom:12px;">'
            f'<p style="margin:0 0 4px 0;font-weight:bold;color:#7d5a00;">📌 {m["item"]}</p>'
            f'<ul style="margin:2px 0;padding-left:18px;font-size:13px;">{links}</ul>'
            f'</div>'
        )
    return (
        f'<div style="background:#fff8e1;border-left:4px solid #f39c12;padding:14px 16px;'
        f'border-radius:4px;margin-bottom:24px;">'
        f'<h2 style="color:#e67e22;margin:0 0 4px 0;font-size:16px;">🔭 Acompanhamento</h2>'
        f'<p style="font-size:12px;color:#999;margin:0 0 12px 0;">Itens em monitoramento com novas coberturas</p>'
        f'{itens_html}</div>'
    )


def _indice_html(topicos: list[str]) -> str:
    links = " &nbsp;|&nbsp; ".join(
        f'<a href="#{t.lower().replace(" ", "-").replace("&", "").replace(",", "")}" style="color:#2980b9;text-decoration:none;">{t}</a>'
        for t in topicos
    )
    return f'<p style="font-size:13px;background:#eaf4fb;padding:10px;border-radius:4px;">📌 {links}</p>'


def _secao_html(topico: str, resumo: str, artigos: list[dict]) -> str:
    links_html = "".join(
        f'<li><a href="{a["url"]}">{a["titulo"]}</a> <span style="color:#888;font-size:12px;">({a["fonte"]})</span></li>'
        for a in artigos
    )
    resumo_html = markdown.markdown(re.sub(r'([^\n])\n(\s*[\*\-] )', r'\1\n\n\2', resumo))
    anchor = topico.lower().replace(" ", "-").replace("&", "").replace(",", "")
    nome, cargo = PERSONAS.get(topico, (None, None))
    persona_html = (
        f'<p style="margin:2px 0 10px 0;font-size:12px;color:#666;">📝 por <b>{nome}</b> · {cargo}</p>'
        if nome else ""
    )
    return f"""
    <div id="{anchor}" style="margin-bottom:32px;">
        <h2 style="color:#1a3c5e;border-bottom:2px solid #d4e6f1;padding-bottom:6px;margin-bottom:4px;">{topico}</h2>
        {persona_html}
        <div style="background:#f0f7ff;padding:12px;border-left:4px solid #2980b9;border-radius:4px;">
            {resumo_html}
        </div>
        <ul style="font-size:13px;color:#333;">{links_html}</ul>
    </div>
    """


def montar_email_html(digest: dict[str, dict], monitoramento: list[dict] | None = None) -> str:
    data_str = datetime.now().strftime("%d/%m/%Y")
    eh_sexta = datetime.now().weekday() == 4
    tipo = "Digest Semanal" if eh_sexta else "News Digest"
    subtitulo = f"Resumo semanal — semana de {data_str}" if eh_sexta else f"Resumo automático das últimas {HORAS_ATRAS}h"
    # D: filtrar tópicos sem conteúdo (0 artigos ou resumo None)
    digest_valido = {
        t: d for t, d in digest.items()
        if d.get("resumo") is not None and d.get("artigos")
    }
    indice = _indice_html(list(digest_valido.keys()))
    secoes = "".join(
        _secao_html(topico, dados["resumo"], dados["artigos"])
        for topico, dados in digest_valido.items()
    )
    secao_monitor = _secao_monitoramento_html(monitoramento or [])
    return f"""
    <html><body style="font-family:Segoe UI,Arial,sans-serif;max-width:700px;margin:auto;color:#222;">
        <h1 style="color:#1a3c5e;">📰 {tipo} — {data_str}</h1>
        <p style="color:#666;font-size:13px;">{subtitulo}</p>
        {indice}
        <hr>
        {secao_monitor}
        {secoes}
        <hr>
        <p style="font-size:11px;color:#aaa;">Gerado automaticamente por news_digest_v1.1.py</p>
    </body></html>
    """

# ---------------------------------------------------------------------------
# Brief Semanal (apenas sextas) — e-mail separado com mesa redonda
# ---------------------------------------------------------------------------

PERSONAS = {
    "Geopolítica & Economia Global": ("Maquiavel", "analista geopolítico"),
    "Finanças":                     ("Midas",     "analista de mercados financeiros"),
    "Tech, Consultoria & Mercado": ("Edison",    "consultor de estratégia e tecnologia"),
    "Inteligência Artificial":     ("Alan",      "pesquisador de IA e sistemas inteligentes"),
    "Agro":                         ("Deméter",   "analista sênior de agronegócio"),
    "Cibersegurança":              ("Loki",      "especialista em cibersegurança ofensiva e defensiva"),
    "Saúde & Biotech":             ("Asclépio",  "analista de healthcare e biotecnologia"),
    "Carreira de Dados":            ("Florence",  "mentora de carreira em dados e analytics"),
}


def salvar_resumos_semana(digest: dict):
    """Acumula os resumos do dia em semana_atual.json, organizado por persona."""
    data_hoje = datetime.now().strftime("%Y-%m-%d")
    acum: dict = {}
    if SEMANA_FILE.exists():
        try:
            acum = json.loads(SEMANA_FILE.read_text(encoding="utf-8"))
        except Exception:
            acum = {}
    for topico, dados in digest.items():
        nome, _ = PERSONAS.get(topico, (topico, ""))
        if nome not in acum:
            acum[nome] = {}
        acum[nome][data_hoje] = dados["resumo"]
    SEMANA_FILE.write_text(json.dumps(acum, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[Semana] Resumos de {data_hoje} salvos em {SEMANA_FILE.name}")


def arquivar_semana(texto_brief: str):
    """Arquiva semana_atual.json, extrai 'Para monitorar' e reseta o acumulador."""
    SEMANAS_DIR.mkdir(parents=True, exist_ok=True)
    data_sexta = datetime.now().strftime("%Y-%m-%d")
    if SEMANA_FILE.exists():
        destino = SEMANAS_DIR / f"semana_{data_sexta}.json"
        destino.write_text(SEMANA_FILE.read_text(encoding="utf-8"), encoding="utf-8")
        SEMANA_FILE.write_text("{}", encoding="utf-8")
        print(f"[Semana] Acumulador arquivado em {destino.name} e resetado")
    # Resetar histórico de URLs da semana
    URLS_SEMANA_FILE.write_text("[]", encoding="utf-8")
    print(f"[Semana] urls_semana.json resetado")
    # Extrair seção "Para monitorar" e salvar para a próxima semana
    match = re.search(r"\*\*Para monitorar:\*\*(.*?)(?:\n---|\n##|\Z)", texto_brief, re.DOTALL)
    if match:
        pontos = match.group(1).strip()
        data_str = datetime.now().strftime("%d/%m/%Y")
        PARA_MONITORAR_FILE.write_text(
            f"# Para monitorar — semana de {data_str}\n\n{pontos}\n",
            encoding="utf-8"
        )
        print(f"[Semana] 'Para monitorar' salvo em {PARA_MONITORAR_FILE.name}")


def gerar_brief_semanal(digest: dict) -> dict:
    """Chama o LLM para gerar o Quick Brief e a Mesa Redonda a partir dos resumos da semana completa."""
    _DIAS_PT = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

    # Tentar usar acumulador da semana inteira; fallback para digest do dia
    acum: dict = {}
    if SEMANA_FILE.exists():
        try:
            acum = json.loads(SEMANA_FILE.read_text(encoding="utf-8"))
        except Exception:
            acum = {}

    blocos = []
    if acum:
        for topico, (nome, cargo) in PERSONAS.items():
            if nome in acum and acum[nome]:
                entradas = sorted(acum[nome].items())  # ordem cronológica
                dias_str = "\n".join(
                    f"{_DIAS_PT[datetime.strptime(d, '%Y-%m-%d').weekday()]} ({datetime.strptime(d, '%Y-%m-%d').strftime('%d/%m')}): {r[:400]}"
                    for d, r in entradas
                )
                blocos.append(f"[{nome} — {cargo}]\n{dias_str}")
    else:
        for topico, dados in digest.items():
            nome, cargo = PERSONAS.get(topico, (topico, "especialista"))
            blocos.append(f"[{nome} — {cargo}]\n{dados['resumo']}")

    resumos_concat = "\n\n".join(blocos)

    # Contexto de monitoramento da semana anterior
    contexto_str = ""
    if PARA_MONITORAR_FILE.exists():
        try:
            pontos_anteriores = PARA_MONITORAR_FILE.read_text(encoding="utf-8").strip()
            if pontos_anteriores:
                contexto_str = f"\n## Assuntos em monitoramento da semana passada\n{pontos_anteriores}\n\n---\n"
        except Exception:
            pass

    prompt = f"""Você é um editor-chefe que recebeu os briefings da semana completa de 8 especialistas.
{contexto_str}
Abaixo estão os resumos de cada especialista organizados por dia da semana:

{resumos_concat}

---

REGRAS OBRIGATÓRIAS:
1. Use APENAS as informações dos briefings acima. Não adicione contexto externo ou conhecimento geral.
2. Os "Sinais fracos" devem ser baseados em artigos ou tendências específicos mencionados nos briefings desta semana — não em especulações gerais.
3. Os itens de "Para monitorar" devem referenciar situações concretas dos briefings.
4. Evite completamente frases de enchimento como "é importante notar", "vale ressaltar", "é fundamental".
5. Na Mesa Redonda, não force conexões entre áreas que não aparecem nos briefings desta semana.

Gere um email semanal em PT-BR com DUAS partes exatamente neste formato:

## ⚡ QUICK BRIEF
**Tema da semana:** [uma frase que capture o fio central desta semana]
**Sinais fracos:** [2 bullets de tendências emergentes baseadas nos briefings desta semana]
**Para monitorar:** [3 itens concretos para a próxima semana, referenciando situações dos briefings]

---

## 🎭 MESA REDONDA
[Simule um diálogo casual e envolvente entre os especialistas. Cada um fala pelo menos uma vez, trazendo o highlight da sua semana. Deixe conexões entre áreas emergirem naturalmente na conversa — não force ligações que não existem nesta semana. O tom é inteligente mas descontraído, como colegas discutindo após uma semana intensa. Use o nome de cada especialista ao falar. Máximo 500 palavras.]

**A pergunta da semana:** [Uma única frase provocativa e aberta para reflexão no fim de semana]"""

    try:
        _groq_rate_limiter.wait()
        resp = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1200,
            temperature=0.5,
        )
        texto = resp.choices[0].message.content.strip()
        tema = ""
        for linha in texto.splitlines():
            if linha.startswith("**Tema da semana:"):
                tema = linha.replace("**Tema da semana:", "").replace("**", "").strip()
                break
        return {"texto": texto, "tema": tema}
    except Exception as e:
        print(f"[Brief] Erro ao gerar brief semanal: {e}")
        return {"texto": "Erro ao gerar brief semanal.", "tema": ""}


def montar_email_brief(texto: str, tema: str) -> str:
    data_str = datetime.now().strftime("%d/%m/%Y")
    personas_html = "".join(
        f'<span style="display:inline-block;margin:3px;padding:4px 10px;background:#e8f4fd;border-radius:12px;font-size:12px;">'
        f'<b>{nome}</b> <span style="color:#888;">({cargo})</span></span>'
        for nome, cargo in PERSONAS.values()
    )
    conteudo_html = markdown.markdown(
        re.sub(r'([^\n])\n(\s*[\*\-] )', r'\1\n\n\2',           # linha em branco antes de bullets
        re.sub(r'[ \t]*\*\*(Para monitorar|Sinais fracos|Tema da semana|A pergunta da semana):\*\*',
               r'\n\n**\1:**', texto))                            # seções do brief em linha própria
    )
    return f"""
    <html><body style="font-family:Segoe UI,Arial,sans-serif;max-width:700px;margin:auto;color:#222;">
        <h1 style="color:#1a3c5e;">🧠 Weekly Intel — {data_str}</h1>
        <p style="color:#666;font-size:13px;font-style:italic;">{tema}</p>
        <p style="font-size:12px;color:#999;">Participantes desta semana: {personas_html}</p>
        <hr>
        <div style="line-height:1.7;">{conteudo_html}</div>
        <hr>
        <p style="font-size:11px;color:#aaa;">Gerado automaticamente por news_digest_v1.1.py</p>
    </body></html>
    """


def enviar_email_brief(html: str, tema: str) -> None:
    data_str = datetime.now().strftime("%d/%m/%Y")
    subject = f"Weekly Intel — {tema}" if tema else f"Weekly Intel — {data_str}"
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.HTMLBody = html
    mail.To = EMAIL_TO
    mail.Send()
    print(f"[E-mail] Weekly Intel enviado: '{subject}'")


# ---------------------------------------------------------------------------
# Envio de e-mail via Outlook desktop (win32com — sem senha, usa sessão aberta)
# ---------------------------------------------------------------------------

def enviar_email(html: str) -> None:
    data_str = datetime.now().strftime("%d/%m/%Y")
    eh_sexta = datetime.now().weekday() == 4
    tipo = "Digest Semanal" if eh_sexta else "News Digest"
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = f"{tipo} — {data_str}"
    mail.HTMLBody = html
    mail.To = EMAIL_TO
    mail.Send()
    print(f"[E-mail] {tipo} enviado para {EMAIL_TO}")

# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------

def main():
    print(f"[Início] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Validação de variáveis de ambiente
    ausentes = [v for v, val in [("GROQ_API_KEY", GROQ_API_KEY), ("EMAIL_TO", EMAIL_TO)] if not val]
    if USAR_NEWSAPI and not NEWSAPI_KEY:
        ausentes.append("NEWSAPI_KEY")
    if ausentes:
        raise EnvironmentError(f"Variáveis ausentes no .env: {', '.join(ausentes)}")

    digest = {}
    erros: list[str] = []
    status_email = "ok"

    print("[Coletando] Todos os tópicos em paralelo...")
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {
            executor.submit(coletar_topico, config): topico
            for topico, config in TOPICS.items()
        }
        coletas = {}
        for future in as_completed(futures):
            topico = futures[future]
            try:
                artigos = future.result()
            except Exception as e:
                artigos = []
                erros.append(f"Coleta falhou: {topico}: {e}")
            print(f"  [{topico}] → {len(artigos)} artigos")
            coletas[topico] = artigos

    # Deduplicação cross-tópico: remove URLs já vistas em tópicos anteriores
    urls_vistas: set[str] = set()
    for topico in TOPICS:
        artigos_unicos = []
        for art in coletas.get(topico, []):
            url = art["url"].strip().rstrip("/")
            if url not in urls_vistas:
                urls_vistas.add(url)
                artigos_unicos.append(art)
        coletas[topico] = artigos_unicos
        if len(artigos_unicos) < len(coletas.get(topico, [])):
            print(f"  [{topico}] dedup cross-tópico: {len(coletas.get(topico, []))} → {len(artigos_unicos)}")

    # Deduplicação cross-day: remove URLs já cobertas em dias anteriores desta semana
    urls_semana: set[str] = set()
    if URLS_SEMANA_FILE.exists():
        try:
            urls_semana = set(json.loads(URLS_SEMANA_FILE.read_text(encoding="utf-8")))
        except Exception:
            urls_semana = set()
    for topico in TOPICS:
        antes = len(coletas.get(topico, []))
        coletas[topico] = [
            a for a in coletas.get(topico, [])
            if a["url"].strip().rstrip("/") not in urls_semana
        ]
        depois = len(coletas[topico])
        if depois < antes:
            print(f"  [{topico}] dedup cross-day: {antes} → {depois} artigos")
    # Registrar URLs desta execução no histórico semanal
    todas_urls = {a["url"].strip().rstrip("/") for arts in coletas.values() for a in arts}
    urls_semana |= todas_urls
    URLS_SEMANA_FILE.write_text(json.dumps(list(urls_semana), ensure_ascii=False), encoding="utf-8")

    print("[Sumarizando] Todos os tópicos em paralelo...")
    # D: pular chamada LLM para tópicos sem artigos (economiza cota e evita seções vazias)
    topicos_com_artigos = [t for t in TOPICS if coletas.get(t)]
    topicos_sem_artigos = [t for t in TOPICS if not coletas.get(t)]
    if topicos_sem_artigos:
        print(f"  [Skip LLM] Sem artigos: {', '.join(topicos_sem_artigos)}")
    with ThreadPoolExecutor(max_workers=4) as executor:  # 4 para respeitar limite de 6k tokens/min do Groq free tier
        futures_llm = {
            executor.submit(sumarizar, topico, coletas.get(topico, []), TOPICS.get(topico)): topico
            for topico in topicos_com_artigos
        }
        resumos = {}
        for future in as_completed(futures_llm):
            topico = futures_llm[future]
            try:
                resumos[topico] = future.result()
            except Exception as e:
                resumos[topico] = "Erro ao gerar resumo."
                erros.append(f"LLM falhou: {topico}: {e}")
            print(f"  [{topico}] ✓")

    for topico in TOPICS:  # mantém ordem original dos tópicos
        artigos = coletas.get(topico, [])
        resumo = resumos.get(topico)  # None para tópicos pulados (sem artigos)
        digest[topico] = {"artigos": artigos, "resumo": resumo}

    monitoramento = verificar_monitoramento(digest)
    if monitoramento:
        print(f"[Monitor] {len(monitoramento)} item(ns) em acompanhamento com cobertura hoje")
    html = montar_email_html(digest, monitoramento)

    # Salvar cópia HTML local
    HTML_DIR.mkdir(parents=True, exist_ok=True)
    ano = datetime.now().strftime("%Y")
    ano_dir = HTML_DIR / ano
    ano_dir.mkdir(parents=True, exist_ok=True)
    html_path = ano_dir / f"digest_{datetime.now().strftime('%Y-%m-%d')}.html"
    html_path.write_text(html, encoding="utf-8")
    print(f"[HTML] Cópia salva em {html_path}")

    try:
        enviar_email(html)
    except Exception as e:
        status_email = f"erro: {e}"
        erros.append(f"Email: {e}")
        print(f"[Erro] Envio de e-mail falhou: {e}")

    salvar_log(digest, status_email, erros)

    # Acumular resumos da semana (todo dia)
    salvar_resumos_semana(digest)

    # Brief semanal — apenas sextas
    if datetime.now().weekday() == 4:
        print("[Brief] Gerando Weekly Intel...")
        brief = gerar_brief_semanal(digest)
        html_brief = montar_email_brief(brief["texto"], brief["tema"])
        try:
            enviar_email_brief(html_brief, brief["tema"])
        except Exception as e:
            print(f"[Erro] Envio do Weekly Intel falhou: {e}")
        arquivar_semana(brief["texto"])

    print("[Concluído]")


if __name__ == "__main__":
    main()
