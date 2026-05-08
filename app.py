import sqlite3
import uuid
import io
import csv
import json
import re
import threading
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from flask import Flask, request, jsonify, send_file, send_from_directory
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

app = Flask(__name__, static_folder='static', static_url_path='')

DB_PATH = Path(__file__).parent / 'audits.db'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
    'Connection': 'keep-alive',
}

# ── Portuguese stopwords ───────────────────────────────────────────────────────
PT_STOPWORDS = {
    'de','a','o','que','e','do','da','em','um','para','com','uma','os','no',
    'se','na','por','mais','as','dos','como','mas','ao','ele','das','seu',
    'sua','ou','quando','muito','nos','já','eu','também','só','pelo','pela',
    'até','isso','ela','entre','depois','sem','mesmo','aos','seus','quem',
    'nas','me','esse','eles','você','essa','num','nem','suas','meu','às',
    'minha','numa','pelos','elas','havia','seja','qual','será','nós','tenho',
    'lhe','deles','essas','esses','pelas','este','fosse','dele','são','foi',
    'ser','está','tem','não','há','ter','pode','fazer','foram','isso','sobre',
    'todas','todos','qual','assim','ainda','cada','nesse','nessa','esta',
    'neste','nesta','isto','aqui','então','bem','nossa','nosso','nossas',
    'nossos','onde','porque','pois','porém','contudo','entanto','além',
    'dentro','fora','antes','sempre','nunca','vez','vezes','tanto','tanta',
    'tantos','tantas','qualquer','alguns','algumas','algum','alguma',
    'nenhum','nenhuma','através','partir','maior','menor','melhor','pior',
    'primeiro','última','último','segundo','terceiro','quatro','cinco',
    'empresa','todas','todos','novo','novos','nova','novas','grande','grandes',
    'pequeno','pequenos','outro','outros','outra','outras','todo','toda',
    'sendo','tendo','ficando','tornando','sendo','tendo','fazendo','usando',
    'esse','essa','esses','essas','aquele','aquela','aqueles','aquelas',
    'isso','aquilo','isto','cujo','cuja','cujos','cujas','qual','quais',
    'quem','quando','quanto','quanta','quantos','quantas','tudo','nada',
}


def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS audits (
        id TEXT PRIMARY KEY,
        created_at TEXT,
        name TEXT,
        status TEXT,
        total_urls INTEGER DEFAULT 0,
        completed_urls INTEGER DEFAULT 0
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS audit_results (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        audit_id TEXT,
        url TEXT,
        title TEXT,
        title_len INTEGER,
        title_status TEXT,
        meta_desc TEXT,
        meta_desc_len INTEGER,
        meta_status TEXT,
        h1_count INTEGER,
        h1_status TEXT,
        h2h3_count INTEGER,
        h2h3_status TEXT,
        error_count INTEGER,
        warn_count INTEGER,
        crawl_error TEXT,
        FOREIGN KEY (audit_id) REFERENCES audits(id)
    )''')
    for col, defn in [
        ('url_len',          'INTEGER DEFAULT 0'),
        ('url_status',       "TEXT DEFAULT 'ok'"),
        ('content_analysis', "TEXT DEFAULT '{}'"),
    ]:
        try:
            c.execute(f'ALTER TABLE audit_results ADD COLUMN {col} {defn}')
        except sqlite3.OperationalError:
            pass
    conn.commit()
    conn.close()


# ── Status helpers ─────────────────────────────────────────────────────────────

def get_url_status(url):
    l = len(url)
    if l <= 70:  return 'ok'
    if l <= 80:  return 'warn'
    return 'err'


def get_title_status(title):
    if not title: return 'absent'
    l = len(title)
    if l <= 70: return 'ok'
    if l <= 90: return 'warn'
    return 'long'


def get_meta_status(meta):
    if not meta: return 'absent'
    l = len(meta)
    if 140 <= l <= 160: return 'ok'
    return 'short' if l < 140 else 'long'


def get_h1_status(count):
    if count == 0: return 'absent'
    return 'ok' if count == 1 else 'multiple'


def get_h2h3_status(count):
    return 'ok' if count > 0 else 'absent'


def count_visible_h1s(soup):
    seen_texts = set()
    count = 0
    for h in soup.find_all('h1'):
        if h.find_parent('template'):
            continue
        if h.get('aria-hidden') == 'true':
            continue
        hidden = False
        for parent in h.parents:
            if not hasattr(parent, 'get'):
                continue
            if parent.get('aria-hidden') == 'true':
                hidden = True; break
            style = parent.get('style', '').replace(' ', '').lower()
            if 'display:none' in style or 'visibility:hidden' in style:
                hidden = True; break
        if hidden:
            continue
        own_style = h.get('style', '').replace(' ', '').lower()
        if 'display:none' in own_style or 'visibility:hidden' in own_style:
            continue
        text = ' '.join(h.get_text().split()).lower()
        if text in seen_texts:
            continue
        seen_texts.add(text)
        count += 1
    return count


def _trunc(text, n):
    return text[:n] if len(text) > n else text


# ── Keyword extraction ─────────────────────────────────────────────────────────

def extract_keywords(soup, n=5):
    """Extract top N meaningful keywords from visible body content (not meta tags).

    Strategy:
    - Read only semantic content tags: p, h1-h5, li, blockquote, td
    - Boost words that appear in headings (h1-h3) with 3× weight — they signal topic focus
    - Ignore nav, footer, header, script, style (not targeted but excluded by tag selection)
    - Fallback to full page text if no content tags are found
    """
    WORD_RE = r'\b[a-záàâãéèêíïóôõöúüçñ]{4,}\b'

    # Heading words signal topic — use for boosting
    heading_tags  = soup.find_all(['h1', 'h2', 'h3'])
    heading_text  = ' '.join(t.get_text(' ', strip=True) for t in heading_tags).lower()
    heading_words = set(re.findall(WORD_RE, heading_text))

    # Collect text from content-rich tags only
    content_tags = soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'li', 'blockquote', 'td'])
    if content_tags:
        text = ' '.join(t.get_text(' ', strip=True) for t in content_tags).lower()
    else:
        # Fallback: full visible text
        text = soup.get_text(' ', strip=True).lower()

    words    = re.findall(WORD_RE, text)
    filtered = [w for w in words if w not in PT_STOPWORDS]
    counts   = Counter(filtered)

    # Boost heading words 3× — they carry the most topical signal
    for w in heading_words:
        if w in counts:
            counts[w] *= 3

    keywords = []
    for word, _ in counts.most_common(n * 4):
        if word not in PT_STOPWORDS and len(word) >= 4:
            keywords.append(word.capitalize())
            if len(keywords) >= n:
                break
    return keywords


# ── Pauta generation ───────────────────────────────────────────────────────────

def generate_pautas(title, keywords, soup):
    """Generate 3 Vertigo-positioned pauta suggestions based on page content."""
    kw1 = keywords[0].lower() if len(keywords) > 0 else 'tecnologia digital'
    kw2 = keywords[1].lower() if len(keywords) > 1 else 'transformação digital'
    kw3 = keywords[2].lower() if len(keywords) > 2 else 'inteligência de negócios'

    return [
        {
            'titulo': _trunc(f'Como {kw1.capitalize()} redefine a estratégia de TI em empresas de grande porte', 70),
            'angulo': (f'Explorar como {kw1} e {kw2} estão redefinindo o papel do CIO/CTO — com benchmarks de mercado, '
                       f'critérios objetivos de priorização e métricas de ROI para justificar investimento ao board.'),
            'intencao': (f'Profissional de TI buscando argumentos concretos para priorizar {kw1} '
                         f'no planejamento estratégico do próximo ciclo orçamentário.'),
            'por_que_vertigo': (f'Conecta diretamente ao posicionamento "Digital Intelligence For Business": '
                                f'{kw1} como inteligência aplicada a resultados, não tecnologia pela tecnologia. '
                                f'Demonstra autoridade consultiva da Vertigo.')
        },
        {
            'titulo': _trunc(f'Os riscos de postergar {kw2} em setores regulados', 70),
            'angulo': (f'Mapeamento dos riscos operacionais, regulatórios e competitivos de adiar iniciativas '
                       f'de {kw2} — com framework de avaliação de impacto por setor e porte de empresa.'),
            'intencao': (f'CIOs e diretores que enfrentam resistência interna ou orçamento limitado e precisam de '
                         f'argumentos concretos sobre o custo real de não agir.'),
            'por_que_vertigo': (f'Narrativa central da Vertigo: a vertigem da aceleração se domina, não se ignora. '
                                f'Postergar {kw2} é escolher a vertigem sem controle — posiciona a Vertigo como '
                                f'parceira que guia a execução com segurança.')
        },
        {
            'titulo': _trunc(f'Governança de {kw3}: framework prático para CIOs e CTOs', 70),
            'angulo': (f'Guia com framework de governança para implementar {kw3} com clareza de papéis, '
                       f'KPIs definidos e alinhamento a objetivos estratégicos — testado em empresas com 500+ colaboradores.'),
            'intencao': (f'Executivos de TI que já decidiram avançar com {kw3} mas precisam de um modelo de '
                         f'execução que minimize riscos e maximize ROI de forma mensurável.'),
            'por_que_vertigo': (f'Demonstra a capacidade de execução consultiva da Vertigo: não apenas visão '
                                f'estratégica, mas método e entrega. Diferencia de concorrentes que só "vendem sonhos" '
                                f'sem estrutura de governança.')
        }
    ]


# ── Deep content analysis ──────────────────────────────────────────────────────

def analyze_content_deep(soup, title, keywords):
    """Full content analysis: structure, SEO, GEO, CIO/CTO audience, CTA."""
    full_text = soup.get_text(' ', strip=True)
    words     = full_text.split()
    word_count = len(words)
    paragraphs = [p.get_text(strip=True) for p in soup.find_all('p') if len(p.get_text(strip=True)) > 30]
    headings   = soup.find_all(['h2', 'h3', 'h4'])
    lists_count = len(soup.find_all(['ul', 'ol']))

    criteria  = []
    issues    = 0
    strengths = []

    # ── Estrutura e Clareza ────────────────────────────────────
    if paragraphs and len(paragraphs[0].split()) >= 15:
        criteria.append({'cat': 'Estrutura', 'item': 'Ideia principal no início', 'status': 'ok',
                         'text': 'Introdução com conteúdo substancial nos primeiros parágrafos — bom para leitores e IAs.'})
        strengths.append('Introdução direta e clara')
    else:
        criteria.append({'cat': 'Estrutura', 'item': 'Ideia principal no início', 'status': 'warn',
                         'text': 'Primeiros parágrafos curtos ou ausentes. Inicie com a ideia principal de forma direta e objetiva.'})
        issues += 1

    h_count = len(headings)
    if h_count >= 3:
        criteria.append({'cat': 'Estrutura', 'item': 'Hierarquia de conteúdo', 'status': 'ok',
                         'text': f'{h_count} subtítulos detectados — boa organização por seções facilita leitura e escaneamento.'})
        strengths.append(f'Estrutura hierárquica com {h_count} subtítulos')
    elif h_count >= 1:
        criteria.append({'cat': 'Estrutura', 'item': 'Hierarquia de conteúdo', 'status': 'warn',
                         'text': f'Apenas {h_count} subtítulo(s). Adicione mais H2/H3 para dividir o conteúdo em seções temáticas.'})
        issues += 1
    else:
        criteria.append({'cat': 'Estrutura', 'item': 'Hierarquia de conteúdo', 'status': 'err',
                         'text': 'Nenhum subtítulo H2/H3. Sem hierarquia, leitores e IAs não identificam as seções do conteúdo.'})
        issues += 2

    if word_count >= 800:
        criteria.append({'cat': 'Estrutura', 'item': 'Profundidade', 'status': 'ok',
                         'text': f'{word_count} palavras — conteúdo com profundidade adequada para autoridade e ranqueamento.'})
        strengths.append(f'Conteúdo extenso ({word_count} palavras)')
    elif word_count >= 400:
        criteria.append({'cat': 'Estrutura', 'item': 'Profundidade', 'status': 'warn',
                         'text': f'{word_count} palavras — conteúdo moderado. Considere expandir para 800+ palavras em temas competitivos.'})
        issues += 1
    else:
        criteria.append({'cat': 'Estrutura', 'item': 'Profundidade', 'status': 'err',
                         'text': f'Apenas {word_count} palavras — conteúdo raso. IAs e buscadores priorizam profundidade e completude.'})
        issues += 2

    # ── Qualidade SEO ──────────────────────────────────────────
    qwords = ['como', 'por que', 'quando', 'onde', 'o que', 'quais', 'qual', 'quanto']
    if any(q in full_text.lower()[:500] for q in qwords):
        criteria.append({'cat': 'SEO', 'item': 'Intenção de busca', 'status': 'ok',
                         'text': 'Conteúdo responde a uma intenção informacional clara nos primeiros parágrafos.'})
    else:
        criteria.append({'cat': 'SEO', 'item': 'Intenção de busca', 'status': 'warn',
                         'text': 'Intenção de busca não explícita. Certifique-se de que o conteúdo responde a uma pergunta ou necessidade específica.'})
        issues += 1

    matches = re.findall(r'\b\d+[\.,]?\d*\s*%|\b\d+\s*(?:mil|bilh|trilh)|\b(?:19|20)\d{2}\b|\b\d+x\b',
                          full_text, re.IGNORECASE)
    if len(matches) >= 3:
        criteria.append({'cat': 'SEO', 'item': 'Dados e evidências', 'status': 'ok',
                         'text': f'{len(matches)} referências numéricas detectadas — fortalece autoridade e credibilidade.'})
        strengths.append('Uso consistente de dados quantitativos')
    elif len(matches) >= 1:
        criteria.append({'cat': 'SEO', 'item': 'Dados e evidências', 'status': 'warn',
                         'text': 'Poucos dados concretos. Adicione estatísticas, percentuais e datas para aumentar credibilidade.'})
        issues += 1
    else:
        criteria.append({'cat': 'SEO', 'item': 'Dados e evidências', 'status': 'err',
                         'text': 'Sem dados quantitativos. IAs e buscadores priorizam conteúdo com evidências concretas e verificáveis.'})
        issues += 2

    # ── Qualidade GEO ──────────────────────────────────────────
    if lists_count >= 2:
        criteria.append({'cat': 'GEO', 'item': 'Listas estruturadas', 'status': 'ok',
                         'text': f'{lists_count} listas detectadas — alta probabilidade de extração como snippet por IAs generativas.'})
        strengths.append('Listas estruturadas para extração por IAs')
    elif lists_count == 1:
        criteria.append({'cat': 'GEO', 'item': 'Listas estruturadas', 'status': 'warn',
                         'text': 'Apenas 1 lista. Adicione mais bullet points e listas numeradas para facilitar extração por ChatGPT e Perplexity.'})
        issues += 1
    else:
        criteria.append({'cat': 'GEO', 'item': 'Listas estruturadas', 'status': 'err',
                         'text': 'Sem listas detectadas. Use bullet points e listas numeradas para que IAs extraiam informações facilmente.'})
        issues += 2

    if paragraphs and len(paragraphs[0]) > 50:
        criteria.append({'cat': 'GEO', 'item': 'Resposta direta no início', 'status': 'ok',
                         'text': 'Conteúdo inicia com resposta objetiva — favorável para featured snippets e citações por IAs.'})
    else:
        criteria.append({'cat': 'GEO', 'item': 'Resposta direta no início', 'status': 'warn',
                         'text': 'Resposta principal não clara nos primeiros parágrafos. IAs priorizam conteúdo que responde diretamente na abertura.'})
        issues += 1

    # ── Adequação Público CIO/CTO ──────────────────────────────
    biz_terms = ['roi','retorno','resultado','negócio','estratégia','gestão','governança',
                 'eficiência','redução','custo','receita','crescimento','vantagem',
                 'decisão','executivo','liderança','board','indicador','kpi','meta',
                 'produtividade','competitiv','rentabilidade','valor','escalabilidade']
    biz_hits = sum(1 for t in biz_terms if t in full_text.lower())
    if biz_hits >= 5:
        criteria.append({'cat': 'Público CIO/CTO', 'item': 'Linguagem de negócio', 'status': 'ok',
                         'text': 'Boa presença de termos de negócio — adequado para executivos de tecnologia sênior.'})
        strengths.append('Linguagem orientada a resultados de negócio')
    elif biz_hits >= 2:
        criteria.append({'cat': 'Público CIO/CTO', 'item': 'Linguagem de negócio', 'status': 'warn',
                         'text': 'Pouca ênfase em resultados de negócio. Adicione mais referências a ROI, governança e impacto estratégico.'})
        issues += 1
    else:
        criteria.append({'cat': 'Público CIO/CTO', 'item': 'Linguagem de negócio', 'status': 'err',
                         'text': 'Foco excessivo em features técnicas sem conectar a resultados de negócio. Risco de não engajar CIOs/CTOs.'})
        issues += 2

    # ── CTA ────────────────────────────────────────────────────
    cta_terms = ['entre em contato','fale com','saiba mais','acesse','baixe','cadastre',
                 'inscreva','solicite','converse','agende','clique','veja mais','descubra',
                 'entre em contato','fale conosco','leia mais']
    if any(t in full_text.lower() for t in cta_terms):
        criteria.append({'cat': 'CTA', 'item': 'Chamada para ação', 'status': 'ok',
                         'text': 'CTA detectado — o conteúdo direciona o leitor para um próximo passo claro.'})
    else:
        criteria.append({'cat': 'CTA', 'item': 'Chamada para ação', 'status': 'warn',
                         'text': 'Nenhum CTA identificado. Adicione uma chamada clara para o próximo passo (ex: "Fale com nossos especialistas").'})
        issues += 1

    verdict = 'ok' if issues == 0 else ('warn' if issues <= 3 else 'err')
    priorities = [c['text'] for c in criteria if c['status'] != 'ok']

    return {
        'verdict': verdict,
        'strengths': strengths,
        'criteria': criteria,
        'priorities': priorities[:5],
    }


# ── GEO Content Analysis ───────────────────────────────────────────────────────

def analyze_content(soup, title=''):
    """Full content analysis: GEO insights + keywords + deep analysis + pautas."""
    full_text  = soup.get_text(' ', strip=True)
    words      = full_text.split()
    word_count = len(words)

    paragraphs = [p.get_text(strip=True) for p in soup.find_all('p')
                  if len(p.get_text(strip=True)) > 40]
    headings   = soup.find_all(['h2', 'h3', 'h4'])

    geo_insights = []
    issues       = 0

    # 1. Direct answers
    has_direct = bool(paragraphs and 40 <= len(paragraphs[0]) <= 400)
    if not has_direct:
        geo_insights.append({'criteria': 'Respostas Diretas', 'status': 'warn',
                             'text': 'Inclua uma resposta objetiva nos primeiros parágrafos. IAs como ChatGPT e Perplexity priorizam conteúdo que responde perguntas diretamente.'})
        issues += 1
    else:
        geo_insights.append({'criteria': 'Respostas Diretas', 'status': 'ok',
                             'text': 'Conteúdo inicia com resposta direta e concisa — favorável para citação por IAs generativas.'})

    # 2. Jargon
    alpha_words  = [w for w in words if w.isalpha()]
    long_words   = [w for w in alpha_words if len(w) > 13]
    jargon_ratio = len(long_words) / max(len(alpha_words), 1)
    if jargon_ratio > 0.07:
        geo_insights.append({'criteria': 'Jargão Técnico', 'status': 'warn',
                             'text': f'Alto índice de termos complexos ({round(jargon_ratio * 100)}% das palavras). Explique jargões para tornar o conteúdo mais acessível.'})
        issues += 1
    else:
        geo_insights.append({'criteria': 'Jargão Técnico', 'status': 'ok',
                             'text': 'Linguagem acessível — conteúdo claro para IAs e leitores em geral.'})

    # 3. Structure
    h_count = len(headings)
    if h_count < 2:
        geo_insights.append({'criteria': 'Estrutura Clara', 'status': 'err',
                             'text': 'Poucos subtítulos detectados. Use H2/H3 para organizar seções — IAs escaneiam hierarquias e extraem respostas por tópico.'})
        issues += 2
    else:
        geo_insights.append({'criteria': 'Estrutura Clara', 'status': 'ok',
                             'text': f'{h_count} subtítulos detectados — boa hierarquia de conteúdo para indexação por IAs.'})

    # 4. Data and evidence
    matches = re.findall(
        r'\b\d{1,3}(?:\.\d{3})*(?:,\d+)?\s*%'
        r'|\b\d+\s*(?:mil|bilh|trilh)'
        r'|\b(?:19|20)\d{2}\b'
        r'|\b\d+x\b',
        full_text, re.IGNORECASE)
    if len(matches) < 2:
        geo_insights.append({'criteria': 'Dados e Evidências', 'status': 'warn',
                             'text': 'Poucos dados quantitativos. Adicione estatísticas e percentuais — IAs tendem a citar fontes com dados concretos.'})
        issues += 1
    else:
        geo_insights.append({'criteria': 'Dados e Evidências', 'status': 'ok',
                             'text': 'Bom uso de dados quantitativos — favorável para IAs que priorizam conteúdo factual.'})

    geo_status = 'ok' if issues == 0 else ('warn' if issues <= 2 else 'err')

    # New: keywords, deep analysis, pautas
    keywords     = extract_keywords(soup)
    deep_analysis = analyze_content_deep(soup, title, keywords)
    pautas       = generate_pautas(title, keywords, soup)

    return {
        'geo_status':    geo_status,
        'word_count':    word_count,
        'insights':      geo_insights,
        'keywords':      keywords,
        'deep_analysis': deep_analysis,
        'pautas':        pautas,
    }


# ── Crawl ──────────────────────────────────────────────────────────────────────

def crawl_url(url):
    url_len    = len(url)
    url_status = get_url_status(url)
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'lxml')

        title_tag  = soup.find('title')
        title      = title_tag.get_text().strip() if title_tag else ''

        meta_tag   = soup.find('meta', attrs={'name': lambda x: x and x.lower() == 'description'})
        meta_desc  = meta_tag.get('content', '').strip() if meta_tag else ''

        h1_count   = count_visible_h1s(soup)
        h2h3_count = len(soup.find_all('h2')) + len(soup.find_all('h3'))

        t_status  = get_title_status(title)
        m_status  = get_meta_status(meta_desc)
        h1_stat   = get_h1_status(h1_count)
        h23_stat  = get_h2h3_status(h2h3_count)

        statuses    = [t_status, m_status, h1_stat, h23_stat]
        error_count = sum(1 for s in statuses if s in ('absent', 'long'))
        warn_count  = sum(1 for s in statuses if s in ('short', 'multiple', 'warn'))

        content_analysis = analyze_content(soup, title)

        return {
            'url': url, 'url_len': url_len, 'url_status': url_status,
            'title': title, 'title_len': len(title), 'title_status': t_status,
            'meta_desc': meta_desc, 'meta_desc_len': len(meta_desc), 'meta_status': m_status,
            'h1_count': h1_count, 'h1_status': h1_stat,
            'h2h3_count': h2h3_count, 'h2h3_status': h23_stat,
            'error_count': error_count, 'warn_count': warn_count,
            'crawl_error': None, 'content_analysis': content_analysis,
        }
    except Exception as e:
        return {
            'url': url, 'url_len': url_len, 'url_status': url_status,
            'title': None, 'title_len': 0, 'title_status': 'absent',
            'meta_desc': None, 'meta_desc_len': 0, 'meta_status': 'absent',
            'h1_count': 0, 'h1_status': 'absent',
            'h2h3_count': 0, 'h2h3_status': 'absent',
            'error_count': 4, 'warn_count': 0,
            'crawl_error': str(e),
            'content_analysis': {'geo_status': 'err', 'word_count': 0, 'insights': [],
                                  'keywords': [], 'deep_analysis': {}, 'pautas': []},
        }


def _insert_result(c, audit_id, r):
    c.execute('''INSERT INTO audit_results
        (audit_id, url, title, title_len, title_status, meta_desc, meta_desc_len,
         meta_status, h1_count, h1_status, h2h3_count, h2h3_status,
         error_count, warn_count, crawl_error, url_len, url_status, content_analysis)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (audit_id, r['url'], r['title'], r['title_len'], r['title_status'],
         r['meta_desc'], r['meta_desc_len'], r['meta_status'],
         r['h1_count'], r['h1_status'], r['h2h3_count'], r['h2h3_status'],
         r['error_count'], r['warn_count'], r['crawl_error'],
         r.get('url_len', len(r['url'])),
         r.get('url_status', get_url_status(r['url'])),
         json.dumps(r.get('content_analysis', {}))))


def _deserialize_results(results):
    """Re-calculate statuses from source data and deserialize JSON."""
    for r in results:
        title = r.get('title') or ''
        meta  = r.get('meta_desc') or ''
        url   = r.get('url') or ''

        r['title_status']  = get_title_status(title)
        r['meta_status']   = get_meta_status(meta)
        r['h1_status']     = get_h1_status(r.get('h1_count') or 0)
        r['h2h3_status']   = get_h2h3_status(r.get('h2h3_count') or 0)
        r['url_len']       = r.get('url_len') or len(url)
        r['url_status']    = get_url_status(url)

        statuses = [r['title_status'], r['meta_status'], r['h1_status'], r['h2h3_status']]
        r['error_count'] = sum(1 for s in statuses if s in ('absent', 'long'))
        r['warn_count']  = sum(1 for s in statuses if s in ('short', 'multiple', 'warn'))

        ca = r.get('content_analysis')
        if ca and isinstance(ca, str):
            try:
                r['content_analysis'] = json.loads(ca)
            except Exception:
                r['content_analysis'] = {}
        elif not ca:
            r['content_analysis'] = {}
    return results


def run_audit(audit_id, urls):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('UPDATE audits SET total_urls=?, status=? WHERE id=?',
              (len(urls), 'running', audit_id))
    conn.commit()

    for i, url in enumerate(urls):
        r = crawl_url(url.strip())
        _insert_result(c, audit_id, r)
        c.execute('UPDATE audits SET completed_urls=? WHERE id=?', (i + 1, audit_id))
        conn.commit()

    c.execute('UPDATE audits SET status=? WHERE id=?', ('completed', audit_id))
    conn.commit()
    conn.close()


# ── Static assets ──────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/vertigo_logo_small.png')
def serve_logo_small():
    return send_from_directory(str(Path(__file__).parent), 'vertigo_logo_small.png')


# ── API Routes ─────────────────────────────────────────────────────────────────

@app.route('/api/audit', methods=['POST'])
def start_audit():
    audit_id   = str(uuid.uuid4())
    created_at = datetime.now().isoformat()
    urls = []
    name = ''

    if 'file' in request.files:
        file    = request.files['file']
        fname   = (file.filename or '').lower()
        content = file.read()
        name    = file.filename

        if fname.endswith('.xlsx') or fname.endswith('.xls'):
            # Parse Excel file
            try:
                wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
                ws = wb.active
                for row in ws.iter_rows(values_only=True):
                    if row and row[0]:
                        cell = str(row[0]).strip()
                        if cell.startswith('http'):
                            urls.append(cell)
            except Exception as e:
                return jsonify({'error': f'Erro ao processar XLSX: {str(e)}'}), 400
        else:
            # Parse CSV / text file
            try:
                decoded = content.decode('utf-8-sig')
            except Exception:
                decoded = content.decode('latin-1', errors='replace')
            reader = csv.reader(io.StringIO(decoded))
            for row in reader:
                if row:
                    cell = row[0].strip()
                    if cell.startswith('http'):
                        urls.append(cell)
    else:
        data = request.get_json(silent=True) or {}
        if 'urls' in data:
            for u in data.get('urls', []):
                u = u.strip()
                if u:
                    if not u.startswith('http'):
                        u = 'https://' + u
                    urls.append(u)
            name = f'{len(urls)} URLs' if len(urls) > 1 else (urls[0] if urls else '')
        else:
            url = data.get('url', '').strip()
            if url:
                if not url.startswith('http'):
                    url = 'https://' + url
                urls = [url]
                name = url

    # Deduplicate while preserving order
    seen = set()
    unique_urls = []
    for u in urls:
        norm = u.rstrip('/')
        if norm not in seen:
            seen.add(norm)
            unique_urls.append(u)
    urls = unique_urls

    if not urls:
        return jsonify({'error': 'Nenhuma URL válida encontrada'}), 400

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        'INSERT INTO audits (id, created_at, name, status, total_urls, completed_urls) VALUES (?,?,?,?,?,?)',
        (audit_id, created_at, name, 'pending', len(urls), 0))
    conn.commit()
    conn.close()

    threading.Thread(target=run_audit, args=(audit_id, urls), daemon=True).start()
    return jsonify({'audit_id': audit_id})


@app.route('/api/audit/<audit_id>/add', methods=['POST'])
def add_url(audit_id):
    data = request.get_json(silent=True) or {}
    url  = data.get('url', '').strip()
    if not url:
        return jsonify({'error': 'URL inválida'}), 400
    if not url.startswith('http'):
        url = 'https://' + url

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT id FROM audits WHERE id=?', (audit_id,))
    if not c.fetchone():
        conn.close()
        return jsonify({'error': 'Auditoria não encontrada'}), 404

    # Check for duplicate URL (normalize trailing slash)
    url_norm = url.rstrip('/')
    c.execute('SELECT id FROM audit_results WHERE audit_id=? AND (url=? OR url=?)',
              (audit_id, url, url_norm))
    if c.fetchone():
        conn.close()
        return jsonify({'duplicate': True, 'message': 'URL já auditada — ignorada.'}), 200

    r = crawl_url(url)
    _insert_result(c, audit_id, r)
    c.execute(
        'UPDATE audits SET total_urls=total_urls+1, completed_urls=completed_urls+1 WHERE id=?',
        (audit_id,))
    conn.commit()
    conn.close()
    return jsonify({'result': r})


@app.route('/api/audit/<audit_id>/result/<int:result_id>', methods=['DELETE'])
def delete_result(audit_id, result_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM audit_results WHERE id=? AND audit_id=?', (result_id, audit_id))
    if c.rowcount > 0:
        c.execute(
            'UPDATE audits SET total_urls=MAX(0,total_urls-1), completed_urls=MAX(0,completed_urls-1) WHERE id=?',
            (audit_id,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/audit/<audit_id>/status')
def audit_status(audit_id):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM audits WHERE id=?', (audit_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        return jsonify({'error': 'Não encontrado'}), 404
    return jsonify(dict(row))


@app.route('/api/audit/<audit_id>/results')
def audit_results(audit_id):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM audits WHERE id=?', (audit_id,))
    audit = c.fetchone()
    c.execute('SELECT * FROM audit_results WHERE audit_id=? ORDER BY error_count DESC, warn_count DESC',
              (audit_id,))
    results = _deserialize_results([dict(r) for r in c.fetchall()])
    conn.close()
    if not audit:
        return jsonify({'error': 'Não encontrado'}), 404
    return jsonify({'audit': dict(audit), 'results': results})


@app.route('/api/history')
def history():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM audits ORDER BY created_at DESC')
    audits = [dict(r) for r in c.fetchall()]
    for a in audits:
        if a['status'] == 'completed' and a['total_urls'] > 0:
            c.execute('''SELECT
                SUM(CASE WHEN title_status  != "ok" THEN 1 ELSE 0 END)*100.0/COUNT(*) as title_pct,
                SUM(CASE WHEN meta_status   != "ok" THEN 1 ELSE 0 END)*100.0/COUNT(*) as meta_pct,
                SUM(CASE WHEN h1_status     != "ok" THEN 1 ELSE 0 END)*100.0/COUNT(*) as h1_pct,
                SUM(CASE WHEN h2h3_status   != "ok" THEN 1 ELSE 0 END)*100.0/COUNT(*) as h2h3_pct,
                SUM(error_count) as total_errors
                FROM audit_results WHERE audit_id=?''', (a['id'],))
            row = c.fetchone()
            if row:
                a['stats'] = dict(row)
    conn.close()
    return jsonify(audits)


@app.route('/api/audit/<audit_id>/delete', methods=['DELETE'])
def delete_audit(audit_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM audit_results WHERE audit_id=?', (audit_id,))
    c.execute('DELETE FROM audits WHERE id=?', (audit_id,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/api/audit/<audit_id>/export')
def export_xlsx(audit_id):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM audits WHERE id=?', (audit_id,))
    audit = dict(c.fetchone() or {})
    c.execute('SELECT * FROM audit_results WHERE audit_id=? ORDER BY error_count DESC', (audit_id,))
    results = [dict(r) for r in c.fetchall()]
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Auditoria SEO'

    cols = [
        'URL', 'Empresa', 'Tam. URL', 'Status URL',
        'Title', 'Tam. Title', 'Status Title',
        'Meta Description', 'Tam. Meta', 'Status Meta',
        'Qtd H1', 'Status H1', 'Qtd H2+H3', 'Status H2/H3',
        'Erros (❌)', 'Alertas (⚠️)', 'GEO Score',
        'Palavras-chave', 'Parecer Conteúdo', 'Erro de Crawl',
    ]
    hdr_fill = PatternFill(start_color='0C1557', end_color='0C1557', fill_type='solid')
    hdr_font = Font(bold=True, color='E0E0E0', size=11)
    for col, h in enumerate(cols, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 20

    fill_ok   = PatternFill(start_color='D1FAE5', end_color='D1FAE5', fill_type='solid')
    fill_warn = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')
    fill_err  = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid')
    status_labels = {'ok':'✅ OK','short':'⚠️ Curto','long':'❌ Longo','absent':'❌ Ausente',
                     'multiple':'⚠️ Múltiplos','warn':'⚠️ Atenção','err':'❌ Problema'}
    status_fills  = {'ok':fill_ok,'short':fill_warn,'long':fill_err,'absent':fill_err,
                     'multiple':fill_warn,'warn':fill_warn,'err':fill_err}

    for ri, r in enumerate(results, 2):
        try:
            ca = json.loads(r.get('content_analysis') or '{}')
        except Exception:
            ca = {}
        geo     = ca.get('geo_status', '')
        kws     = ', '.join(ca.get('keywords', []))
        deep_v  = (ca.get('deep_analysis') or {}).get('verdict', '')

        try:
            from urllib.parse import urlparse
            domain = urlparse(r['url']).hostname or ''
            domain = domain.replace('www.', '')
        except Exception:
            domain = ''

        row_data = [
            r['url'], domain, r.get('url_len', len(r['url'])),
            status_labels.get(r.get('url_status',''), r.get('url_status','')),
            r['title'] or '', r['title_len'],
            status_labels.get(r['title_status'], r['title_status']),
            r['meta_desc'] or '', r['meta_desc_len'],
            status_labels.get(r['meta_status'], r['meta_status']),
            r['h1_count'], status_labels.get(r['h1_status'], r['h1_status']),
            r['h2h3_count'], status_labels.get(r['h2h3_status'], r['h2h3_status']),
            r['error_count'], r['warn_count'],
            status_labels.get(geo, geo),
            kws,
            status_labels.get(deep_v, deep_v),
            r['crawl_error'] or '',
        ]
        status_cols = {4: r.get('url_status',''), 7: r['title_status'],
                       10: r['meta_status'], 12: r['h1_status'],
                       14: r['h2h3_status'], 17: geo, 19: deep_v}
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            if ci in status_cols:
                cell.fill = status_fills.get(status_cols[ci], PatternFill())
            cell.alignment = Alignment(wrap_text=(ci in (1, 5, 8, 18)), vertical='top')

    for col in ws.columns:
        max_len = max((len(str(cell.value or '')) for cell in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 80)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    ts = audit.get('created_at', '')[:10]
    return send_file(output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=f'auditoria-seo-{ts}-{audit_id[:6]}.xlsx')


# ── Global error handlers (always return JSON) ─────────────────────────────────

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Rota não encontrada'}), 404


@app.errorhandler(405)
def method_not_allowed(e):
    return jsonify({'error': 'Método não permitido'}), 405


@app.errorhandler(500)
def internal_error(e):
    return jsonify({'error': f'Erro interno: {str(e)}'}), 500


if __name__ == '__main__':
    init_db()
    app.run(debug=False, port=5050, threaded=True, use_reloader=True)
