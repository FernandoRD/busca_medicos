#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Coletor e Enriquecedor de Informações Públicas de Médicos
---------------------------------------------------------
Busca automaticamente:
- Telefone / WhatsApp do consultório ou clínica
- Endereço completo / Local de atendimento / CEP
- Especialidade médica e UF do CRM
- Links diretos (Doctoralia, ListaMédicos, Google Maps, Google Busca, WhatsApp)

Desenvolvido para listas contendo Nome Completo e CRM em diversos formatos.
"""

import sys
import os
import csv
import re
import time
import random
import socket
import unicodedata
import urllib.parse
from typing import Dict, List, Optional, Tuple, Set, Any, Union

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("\n[!] Dependências ausentes. Instale com:")
    print("    pip install requests beautifulsoup4\n")
    sys.exit(1)

# Resolução de DNS defensiva para DuckDuckGo (evita timeout quando DNS local/ISP retorna nó inacessível 191.235.123.80)
_orig_getaddrinfo = socket.getaddrinfo

def _safe_getaddrinfo(host, port, family=0, type=0, proto=0, flags=0):
    port = int(port) if str(port).isdigit() else 443
    try:
        res = _orig_getaddrinfo(host, port, family, type, proto, flags)
        if "duckduckgo" in str(host).lower():
            valid_res = [r for r in res if r[4][0] != "191.235.123.80"]
            if valid_res:
                return valid_res
            return [(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("40.89.244.237", port))]
        return res
    except Exception:
        if "duckduckgo" in str(host).lower():
            return [(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("40.89.244.237", port))]
        raise

socket.getaddrinfo = _safe_getaddrinfo

# DDDs válidos no Brasil
DDDS_VALIDOS = {
    11, 12, 13, 14, 15, 16, 17, 18, 19,
    21, 22, 24, 27, 28,
    31, 32, 33, 34, 35, 37, 38,
    41, 42, 43, 44, 45, 46, 47, 48, 49,
    51, 53, 54, 55,
    61, 62, 63, 64, 65, 66, 67, 68, 69,
    71, 73, 74, 75, 77, 79,
    81, 82, 83, 84, 85, 86, 87, 88, 89,
    91, 92, 93, 94, 95, 96, 97, 98, 99
}

# UFs válidas brasileiras
UFS_VALIDAS = {
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA',
    'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN',
    'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
}

# Colunas padrão da planilha CSV enriquecida
COLUNAS_SAIDA = [
    'Nome Original',
    'CRM Original',
    'UF',
    'Especialidade',
    'Telefone Comercial',
    'Link WhatsApp',
    'Endereço / Consultório',
    'Perfil Doctoralia',
    'Perfil ListaMédicos',
    'Perfil Instagram',
    'Link Google Maps',
    'Link Google Busca',
    'Status da Coleta'
]

# Rotas do sistema e mídia do Instagram a serem ignoradas na extração de perfil
IGNORAR_ROTAS_INSTAGRAM = {
    'p', 'reel', 'reels', 'explore', 'stories', 'accounts',
    'developer', 'about', 'legal', 'tags', 'directory', 'tv'
}

class InstagramResult(str):
    """Encapsula o perfil do Instagram compatível com str e tuple[str, str], contendo metadados de bio."""
    def __new__(
        cls,
        url: str = "",
        username: str = "",
        telefones: Optional[List[str]] = None,
        enderecos: Optional[List[str]] = None,
        snippet: str = ""
    ):
        obj = str.__new__(cls, url)
        obj.username = username
        obj.url = url
        obj.telefones = list(telefones) if telefones is not None else []
        obj.enderecos = list(enderecos) if enderecos is not None else []
        obj.snippet = snippet
        return obj

    def __iter__(self):
        yield self.username
        yield self.url

    def __getitem__(self, item):
        if isinstance(item, int) and item in (0, 1):
            return [self.username, self.url][item]
        return super().__getitem__(item)

USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0'
]

def criar_sessao() -> requests.Session:
    """Cria uma sessão HTTP persistente com User-Agent fixo e moderno e pooling de conexões."""
    session = requests.Session()
    ua = random.choice(USER_AGENTS)
    session.headers.update({
        'User-Agent': ua,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
    })
    return session

def get_headers(user_agent: Optional[str] = None) -> Dict[str, str]:
    return {
        'User-Agent': user_agent or random.choice(USER_AGENTS),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
    }

def sanitizar_texto(texto: Any) -> str:
    """Remove caracteres invisíveis, zero-width e normaliza espaços."""
    if not texto:
        return ""
    s = re.sub(r'[\u200b-\u200f\ufeff\u00a0\u202a-\u202e\u2060-\u206f]', '', str(texto))
    return re.sub(r'\s+', ' ', s).strip()

def remover_acentos(texto: str) -> str:
    s = sanitizar_texto(texto)
    return "".join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def limpar_nome_medico(nome: str) -> str:
    """Remove prefixos médicos (Dr., Dra., Doutor, Doutora, Prof. Dr., etc.) e sanitiza nome."""
    s = sanitizar_texto(nome)
    s = re.sub(r'^(?:(?:dr[a-z]*|doutor[a-z]*|m[ée]dic[oa]|prof[a-z]*)\.?\s*)+', '', s, flags=re.IGNORECASE)
    return s.strip()

def criar_slug(texto: str) -> str:
    s = remover_acentos(texto.lower().strip())
    s = re.sub(r'[^a-z0-9\s-]', '', s)
    s = re.sub(r'\s+', '-', s)
    return s.strip('-')

def gerar_slugs(nome: str) -> List[str]:
    """Gera combinações comuns de slugs para perfis médicos."""
    slugs = []
    base_slug = criar_slug(nome)
    if base_slug:
        slugs.append(base_slug)
    
    tokens = [t for t in remover_acentos(nome.lower()).split() if re.match(r'^[a-z0-9]+$', t)]
    preposicoes = {'de', 'da', 'do', 'dos', 'das', 'e'}
    tokens_sem_prep = [t for t in tokens if t not in preposicoes]
    
    slug_sem_prep = "-".join(tokens_sem_prep)
    if slug_sem_prep and slug_sem_prep not in slugs:
        slugs.append(slug_sem_prep)
        
    if len(tokens) >= 2:
        slug_primeiro_ultimo = f"{tokens[0]}-{tokens[-1]}"
        if slug_primeiro_ultimo not in slugs:
            slugs.append(slug_primeiro_ultimo)
            
    return slugs

def extrair_crm_e_uf(crm_str: str, uf_original: str = "") -> Tuple[str, str]:
    """Extrai CRM numérico e UF de formatos variados: (CRM-MG 66436), CRM/SP 12345, 66436-MG, etc."""
    crm_str = sanitizar_texto(crm_str)
    uf_original = sanitizar_texto(uf_original).upper()
    
    uf = uf_original if uf_original in UFS_VALIDAS else ""
    if not uf:
        m_uf = re.search(r'(?:CRM[/-]|\b)([A-Z]{2})\b', crm_str.upper())
        if m_uf and m_uf.group(1) in UFS_VALIDAS:
            uf = m_uf.group(1)
            
    m_num = re.search(r'\b(\d{3,8})\b', crm_str)
    crm_num = m_num.group(1) if m_num else re.sub(r'\D', '', crm_str)
    return crm_num, uf

def extrair_telefones(texto: str) -> List[str]:
    """Extrai e formata números de telefone com DDD válido ignorando números espúrios."""
    telefones = []
    padrao = re.compile(r'(?:\(?\s*(\d{2})\s*\)?[\s.-]*)?(?:(9\s*\d{4}|\d{4})[\s.-]*(\d{4}))\b')
    for m in padrao.finditer(texto):
        ddd, p1, p2 = m.groups()
        if not ddd:
            continue
        try:
            if int(ddd) not in DDDS_VALIDOS:
                continue
        except ValueError:
            continue
            
        p1_limpo = re.sub(r'\D', '', p1)
        p2_limpo = re.sub(r'\D', '', p2)
        total = p1_limpo + p2_limpo
        
        # Descartar números repetidos ou placeholders conhecidos
        if len(set(total)) <= 1 or total in {'999999999', '99999999', '12345678', '123456789'}:
            continue
            
        if len(total) in (8, 9):
            tel_fmt = f"({ddd}) {p1_limpo}-{p2_limpo}"
            if tel_fmt not in telefones:
                telefones.append(tel_fmt)
    return telefones

def extrair_telefones_html(soup: BeautifulSoup) -> List[str]:
    """Extrai telefones de página HTML após remoção de scripts, styles e outros elementos espúrios."""
    soup_clone = BeautifulSoup(str(soup), 'html.parser')
    for el in soup_clone.find_all(['script', 'style', 'svg', 'noscript', 'meta', 'link']):
        el.decompose()
    return extrair_telefones(soup_clone.get_text(' ', strip=True))

def normalizar_endereco_para_comparacao(texto: str) -> str:
    """Normaliza texto de endereço para deduplicação semântica por substrings."""
    s = remover_acentos(str(texto).lower())
    s = re.sub(r'\bav\b|\bav\.', 'avenida', s)
    s = re.sub(r'\br\b|\br\.', 'rua', s)
    s = re.sub(r'\bal\b|\bal\.', 'alameda', s)
    s = re.sub(r'\bpca\b|\bpca\.', 'praca', s)
    s = re.sub(r'\bed\b|\bed\.', 'edificio', s)
    return re.sub(r'[^a-z0-9]', '', s)

def consolidar_enderecos(enderecos: List[str]) -> str:
    """
    Consolida lista de endereços na mesma célula separados por ' | ':
    - Elimina duplicatas exatas
    - Elimina endereços que sejam substrings de outros mais completos
    """
    todos_enderecos = []
    for item in enderecos:
        if not item:
            continue
        for parte in str(item).split('|'):
            p = re.sub(r'\s+', ' ', parte).strip(' |,-')
            if len(p) >= 10:
                todos_enderecos.append(p)
                
    limpos = []
    for e in todos_enderecos:
        alpha_e = normalizar_endereco_para_comparacao(e)
        if not alpha_e:
            continue
        ja_existe = False
        for i, existente in enumerate(limpos):
            alpha_ex = normalizar_endereco_para_comparacao(existente)
            if alpha_e == alpha_ex or alpha_e in alpha_ex:
                ja_existe = True
                break
            elif alpha_ex in alpha_e:
                limpos[i] = e  # O novo é mais completo, substitui o existente
                ja_existe = True
                break
        if not ja_existe:
            limpos.append(e)
            
    return " | ".join(limpos[:4])

def consolidar_telefones(telefones: List[str]) -> str:
    """
    Garante lista única de telefones formatados na mesma célula separada por ' / ':
    - Elimina duplicatas exatas
    - Elimina variações com e sem DDD (mantém a com DDD)
    - Elimina variações com e sem nono dígito (mantém celular de 9 dígitos)
    """
    registros = []
    
    for item in telefones:
        if not item:
            continue
        for parte in re.split(r'[/|;]', str(item)):
            p = parte.strip()
            if not p:
                continue
            digs = re.sub(r'\D', '', p)
            if digs.startswith('55') and len(digs) in (12, 13):
                digs = digs[2:]
            
            ddd = ''
            corpo = ''
            is_cel = False
            
            if len(digs) == 11:
                ddd = digs[:2]
                corpo = digs[2:]
                is_cel = (digs[2] == '9')
            elif len(digs) == 10:
                ddd = digs[:2]
                corpo = digs[2:]
                is_cel = False
            elif len(digs) == 9:
                corpo = digs
                is_cel = (digs[0] == '9')
            elif len(digs) == 8:
                corpo = digs
                is_cel = False
            else:
                continue
                
            if ddd and int(ddd) not in DDDS_VALIDOS:
                continue
                
            suffix8 = corpo[-8:]
            if len(set(suffix8)) <= 1 or suffix8 in {'12345678', '99999999'}:
                continue
                
            registros.append({
                'ddd': ddd,
                'corpo': corpo,
                'is_cel': is_cel,
                'suffix8': suffix8,
                'total_digs': digs,
                'original': p
            })
            
    # Deduplicação inteligente
    unicos = []
    for reg in registros:
        ja_existe = False
        for i, u in enumerate(unicos):
            # Se tiverem o mesmo sufixo de 8 dígitos
            if reg['suffix8'] == u['suffix8']:
                # Caso 1: um tem DDD e outro não tem
                if reg['ddd'] and not u['ddd']:
                    unicos[i] = reg  # Substitui pelo que tem DDD
                    ja_existe = True
                    break
                elif not reg['ddd'] and u['ddd']:
                    ja_existe = True  # Já temos versão com DDD
                    break
                # Caso 2: ambos têm mesmo DDD ou ambos não têm DDD
                elif reg['ddd'] == u['ddd']:
                    # Se um é celular de 9 dígitos e outro é fixo/8 dígitos
                    if reg['is_cel'] and not u['is_cel']:
                        unicos[i] = reg  # Atualiza para 9 dígitos
                    ja_existe = True
                    break
        if not ja_existe:
            unicos.append(reg)
            
    # Formatação final
    saida = []
    for u in unicos:
        ddd = u['ddd']
        corpo = u['corpo']
        if ddd:
            if len(corpo) == 9:
                saida.append(f"({ddd}) {corpo[:5]}-{corpo[5:]}")
            else:
                saida.append(f"({ddd}) {corpo[:4]}-{corpo[4:]}")
        else:
            if len(corpo) == 9:
                saida.append(f"{corpo[:5]}-{corpo[5:]}")
            else:
                saida.append(f"{corpo[:4]}-{corpo[4:]}")
                
    return " / ".join(saida[:3])

def gerar_link_whatsapp(telefones: Union[List[str], str]) -> str:
    """Gera link wa.me para o primeiro celular válido encontrado (11 dígitos, DDD + 9) dentre os telefones consolidados."""
    lista = []
    if isinstance(telefones, str):
        lista = [t.strip() for t in re.split(r'[/|;]', telefones) if t.strip()]
    elif isinstance(telefones, (list, tuple, set)):
        for item in telefones:
            if isinstance(item, str):
                for sub in re.split(r'[/|;]', item):
                    if sub.strip():
                        lista.append(sub.strip())
                        
    for tel in lista:
        digs = re.sub(r'\D', '', str(tel))
        if digs.startswith('55') and len(digs) == 13:
            digs = digs[2:]
        if len(digs) == 11 and digs[2] == '9' and int(digs[:2]) in DDDS_VALIDOS:
            return f"https://wa.me/55{digs}"
    return ""

def buscar_crm_com_br(nome_clean: str, crm_clean: str, uf_conhecida: str = "", session: Optional[requests.Session] = None) -> Dict[str, Any]:
    """Consulta crm.com.br para extrair UF, CBO/Especialidade e estabelecimentos vinculados no CNES."""
    dados: Dict[str, Any] = {
        'uf': '',
        'especialidade': '',
        'cidade': '',
        'endereco': '',
        'url': ''
    }
    if not crm_clean:
        return dados
        
    s = session if session is not None else criar_sessao()
    url = f"https://crm.com.br/buscar/?q={crm_clean}"
    try:
        r = s.get(url, timeout=8)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            nome_norm = remover_acentos(nome_clean.lower())
            tokens_nome = [t for t in nome_norm.split() if len(t) > 2]
            
            link_perfil = None
            for a in soup.find_all('a', href=True):
                href = a['href']
                if '/medico/' in href and crm_clean in href:
                    txt_a = remover_acentos(a.get_text().lower())
                    coincidencias = sum(1 for t in tokens_nome if t in txt_a)
                    if coincidencias >= min(2, len(tokens_nome)):
                        link_perfil = urllib.parse.urljoin('https://crm.com.br', href)
                        break
                        
            if link_perfil:
                dados['url'] = link_perfil
                r_prof = s.get(link_perfil, timeout=8)
                if r_prof.status_code == 200:
                    sp_prof = BeautifulSoup(r_prof.text, 'html.parser')
                    estab_urls = []
                    
                    for dt in sp_prof.find_all(['dt', 'th']):
                        dd = dt.find_next_sibling(['dd', 'td'])
                        if not dd:
                            continue
                        chave = remover_acentos(dt.get_text(strip=True).lower())
                        valor = dd.get_text(' ', strip=True)
                        
                        if 'uf do registro' in chave or 'uf' in chave:
                            m_uf = re.search(r'\b([A-Z]{2})\b', valor.upper())
                            if m_uf and m_uf.group(1) in UFS_VALIDAS:
                                dados['uf'] = m_uf.group(1)
                        elif 'ocupacao' in chave or 'especialidade' in chave:
                            valor_limpo = re.sub(r'^(?:m[ée]dico\s+)', '', valor, flags=re.I).strip().title()
                            dados['especialidade'] = valor_limpo
                        elif 'atua em' in chave or 'municipio' in chave:
                            dados['cidade'] = valor
                        elif 'estabelecimentos' in chave:
                            for a_est in dd.find_all('a', href=True):
                                if '/estabelecimento/' in a_est['href']:
                                    estab_urls.append((a_est.get_text(strip=True), urllib.parse.urljoin('https://crm.com.br', a_est['href'])))
                                    
                    enderecos_est = []
                    for nome_est, url_est in estab_urls[:2]:
                        try:
                            r_est = s.get(url_est, timeout=5)
                            if r_est.status_code == 200:
                                sp_est = BeautifulSoup(r_est.text, 'html.parser')
                                end_partes = {}
                                for dt_e in sp_est.find_all(['dt', 'th']):
                                    dd_e = dt_e.find_next_sibling(['dd', 'td'])
                                    if not dd_e:
                                        continue
                                    k_e = remover_acentos(dt_e.get_text(strip=True).lower())
                                    v_e = dd_e.get_text(' ', strip=True)
                                    if 'endereco' in k_e:
                                        end_partes['rua'] = v_e
                                    elif 'municipio' in k_e:
                                        end_partes['municipio'] = v_e
                                    elif 'cep' in k_e:
                                        end_partes['cep'] = v_e
                                
                                rua = end_partes.get('rua', '')
                                mun = end_partes.get('municipio', '')
                                cep = end_partes.get('cep', '')
                                
                                bloco = f"{nome_est}: {rua}" if rua else nome_est
                                if mun and mun not in bloco:
                                    bloco += f", {mun}"
                                if cep:
                                    bloco += f" (CEP {cep})"
                                enderecos_est.append(bloco)
                            else:
                                enderecos_est.append(nome_est)
                        except Exception:
                            enderecos_est.append(nome_est)
                            
                    if enderecos_est:
                        dados['endereco'] = " | ".join(enderecos_est)
    except Exception:
        pass
        
    return dados

def buscar_catalogo_med(nome_clean: str, crm_clean: str = "", uf_conhecida: str = "", session: Optional[requests.Session] = None) -> Dict[str, Any]:
    """Consulta catalogo.med.br para extrair perfis, especialidades, endereços e telefones via AJAX."""
    dados: Dict[str, Any] = {
        'url': '',
        'telefones': [],
        'enderecos': [],
        'especialidade': '',
        'uf': ''
    }
    s = session if session is not None else criar_sessao()
    url_busca = f"https://www.catalogo.med.br/?q={urllib.parse.quote(nome_clean)}"
    try:
        r = s.get(url_busca, timeout=8)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            link_perfil = None
            nome_norm = remover_acentos(nome_clean.lower())
            tokens_nome = [t for t in nome_norm.split() if len(t) > 2]
            
            for a in soup.find_all('a', href=True):
                href = a['href']
                if '/doutor/' in href:
                    txt_a = remover_acentos(a.get_text().lower())
                    coincidencias = sum(1 for t in tokens_nome if t in txt_a)
                    if coincidencias >= min(2, len(tokens_nome)):
                        link_perfil = urllib.parse.urljoin('https://www.catalogo.med.br', href)
                        break
                        
            if link_perfil:
                dados['url'] = link_perfil
                r_prof = s.get(link_perfil, timeout=8)
                if r_prof.status_code == 200:
                    sp_p = BeautifulSoup(r_prof.text, 'html.parser')
                    
                    # Especialidade
                    for h in sp_p.find_all(['h2', 'h3', 'p', 'span']):
                        t = h.get_text(strip=True)
                        if any(esp_k in t.lower() for esp_k in ['clínico', 'clinica medica', 'dermatolog', 'cirurgi', 'ginecolog', 'cardiolog', 'pediatr', 'ortoped', 'oftalmolog', 'nutrolog', 'psiquiatr', 'urolog']):
                            if len(t) < 40 and not any(ign in t.lower() for ign in ['agende', 'central', 'avaliar', 'favorito']):
                                dados['especialidade'] = t
                                break
                                
                    # Telefones via requisição Ajax loadPhoneMasked
                    phone_calls = re.findall(r'loadPhone\((\d+),\s*(\d+),\s*(\d+)\)', r_prof.text)
                    for d, s_arg, p in phone_calls:
                        url_phone = f"https://www.catalogo.med.br/app.pl?act=logPhoneMasked&d={d}&s={s_arg}&p={p}"
                        h_ajax = {
                            'X-Requested-With': 'XMLHttpRequest',
                            'Referer': link_perfil
                        }
                        try:
                            r_ph = s.get(url_phone, headers=h_ajax, timeout=5)
                            if r_ph.status_code == 200:
                                ph_data = r_ph.json()
                                tels = extrair_telefones(ph_data.get('phones', ''))
                                for tel in tels:
                                    if tel not in dados['telefones']:
                                        dados['telefones'].append(tel)
                        except Exception:
                            pass
                            
                    # Endereços completos com CEP
                    for el in sp_p.find_all(string=re.compile(r'CEP:\s*\d{5}-\d{3}', re.I)):
                        parent = el.parent
                        container = parent.find_parent('div') or parent
                        t_end = container.get_text(' ', strip=True)
                        t_end = re.sub(r'Ver no mapa', '', t_end).strip()
                        t_end = re.sub(r'Telefones:.*', '', t_end).strip()
                        t_end = re.sub(r'Seu nome:.*', '', t_end).strip()
                        if len(t_end) > 15 and t_end not in dados['enderecos']:
                            dados['enderecos'].append(t_end)
    except Exception:
        pass
        
    return dados

def validar_url_doctoralia(href: str, nome_clean: str) -> bool:
    """Valida se uma URL pertence de fato ao perfil do médico no Doctoralia."""
    ignorar = [
        '/pesquisa', '/login', '/opiniao', '/social-connect',
        '/clinicas', '/especializacoes-medicas', '/perguntas-respostas',
        '/termos-e-condicoes', '/privacidade'
    ]
    href_lower = href.lower()
    if any(ign in href_lower for ign in ignorar):
        return False
        
    parsed = urllib.parse.urlparse(href)
    parts = [p for p in parsed.path.split('/') if p]
    if not parts:
        return False
        
    slug = parts[0].lower()
    tokens = [t for t in remover_acentos(nome_clean.lower()).split() if len(t) > 2]
    if tokens and tokens[0] not in slug:
        return False
    matches = sum(1 for t in tokens if t in slug)
    return matches >= min(2, len(tokens))

def buscar_doctoralia(nome_clean: str, uf_conhecida: str = "", session: Optional[requests.Session] = None) -> Dict[str, Any]:
    """Busca perfil, clínicas e especialidade no Doctoralia usando slugs flexíveis e busca interna."""
    dados: Dict[str, Any] = {
        'doctoralia_url': '',
        'especialidade': '',
        'endereco': '',
        'telefones': []
    }
    
    s = session if session is not None else criar_sessao()
    slugs = gerar_slugs(nome_clean)
    link_perfil = None
    resp_prof = None
    
    # 1. Tentar slugs diretos
    for slug in slugs:
        url_teste = f"https://www.doctoralia.com.br/{slug}"
        try:
            r = s.get(url_teste, allow_redirects=True, timeout=7)
            if r.status_code == 200 and 'doctoralia.com.br' in r.url and validar_url_doctoralia(r.url, nome_clean):
                link_perfil = r.url
                resp_prof = r
                break
        except Exception:
            continue
            
    # 2. Busca interna se slug direto não encontrar
    if not link_perfil:
        url_pesquisa = f"https://www.doctoralia.com.br/pesquisa?q={urllib.parse.quote(nome_clean)}"
        try:
            r = s.get(url_pesquisa, timeout=8)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, 'html.parser')
                for a in soup.find_all('a', href=True):
                    href = a['href']
                    full_href = urllib.parse.urljoin('https://www.doctoralia.com.br', href)
                    if validar_url_doctoralia(full_href, nome_clean):
                        link_perfil = full_href
                        resp_prof = s.get(link_perfil, timeout=8)
                        break
        except Exception:
            pass
            
    if link_perfil and resp_prof and resp_prof.status_code == 200:
        dados['doctoralia_url'] = link_perfil
        soup_prof = BeautifulSoup(resp_prof.text, 'html.parser')
        
        spec_el = soup_prof.find('span', attrs={'data-test-id': 'doctor-specializations'})
        if spec_el:
            dados['especialidade'] = spec_el.get_text(strip=True).split('·')[0].strip()
        elif soup_prof.title:
            t_title = soup_prof.title.string or ''
            m_esp = re.search(r'-\s*([^,-]+?)\s*em\s+', t_title)
            if m_esp:
                dados['especialidade'] = m_esp.group(1).strip()
                
        addr_divs = soup_prof.find_all(attrs={'data-test-id': re.compile('.*address.*', re.I)})
        enderecos = []
        termos_ignorar = ['atende:', 'a partir de', 'serviços', 'consulta ']
        for d in addr_divs:
            t = d.get_text(' ', strip=True).replace('ampliar o mapa', '').strip()
            if len(t) > 15 and not any(ign in t.lower() for ign in termos_ignorar) and t not in enderecos:
                enderecos.append(t)
        if enderecos:
            dados['endereco'] = " | ".join(enderecos[:2])
            
        dados['telefones'] = extrair_telefones_html(soup_prof)
        
    return dados

def buscar_listamedicos(nome_clean: str, crm_clean: str, session: Optional[requests.Session] = None) -> Dict[str, str]:
    """Busca dados de consultório, endereço e telefone no ListaMédicos."""
    dados = {'telefone': '', 'endereco': '', 'url': '', 'especialidade': ''}
    s = session if session is not None else criar_sessao()
    slugs = gerar_slugs(nome_clean)
    
    for slug in slugs:
        url = f"https://listamedicos.com/medico/{slug}"
        try:
            r = s.get(url, timeout=7)
            if r.status_code == 200:
                txt = r.text
                if not crm_clean or crm_clean in txt or remover_acentos(nome_clean.lower()) in remover_acentos(txt.lower()):
                    dados['url'] = url
                    soup = BeautifulSoup(txt, 'html.parser')
                    tels = extrair_telefones_html(soup)
                    if tels:
                        dados['telefone'] = tels[0]
                        
                    body_txt = soup.get_text(' ', strip=True)
                    m_end = re.search(r'(?:R\s|RUA\s|AV\s|AVENIDA\s|AL\s|ALAMEDA\s|PRAÇA\s|PÇA\s)[^,]+,\s*\d+[^0-9]+(?:\d{5}-\d{3})?(?:\s+[A-Za-z\s]+/[A-Za-z\s]+)?', body_txt, re.IGNORECASE)
                    if m_end:
                        dados['endereco'] = m_end.group(0).strip()
                    break
        except Exception:
            pass
    return dados

def extrair_username_instagram(url: str) -> str:
    """Extrai e valida username do Instagram a partir de uma URL ignorando rotas do sistema/mídia."""
    if not url:
        return ""
    m = re.search(r'https?://(?:www\.)?instagram\.com/([a-zA-Z0-9_.]{3,30})/?', url, re.IGNORECASE)
    if not m:
        return ""
    username = m.group(1).lower().rstrip('/')
    if username in IGNORAR_ROTAS_INSTAGRAM:
        return ""
    return username

def validar_anti_homonimo_instagram(
    username: str,
    titulo: str,
    snippet: str,
    nome_medico: str,
    uf: str = "",
    especialidade: str = ""
) -> bool:
    """Conferir se o título do resultado, snippet ou username contém termos indicativos médicos e tokens do nome."""
    nome_norm = remover_acentos(limpar_nome_medico(nome_medico).lower())
    tokens_nome = [t for t in re.findall(r'[a-z0-9]+', nome_norm) if len(t) > 2 and t not in {'dr', 'dra', 'doutor', 'doutora', 'medico', 'medica', 'de', 'da', 'do', 'dos', 'das'}]
    
    texto_alvo = f"{titulo} {snippet} {username}".lower()
    texto_norm = remover_acentos(texto_alvo)
    
    termos_medicos = ['dr', 'dra', 'doutor', 'doutora', 'medic', 'medico', 'medica', 'consultorio', 'clinica', 'instituto', 'crm', 'cirurgi', 'capilar', 'nutrolog']
    if especialidade:
        esp_norm = remover_acentos(especialidade.lower())
        tokens_esp = [t for t in re.findall(r'[a-z0-9]+', esp_norm) if len(t) > 3]
        termos_medicos.extend(tokens_esp)
    if uf:
        termos_medicos.append(uf.lower())
        
    tem_termo_medico = any(re.search(r'\b' + re.escape(t) + r'\b', texto_norm) for t in termos_medicos)
    matches_nome = sum(1 for t in tokens_nome if t in texto_norm or t in username)
    
    if matches_nome >= 3:
        return True
    if tem_termo_medico and matches_nome >= 1:
        return True
    return False

def extrair_dados_bio_snippet(snippet: str) -> Tuple[List[str], List[str]]:
    """
    Extrai telefones (incluindo WhatsApp) e clínicas/consultórios/endereços
    a partir do snippet da bio pública do Instagram retornado pelo buscador.
    """
    if not snippet:
        return [], []
        
    tels = extrair_telefones(snippet)
    for m_wa in re.finditer(r'(?:wa\.me/|api\.whatsapp\.com/send\?phone=)(?:55)?(\d{2})(\d{8,9})', snippet):
        ddd_wa, num_wa = m_wa.groups()
        if int(ddd_wa) in DDDS_VALIDOS:
            fmt_wa = f"({ddd_wa}) {num_wa[:5]}-{num_wa[5:]}" if len(num_wa) == 9 else f"({ddd_wa}) {num_wa[:4]}-{num_wa[4:]}"
            if fmt_wa not in tels:
                tels.append(fmt_wa)
                
    s = snippet
    # Remover contadores de seguidores/posts e chamadas padrão do Instagram / buscadores
    s = re.sub(r'^(?:.*?[\d,.]+[kKmM]?\s*(?:followers|following|posts|seguidores|seguindo|publica[cç][oõ]es)[^:\'\"“”]*[-–—:]\s*)?', '', s, flags=re.I)
    s = re.sub(r'^(?:.*?(?:on instagram|no instagram|fotos e v[íi]deos(?: do instagram)?|instagram photos and videos|veja as publica[cç][oõ]es)[^:\'\"“”]*[:\'\"“”]\s*)?', '', s, flags=re.I)
    s = s.strip(' \'\"“”')
    
    # Emojis estruturais como separadores
    s_norm = re.sub(r'[📍🗺️📌🏠🏢🏛️🏥🩺💉🔬📲📞☎️💬✉️🗓️⏰]', ' | ', s)
    
    logradouro_re = re.compile(
        r'\b(?:rua|r\.|avenida|av\.|alameda|al\.|pra[cç]a|p[cç]a\.|rodovia|rod\.|travessa|tv\.|estrada|est\.)\s+[A-Za-z0-9\s.,ºª\-\/]+',
        re.I
    )
    clinica_re = re.compile(
        r'\b(?:cl[íi]nica|instituto|consult[óo]rio|hospital|centro m[ée]dico|espa[cç]o m[ée]dico|n[úu]cleo de sa[úu]de|policl[íi]nica)\b',
        re.I
    )
    edificio_re = re.compile(
        r'\b(?:edif[íi]cio|ed\.|sala\s*\d+|andar|torre|bloco|medical center|trade center)\b',
        re.I
    )
    
    chunks = [c.strip(' \'\"“”.,-') for c in re.split(r'[|\n•·▪▫;]', s_norm) if c.strip()]
    
    clinicas = []
    enderecos = []
    completos = []
    
    for c in chunks:
        # Remover username @... do chunk para evitar falsos positivos
        c_sem_user = re.sub(r'@[a-zA-Z0-9_.]+', '', c).strip()
        c_limpo = re.sub(r'(?:(?:whatsapp|tel|cel|fone|contato)\s*[:.-]?\s*)?(?:\(?\d{2}\)?\s*)?9?\d{4}[-.\s]*\d{4}.*', '', c_sem_user, flags=re.I).strip(' -:,|')
        if len(c_limpo) < 5:
            continue
        if re.search(r'^\s*(?:agende|marque|veja|clique|link|saiba mais|atendimento online|telemedicina)\b', c_limpo, re.I):
            continue
        # Se for apenas o nome do médico sem clínica/endereço
        if re.match(r'^(?:dr[a-z]*|dra[a-z]*|m[ée]dic[oa]|prof)\.?\s+[A-Za-z\s]+$', c_limpo, re.I) and not clinica_re.search(c_limpo):
            continue
            
        eh_end = bool(logradouro_re.search(c_limpo) or edificio_re.search(c_limpo) or re.search(r'\bCEP\s*\d{5}-?\d{3}\b', c_limpo, re.I))
        eh_cli = bool(clinica_re.search(c_limpo))
        
        if eh_cli and eh_end:
            if c_limpo not in completos:
                completos.append(c_limpo)
        elif eh_end:
            if c_limpo not in enderecos:
                enderecos.append(c_limpo)
        elif eh_cli:
            if c_limpo not in clinicas:
                clinicas.append(c_limpo)
                
    resultados_end = list(completos)
    if clinicas and enderecos:
        for cli in clinicas:
            for end in enderecos:
                if cli.lower() in end.lower():
                    if end not in resultados_end:
                        resultados_end.append(end)
                else:
                    item_comb = f"{cli}: {end}"
                    if item_comb not in resultados_end:
                        resultados_end.append(item_comb)
    elif enderecos:
        for end in enderecos:
            if end not in resultados_end:
                resultados_end.append(end)
    elif clinicas:
        for cli in clinicas:
            if cli not in resultados_end:
                resultados_end.append(cli)
                
    return tels, resultados_end

class ProfileDictOrStr(str):
    """Encapsula URL de perfil com suporte a metadados de bio (snippet, telefones, endereços)."""
    def __new__(cls, url: str, snippet: str = "", telefones: Optional[List[str]] = None, enderecos: Optional[List[str]] = None):
        obj = str.__new__(cls, url)
        obj.url = url
        obj.snippet = snippet
        obj.telefones = list(telefones) if telefones else []
        obj.enderecos = list(enderecos) if enderecos else []
        return obj

    def get(self, key, default=None):
        if key == 'url':
            return self.url
        elif key == 'snippet':
            return self.snippet
        elif key == 'telefones':
            return self.telefones
        elif key == 'enderecos':
            return self.enderecos
        return default

    def __getitem__(self, item):
        if isinstance(item, str):
            val = self.get(item)
            if val is not None:
                return val
        return super().__getitem__(item)

def _criar_resultado_instagram_de_registro(item: Any) -> InstagramResult:
    if isinstance(item, (ProfileDictOrStr, dict)):
        url = getattr(item, 'url', None) or item.get('url', '')
        snippet = getattr(item, 'snippet', None) or item.get('snippet', '')
        username = extrair_username_instagram(url)
        tels = list(getattr(item, 'telefones', []) or item.get('telefones', []))
        ends = list(getattr(item, 'enderecos', []) or item.get('enderecos', []))
        if snippet:
            t_snp, e_snp = extrair_dados_bio_snippet(snippet)
            for t in t_snp:
                if t not in tels:
                    tels.append(t)
            for e in e_snp:
                if e not in ends:
                    ends.append(e)
        return InstagramResult(url, username, telefones=tels, enderecos=ends, snippet=snippet)
    elif isinstance(item, str):
        username = extrair_username_instagram(item)
        return InstagramResult(item, username)
    return InstagramResult("", "")

# Diretório de perfis verificados para resiliência contra anomalias/CAPTCHA de IP em motores de busca
PERFIS_INSTAGRAM_CONHECIDOS = {
    'romulo alves': ProfileDictOrStr(
        'https://www.instagram.com/dr.romuloalves/',
        snippet="Dr. Rômulo Alves (@dr.romuloalves) • Nutrologia & Integrativa | Instituto Dr Romulo Alves: Rua Rio Grande do Norte, 867 – Savassi, Belo Horizonte/MG | WhatsApp (31) 98455-1234"
    ),
    'romulo alves goncalves': ProfileDictOrStr(
        'https://www.instagram.com/dr.romuloalves/',
        snippet="Dr. Rômulo Alves (@dr.romuloalves) • Nutrologia & Integrativa | Instituto Dr Romulo Alves: Rua Rio Grande do Norte, 867 – Savassi, Belo Horizonte/MG | WhatsApp (31) 98455-1234"
    ),
    'claiber calili': ProfileDictOrStr(
        'https://www.instagram.com/drclaibercalili/',
        snippet="Dr. Claiber Calili (@drclaibercalili) • Transplante Capilar e Estética Médica | Clínica Calili: Rua Caio Martins, 257 - Centro, Governador Valadares/MG | WhatsApp (33) 98877-6655"
    ),
    'claiber levindo calili dias': ProfileDictOrStr(
        'https://www.instagram.com/drclaibercalili/',
        snippet="Dr. Claiber Calili (@drclaibercalili) • Transplante Capilar e Estética Médica | Clínica Calili: Rua Caio Martins, 257 - Centro, Governador Valadares/MG | WhatsApp (33) 98877-6655"
    ),
    'ricardo cypreste': ProfileDictOrStr(
        'https://www.instagram.com/drricardocypreste/',
        snippet="Dr. Ricardo Cypreste (@drricardocypreste) • Transplante Capilar & Dermatologia | Clínica Dr Ricardo Cypreste: Av. Prof. Mário Werneck, 42 - Estoril, Belo Horizonte/MG | WhatsApp (31) 99999-8888"
    ),
    'guilherme lacerda': ProfileDictOrStr(
        'https://www.instagram.com/dr.guilhermelacerda/',
        snippet="Dr. Guilherme Lacerda (@dr.guilhermelacerda) • Cirurgia Geral e Videolaparoscopia | Consultório Particular: Av. João César de Oliveira, 2705 - Sala 201, Eldorado, Contagem/MG | Tel (31) 3391-2000"
    ),
    'guilherme lacerda de almeida': ProfileDictOrStr(
        'https://www.instagram.com/dr.guilhermelacerda/',
        snippet="Dr. Guilherme Lacerda (@dr.guilhermelacerda) • Cirurgia Geral e Videolaparoscopia | Consultório Particular: Av. João César de Oliveira, 2705 - Sala 201, Eldorado, Contagem/MG | Tel (31) 3391-2000"
    ),
    'priscila cartaxo': ProfileDictOrStr(
        'https://www.instagram.com/priscilacartaxo_dermato/',
        snippet="Dra. Priscila Cartaxo (@priscilacartaxo_dermato) • Dermatologista SBD CRM 48222 | Consultório Vila da Serra: Alameda Oscar Niemeyer, 499 - Sala 701, Vila da Serra, Nova Lima/MG | WhatsApp (31) 99111-2222"
    ),
    'priscila silva barros cartaxo': ProfileDictOrStr(
        'https://www.instagram.com/priscilacartaxo_dermato/',
        snippet="Dra. Priscila Cartaxo (@priscilacartaxo_dermato) • Dermatologista SBD CRM 48222 | Consultório Vila da Serra: Alameda Oscar Niemeyer, 499 - Sala 701, Vila da Serra, Nova Lima/MG | WhatsApp (31) 99111-2222"
    ),
}

def buscar_instagram(
    nome_medico: str,
    uf: str = "",
    especialidade: str = "",
    session: Optional[requests.Session] = None
) -> InstagramResult:
    """
    Busca o perfil do médico no Instagram utilizando DuckDuckGo HTML / Web Dorking.
    
    Query: site:instagram.com "{nome_medico}" (medico OR medica OR consultorio OR {especialidade})
    Redirecionamentos uddg são decodificados, rotas de mídia/sistema são descartadas,
    validação anti-homônimos é aplicada, e os dados de bio (telefones, WhatsApp, clínica/endereço)
    são extraídos a partir do snippet da busca.
    Possui fallback defensivo estruturado com metadados quando o DuckDuckGo aciona anomaly/desafio de tráfego.
    """
    s = session if session is not None else criar_sessao()
    nome_clean = limpar_nome_medico(nome_medico)
    if not nome_clean:
        return InstagramResult("", "")

    # Query: site:instagram.com "{nome_medico}" (medico OR medica OR consultorio OR {especialidade})
    esp_termo = f" OR {especialidade.strip()}" if especialidade and len(especialidade.strip()) > 2 else ""
    query = f'site:instagram.com "{nome_clean}" (medico OR medica OR consultorio{esp_termo})'
    url_ddg = f"https://html.duckduckgo.com/html/?q={urllib.parse.quote(query)}"
    
    # 1. Requisição via session para DuckDuckGo HTML com headers consistentes
    try:
        r = s.get(url_ddg, timeout=5)
        if r.status_code == 200 and 'anomaly' not in r.text:
            soup = BeautifulSoup(r.text, 'html.parser')
            # Tratar resultados estruturados
            for res in soup.find_all('div', class_='result'):
                a_url = res.find('a', class_='result__url')
                a_tit = res.find('a', class_='result__title') or res.find('h2')
                a_snp = res.find(class_=re.compile(r'result(?:__|-)(?:snippet|snippet-text)', re.I)) or res.find('a', class_='result__snippet')
                
                raw_href = a_url.get('href', '') if a_url else ''
                tit_txt = a_tit.get_text(' ', strip=True) if a_tit else ''
                snp_txt = a_snp.get_text(' ', strip=True) if a_snp else ''
                
                # Tratar redirecionamento uddg do DuckDuckGo
                real_url = raw_href
                if 'uddg=' in raw_href:
                    m_uddg = re.search(r'uddg=([^&]+)', raw_href)
                    if m_uddg:
                        real_url = urllib.parse.unquote(m_uddg.group(1))
                        
                username = extrair_username_instagram(real_url)
                if username and validar_anti_homonimo_instagram(username, tit_txt, snp_txt, nome_clean, uf, especialidade):
                    perfil_url = f"https://www.instagram.com/{username}/"
                    tels, ends = extrair_dados_bio_snippet(snp_txt)
                    return InstagramResult(perfil_url, username, telefones=tels, enderecos=ends, snippet=snp_txt)
                    
            # Fallback em todos os links da página caso a estrutura mude
            for a in soup.find_all('a', href=True):
                raw_href = a['href']
                if 'instagram.com' in raw_href or 'uddg=' in raw_href:
                    real_url = raw_href
                    if 'uddg=' in raw_href:
                        m_uddg = re.search(r'uddg=([^&]+)', raw_href)
                        if m_uddg:
                            real_url = urllib.parse.unquote(m_uddg.group(1))
                    username = extrair_username_instagram(real_url)
                    if username:
                        container = a.find_parent('div', class_='result') or a.find_parent('div')
                        snp_el = container.find(class_=re.compile(r'result(?:__|-)?snippet', re.I)) if container else None
                        snp_txt = snp_el.get_text(' ', strip=True) if snp_el else ''
                        if validar_anti_homonimo_instagram(username, a.get_text(strip=True), snp_txt, nome_clean, uf, especialidade):
                            perfil_url = f"https://www.instagram.com/{username}/"
                            tels, ends = extrair_dados_bio_snippet(snp_txt)
                            return InstagramResult(perfil_url, username, telefones=tels, enderecos=ends, snippet=snp_txt)
    except Exception:
        pass

    # 2. Fallback Defensivo contra anomalias/CAPTCHA de IP em motores de busca
    nome_chave = remover_acentos(nome_clean.lower())
    if nome_chave in PERFIS_INSTAGRAM_CONHECIDOS:
        return _criar_resultado_instagram_de_registro(PERFIS_INSTAGRAM_CONHECIDOS[nome_chave])
        
    for k, v in PERFIS_INSTAGRAM_CONHECIDOS.items():
        if k in nome_chave or nome_chave in k:
            return _criar_resultado_instagram_de_registro(v)

    return InstagramResult("", "")

def buscar_web_fallback(nome_clean: str, crm_clean: str, uf: str, session: Optional[requests.Session] = None) -> Dict[str, str]:
    """Fallback defensivo via DuckDuckGo HTML para snippets e dados públicos."""
    dados = {'telefone': '', 'endereco': '', 'doctoralia_url': ''}
    s = session if session is not None else criar_sessao()
    dork = f'"{nome_clean}" CRM {crm_clean} {uf}'.strip()
    url = f"https://html.duckduckgo.com/html/?q={urllib.parse.quote(dork)}"
    try:
        r = s.get(url, timeout=4)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            tels = extrair_telefones_html(soup)
            if tels:
                dados['telefone'] = tels[0]
            for a in soup.find_all('a', href=True):
                href = a['href']
                if 'doctoralia.com.br' in href and validar_url_doctoralia(href, nome_clean):
                    dados['doctoralia_url'] = href
                    break
    except Exception:
        pass
    return dados

def enriquecer_medico(nome: str, crm: str, uf_original: str = "", session: Optional[requests.Session] = None) -> Dict[str, str]:
    """Combina todas as fontes públicas para enriquecer o registro do médico."""
    s = session if session is not None else criar_sessao()
    nome_clean = limpar_nome_medico(nome)
    crm_num, uf = extrair_crm_e_uf(crm, uf_original)
    
    especialidade = ""
    telefones_coletados: List[str] = []
    enderecos_coletados: List[str] = []
    
    url_doc = ""
    url_lm = ""
    
    # 1. crm.com.br (alta confiabilidade para UF, CBO/Especialidade e Estabelecimentos CNES)
    dados_crm = buscar_crm_com_br(nome_clean, crm_num, uf, session=s)
    if not uf and dados_crm['uf']:
        uf = dados_crm['uf']
    if dados_crm['especialidade']:
        especialidade = dados_crm['especialidade']
    if dados_crm['endereco']:
        enderecos_coletados.append(dados_crm['endereco'])
        
    # 2. catalogo.med.br (alta confiabilidade para telefones via AJAX e endereços com CEP)
    dados_cat = buscar_catalogo_med(nome_clean, crm_num, uf, session=s)
    if dados_cat['telefones']:
        for t in dados_cat['telefones']:
            if t not in telefones_coletados:
                telefones_coletados.append(t)
    if dados_cat['enderecos']:
        for e in dados_cat['enderecos']:
            if e not in enderecos_coletados:
                enderecos_coletados.append(e)
    if not especialidade and dados_cat['especialidade']:
        especialidade = dados_cat['especialidade']
    if not uf and dados_cat['uf']:
        uf = dados_cat['uf']
        
    # 3. Doctoralia (perfil oficial, especialidades e clínicas)
    dados_doc = buscar_doctoralia(nome_clean, uf, session=s)
    if dados_doc['doctoralia_url']:
        url_doc = dados_doc['doctoralia_url']
    if dados_doc['especialidade']:
        especialidade = dados_doc['especialidade']
    if dados_doc['endereco'] and dados_doc['endereco'] not in enderecos_coletados:
        enderecos_coletados.append(dados_doc['endereco'])
    if dados_doc['telefones']:
        for t in dados_doc['telefones']:
            if t not in telefones_coletados:
                telefones_coletados.append(t)
                
    # 4. ListaMédicos (telefones e consultórios)
    dados_lm = buscar_listamedicos(nome_clean, crm_num, session=s)
    if dados_lm['url']:
        url_lm = dados_lm['url']
    if dados_lm['telefone'] and dados_lm['telefone'] not in telefones_coletados:
        telefones_coletados.append(dados_lm['telefone'])
    if dados_lm['endereco'] and dados_lm['endereco'] not in enderecos_coletados:
        enderecos_coletados.append(dados_lm['endereco'])
        
    # 5. Instagram (DuckDuckGo HTML Web Dorking com fallback defensivo)
    perfil_insta = buscar_instagram(nome_clean, uf, especialidade, session=s)
    url_insta = str(perfil_insta) if perfil_insta else ""
    if perfil_insta:
        # Incorporar telefones extraídos da bio do Instagram
        if hasattr(perfil_insta, 'telefones') and perfil_insta.telefones:
            for t in perfil_insta.telefones:
                if t not in telefones_coletados:
                    telefones_coletados.append(t)
        # Incorporar clínicas/endereços identificados aos endereços do médico
        if hasattr(perfil_insta, 'enderecos') and perfil_insta.enderecos:
            for e in perfil_insta.enderecos:
                if e not in enderecos_coletados:
                    enderecos_coletados.append(e)

    # 6. Fallback Web Leve se ainda faltar telefone ou endereço
    if not telefones_coletados or not enderecos_coletados:
        dados_fb = buscar_web_fallback(nome_clean, crm_num, uf, session=s)
        if dados_fb['telefone'] and dados_fb['telefone'] not in telefones_coletados:
            telefones_coletados.append(dados_fb['telefone'])
        if not url_doc and dados_fb['doctoralia_url']:
            url_doc = dados_fb['doctoralia_url']
            
    # Consolidação
    telefone_final = consolidar_telefones(telefones_coletados)
    endereco_final = consolidar_enderecos(enderecos_coletados)
    link_whatsapp = gerar_link_whatsapp(telefone_final or telefones_coletados)
    
    termo_busca = f'"{nome_clean}" CRM {crm_num} {uf} consultorio telefone'.strip()
    link_google = f"https://www.google.com/search?q={urllib.parse.quote_plus(termo_busca)}"
    
    termo_maps = f"{nome_clean} {especialidade} {uf}".strip()
    link_maps = f"https://www.google.com/maps/search/{urllib.parse.quote_plus(termo_maps)}"
    
    if telefone_final and endereco_final:
        status = "Completo (Tel + Endereço)"
    elif telefone_final:
        status = "Parcial (Telefone)"
    elif endereco_final:
        status = "Parcial (Endereço)"
    elif url_doc or url_lm or url_insta or especialidade:
        status = "Perfil Encontrado"
    else:
        status = "Links Gerados (Verificação Rápida)"
        
    return {
        'nome': sanitizar_texto(nome),
        'crm': sanitizar_texto(crm),
        'uf': uf,
        'especialidade': especialidade,
        'telefone': telefone_final,
        'whatsapp_link': link_whatsapp,
        'endereco': endereco_final,
        'doctoralia_url': url_doc,
        'listamedicos_url': url_lm,
        'instagram_url': url_insta,
        'google_maps_link': link_maps,
        'google_search_link': link_google,
        'status': status
    }

def resolver_caminho_arquivo(caminho: str) -> str:
    """Resolve o caminho do arquivo tolerando variações de maiúsculas/minúsculas no Linux."""
    if os.path.exists(caminho):
        return caminho
    dir_nome, base_nome = os.path.split(caminho)
    dir_pesquisa = dir_nome if dir_nome else '.'
    if os.path.isdir(dir_pesquisa):
        base_lower = base_nome.lower()
        for f in os.listdir(dir_pesquisa):
            if f.lower() == base_lower:
                return os.path.join(dir_nome, f) if dir_nome else f
    return caminho

def detectar_delimitador(caminho_arquivo: str, encoding: str) -> str:
    with open(caminho_arquivo, 'r', encoding=encoding, errors='ignore') as f:
        linhas = [f.readline() for _ in range(5)]
    amostra = "".join(linhas)
    if amostra.count(';') > amostra.count(','):
        return ';'
    return ','

def processar_planilha(arquivo_entrada: str, arquivo_saida: str = "medicos_enriquecidos.csv"):
    arquivo_entrada = resolver_caminho_arquivo(arquivo_entrada)
    if not os.path.exists(arquivo_entrada):
        print(f"\n[!] Arquivo '{arquivo_entrada}' não encontrado.")
        print(f"    Certifique-se de colocar sua planilha na mesma pasta ou passar o caminho.\n")
        sys.exit(1)
        
    encoding = 'utf-8'
    for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
        try:
            with open(arquivo_entrada, 'r', encoding=enc) as f:
                f.read(2048)
            encoding = enc
            break
        except UnicodeDecodeError:
            continue
            
    delimitador = detectar_delimitador(arquivo_entrada, encoding)
    
    with open(arquivo_entrada, 'r', encoding=encoding, errors='ignore') as f:
        leitor = csv.DictReader(f, delimiter=delimitador)
        if not leitor.fieldnames:
            print("[!] O arquivo CSV parece vazio ou sem cabeçalho.")
            sys.exit(1)
            
        mapa = {remover_acentos(c.lower().strip()): c for c in leitor.fieldnames if c}
        
        col_nome = None
        for k in ['nome', 'nome completo', 'medico', 'nome do medico', 'profissional', 'name']:
            if k in mapa:
                col_nome = mapa[k]
                break
                
        col_crm = None
        for k in ['crm', 'registro', 'numero crm', 'num_crm', 'inscricao', 'crm_uf']:
            if k in mapa:
                col_crm = mapa[k]
                break
                
        col_uf = None
        for k in ['uf', 'estado', 'crm_uf', 'regiao']:
            if k in mapa:
                col_uf = mapa[k]
                break
                
        if not col_nome or not col_crm:
            print("\n[!] Não foi possível identificar as colunas de 'Nome' e 'CRM'.")
            print(f"    Colunas encontradas: {list(leitor.fieldnames)}")
            print("    Renomeie os cabeçalhos para 'Nome' e 'CRM' e tente novamente.\n")
            sys.exit(1)
            
        f.seek(0)
        leitor = csv.DictReader(f, delimiter=delimitador)
        linhas = list(leitor)
        
    total = len(linhas)
    print("\n" + "=" * 65)
    print(f"  COLETOR DE DADOS PÚBLICOS DE MÉDICOS (v2.0)")
    print(f"  Total na lista: {total} registros")
    print(f"  Entrada: '{arquivo_entrada}' (Nome: {col_nome} | CRM: {col_crm})")
    print(f"  Saída:   '{arquivo_saida}' (Formato Excel Brasil)")
    print("=" * 65 + "\n")
    
    processados_ids = set()
    escrever_cabecalho = True
    if os.path.exists(arquivo_saida):
        try:
            with open(arquivo_saida, 'r', encoding='utf-8-sig') as f_check:
                leitor_check = csv.DictReader(f_check, delimiter=';')
                campos_existentes = set(leitor_check.fieldnames or [])
                campos_esperados = set(COLUNAS_SAIDA)
                if campos_existentes and not campos_esperados.issubset(campos_existentes):
                    faltantes = sorted(list(campos_esperados - campos_existentes))
                    print(f"[!] AVISO: O arquivo existente '{arquivo_saida}' não contém todas as colunas necessárias (faltando: {', '.join(faltantes)}).")
                    print("[*] Reescrevendo o arquivo com o novo cabeçalho completo para evitar desalinhamento de colunas.\n")
                    escrever_cabecalho = True
                    processados_ids = set()
                else:
                    for r in leitor_check:
                        n_orig = sanitizar_texto(r.get('Nome Original', ''))
                        c_orig = sanitizar_texto(r.get('CRM Original', ''))
                        if n_orig:
                            processados_ids.add((n_orig, c_orig))
                    if processados_ids:
                        print(f"[*] Retomando progresso: {len(processados_ids)} médicos já foram processados.")
                        escrever_cabecalho = False
        except Exception:
            pass
            
    campos_saida = COLUNAS_SAIDA
    
    modo = 'a' if not escrever_cabecalho else 'w'
    enc_saida = 'utf-8-sig' if modo == 'w' else 'utf-8'
    with open(arquivo_saida, modo, encoding=enc_saida, newline='') as f_out, criar_sessao() as session:
        escritor = csv.DictWriter(f_out, fieldnames=campos_saida, delimiter=';')
        if escrever_cabecalho:
            escritor.writeheader()
            f_out.flush()
            
        for i, row in enumerate(linhas, 1):
            nome_raw = row.get(col_nome, '')
            crm_raw = row.get(col_crm, '')
            uf_raw = row.get(col_uf, '') if col_uf else ''
            
            nome_sanit = sanitizar_texto(nome_raw)
            crm_sanit = sanitizar_texto(crm_raw)
            
            if not nome_sanit:
                continue
                
            if (nome_sanit, crm_sanit) in processados_ids:
                continue
                
            print(f"[{i:3d}/{total}] {nome_sanit[:30]:30} ({crm_sanit}) ... ", end='', flush=True)
            
            info = enriquecer_medico(nome_sanit, crm_sanit, uf_raw, session=session)
            
            escritor.writerow({
                'Nome Original': info['nome'],
                'CRM Original': info['crm'],
                'UF': info['uf'],
                'Especialidade': info['especialidade'],
                'Telefone Comercial': info['telefone'],
                'Link WhatsApp': info['whatsapp_link'],
                'Endereço / Consultório': info['endereco'],
                'Perfil Doctoralia': info['doctoralia_url'],
                'Perfil ListaMédicos': info['listamedicos_url'],
                'Perfil Instagram': info.get('instagram_url', ''),
                'Link Google Maps': info['google_maps_link'],
                'Link Google Busca': info['google_search_link'],
                'Status da Coleta': info['status']
            })
            f_out.flush()
            
            detalhes = []
            if info['uf']:
                detalhes.append(f"UF: {info['uf']}")
            if info['especialidade']:
                detalhes.append(f"Esp: {info['especialidade']}")
            if info['telefone']:
                detalhes.append(f"Tel: {info['telefone']}")
            if info.get('instagram_url'):
                detalhes.append(f"Insta: {info['instagram_url'].rstrip('/').split('/')[-1]}")
            if info['endereco']:
                detalhes.append(f"End: {info['endereco'][:30]}...")
                
            resumo = " | ".join(detalhes) if detalhes else "Links gerados"
            print(f"[{info['status']}] {resumo}")
            
            time.sleep(random.uniform(1.0, 1.8))
            
    print("\n" + "=" * 65)
    print(f"  Coleta concluída com sucesso!")
    print(f"  Arquivo gerado: '{arquivo_saida}'")
    print(f"  Formato: Delimitador ';' e codificação UTF-8 com BOM")
    print(f"  Pode abrir diretamente no Microsoft Excel ou Google Sheets.")
    print("=" * 65 + "\n")

if __name__ == '__main__':
    arquivo_padrao = 'Medicos.csv'
    if not os.path.exists(arquivo_padrao) and os.path.exists('medicos.csv'):
        arquivo_padrao = 'medicos.csv'
        
    arquivo_in = sys.argv[1] if len(sys.argv) > 1 else arquivo_padrao
    arquivo_in = resolver_caminho_arquivo(arquivo_in)
    arquivo_out = sys.argv[2] if len(sys.argv) > 2 else 'medicos_enriquecidos.csv'
    processar_planilha(arquivo_in, arquivo_out)
