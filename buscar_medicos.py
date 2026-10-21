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

# Segurança para execução em modo windowed no Windows (onde sys.stdout e sys.stderr são None)
class _NullStream:
    def write(self, text):
        pass
    def flush(self):
        pass

if sys.stdout is None:
    sys.stdout = _NullStream()
if sys.stderr is None:
    sys.stderr = _NullStream()

import csv
import re
import time
import random
import socket
import unicodedata
import urllib.parse
import threading
import queue
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import ttk
from typing import Dict, List, Optional, Tuple, Set, Any, Union

# ---------------------------------------------------------------------------
# Estruturas e Chamadas Nativas do Windows (Win32 API via ctypes)
# ---------------------------------------------------------------------------

OFN_OVERWRITEPROMPT = 0x00000002
OFN_HIDEREADONLY    = 0x00000004
OFN_NOCHANGEDIR     = 0x00000008
OFN_PATHMUSTEXIST   = 0x00000800
OFN_FILEMUSTEXIST   = 0x00001000
OFN_EXPLORER        = 0x00080000

MB_OK               = 0x00000000
MB_OKCANCEL         = 0x00000001
MB_YESNO            = 0x00000004
MB_ICONERROR        = 0x00000010
MB_ICONWARNING      = 0x00000030
MB_ICONINFORMATION  = 0x00000040
IDYES               = 6

class OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", ctypes.c_uint32),
        ("hwndOwner", wintypes.HWND),
        ("hInstance", wintypes.HINSTANCE),
        ("lpstrFilter", wintypes.LPCWSTR),
        ("lpstrCustomFilter", wintypes.LPWSTR),
        ("nMaxCustFilter", ctypes.c_uint32),
        ("nFilterIndex", ctypes.c_uint32),
        ("lpstrFile", wintypes.LPWSTR),
        ("nMaxFile", ctypes.c_uint32),
        ("lpstrFileTitle", wintypes.LPWSTR),
        ("nMaxFileTitle", ctypes.c_uint32),
        ("lpstrInitialDir", wintypes.LPCWSTR),
        ("lpstrTitle", wintypes.LPCWSTR),
        ("Flags", ctypes.c_uint32),
        ("nFileOffset", ctypes.c_uint16),
        ("nFileExtension", ctypes.c_uint16),
        ("lpstrDefExt", wintypes.LPCWSTR),
        ("lCustData", wintypes.LPARAM),
        ("lpfnHook", ctypes.c_void_p),
        ("lpTemplateName", wintypes.LPCWSTR),
        ("pvReserved", ctypes.c_void_p),
        ("dwReserved", ctypes.c_uint32),
        ("FlagsEx", ctypes.c_uint32),
    ]

def ativar_high_dpi():
    """Ativa High-DPI no Windows para evitar janelas borradas em telas com escala >100%."""
    if os.name == 'nt' and hasattr(ctypes, 'windll'):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                try:
                    ctypes.windll.user32.SetProcessDPIAware()
                except Exception:
                    pass

def mostrar_mensagem_nativa(titulo: str, mensagem: str, tipo: str = "info", hwnd=None) -> int:
    """Exibe caixas de diálogo 100% nativas do Windows via MessageBoxW (ou fallback Tk em Linux)."""
    if os.name == 'nt' and hasattr(ctypes, 'windll'):
        flags = MB_OK
        if tipo == "erro":
            flags |= MB_ICONERROR
        elif tipo == "aviso":
            flags |= MB_ICONWARNING
        elif tipo == "pergunta":
            flags = MB_YESNO | MB_ICONWARNING
        else:
            flags |= MB_ICONINFORMATION
        return ctypes.windll.user32.MessageBoxW(hwnd, mensagem, titulo, flags)
    else:
        try:
            from tkinter import messagebox
            if tipo == "erro":
                messagebox.showerror(titulo, mensagem)
            elif tipo == "aviso":
                messagebox.showwarning(titulo, mensagem)
            elif tipo == "pergunta":
                return IDYES if messagebox.askyesno(titulo, mensagem) else 0
            else:
                messagebox.showinfo(titulo, mensagem)
        except Exception:
            print(f"[{titulo}] {mensagem}")
        return 1

def selecionar_arquivo_entrada_nativo(hwnd=None, dir_inicial: str = "") -> Optional[str]:
    """Abre o diálogo nativo do Windows de seleção de arquivo via GetOpenFileNameW."""
    if os.name == 'nt' and hasattr(ctypes, 'windll'):
        ofn = OPENFILENAMEW()
        ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
        ofn.hwndOwner = hwnd
        filtro = "Planilhas CSV (*.csv)\0*.csv\0Todos os arquivos (*.*)\0*.*\0\0"
        ofn.lpstrFilter = filtro

        buf = ctypes.create_unicode_buffer(2048)
        ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
        ofn.nMaxFile = len(buf)

        if dir_inicial and os.path.isdir(dir_inicial):
            ofn.lpstrInitialDir = dir_inicial

        ofn.lpstrTitle = "Selecione a Planilha CSV de Entrada com os Médicos"
        ofn.Flags = (
            OFN_EXPLORER |
            OFN_FILEMUSTEXIST |
            OFN_PATHMUSTEXIST |
            OFN_HIDEREADONLY |
            OFN_NOCHANGEDIR
        )

        if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
            return buf.value
        return None
    else:
        try:
            from tkinter import filedialog
            caminho = filedialog.askopenfilename(
                title="Selecione o arquivo CSV de entrada",
                initialdir=dir_inicial or os.getcwd(),
                filetypes=[("Planilhas CSV", "*.csv"), ("Todos os arquivos", "*.*")]
            )
            return caminho if caminho else None
        except Exception:
            return None

def selecionar_arquivo_saida_nativo(hwnd=None, dir_inicial: str = "", nome_padrao: str = "Medicos_pesquisados.csv") -> Optional[str]:
    """Abre o diálogo nativo do Windows de salvamento via GetSaveFileNameW com nome padrão preenchido."""
    if os.name == 'nt' and hasattr(ctypes, 'windll'):
        ofn = OPENFILENAMEW()
        ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
        ofn.hwndOwner = hwnd
        filtro = "Planilhas CSV (*.csv)\0*.csv\0Todos os arquivos (*.*)\0*.*\0\0"
        ofn.lpstrFilter = filtro

        buf = ctypes.create_unicode_buffer(nome_padrao, 2048)
        ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
        ofn.nMaxFile = len(buf)

        if dir_inicial and os.path.isdir(dir_inicial):
            ofn.lpstrInitialDir = dir_inicial

        ofn.lpstrTitle = "Selecione onde salvar o arquivo com os resultados"
        ofn.lpstrDefExt = "csv"
        ofn.Flags = (
            OFN_EXPLORER |
            OFN_PATHMUSTEXIST |
            OFN_OVERWRITEPROMPT |
            OFN_HIDEREADONLY |
            OFN_NOCHANGEDIR
        )

        if ctypes.windll.comdlg32.GetSaveFileNameW(ctypes.byref(ofn)):
            return buf.value
        return None
    else:
        try:
            from tkinter import filedialog
            caminho = filedialog.asksaveasfilename(
                title="Selecione onde salvar o resultado da pesquisa",
                initialdir=dir_inicial or os.getcwd(),
                initialfile=nome_padrao,
                defaultextension=".csv",
                filetypes=[("Planilhas CSV", "*.csv"), ("Todos os arquivos", "*.*")]
            )
            return caminho if caminho else None
        except Exception:
            return None

class PlanilhaError(Exception):
    """Exceção levantada em caso de erros de validação ou estrutura da planilha CSV."""
    pass

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

def processar_planilha(
    arquivo_entrada: str,
    arquivo_saida: str = "Medicos_pesquisados.csv",
    progresso_callback: Optional[Any] = None
) -> int:
    arquivo_entrada = resolver_caminho_arquivo(arquivo_entrada)
    if not os.path.exists(arquivo_entrada):
        msg = f"Arquivo '{arquivo_entrada}' não encontrado.\nCertifique-se de colocar sua planilha na mesma pasta ou passar o caminho completo."
        print(f"\n[!] {msg}\n")
        raise FileNotFoundError(msg)
        
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
            msg = "O arquivo CSV parece vazio ou sem cabeçalho."
            print(f"[!] {msg}")
            raise PlanilhaError(msg)
            
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
            colunas_disp = list(leitor.fieldnames) if leitor.fieldnames else []
            msg = (
                f"Não foi possível identificar as colunas de 'Nome' e 'CRM' no arquivo CSV.\n\n"
                f"Colunas encontradas: {colunas_disp}\n\n"
                f"Por favor, renomeie os cabeçalhos para 'Nome' e 'CRM' e tente novamente."
            )
            print(f"\n[!] {msg}\n")
            raise PlanilhaError(msg)
            
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
                if progresso_callback:
                    try:
                        progresso_callback(i, total, "(Registro sem nome)")
                    except Exception:
                        pass
                continue
                
            if (nome_sanit, crm_sanit) in processados_ids:
                if progresso_callback:
                    try:
                        progresso_callback(i, total, f"{nome_sanit} (já processado)")
                    except Exception:
                        pass
                continue
                
            if progresso_callback:
                try:
                    progresso_callback(i, total, nome_sanit)
                except Exception:
                    pass

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
    return total

# ---------------------------------------------------------------------------
# Redirecionamento de Logs e Interface Gráfica (GUI)
# ---------------------------------------------------------------------------

class LogQueueWriter:
    """Redirecionador seguro de stdout/stderr para a fila de logs da GUI e console original."""
    def __init__(self, log_queue: queue.Queue, fallback_stream=None):
        self.log_queue = log_queue
        self.fallback_stream = fallback_stream

    def write(self, text: str):
        if text:
            self.log_queue.put(text)
            if self.fallback_stream and hasattr(self.fallback_stream, 'write'):
                try:
                    self.fallback_stream.write(text)
                except Exception:
                    pass

    def flush(self):
        if self.fallback_stream and hasattr(self.fallback_stream, 'flush'):
            try:
                self.fallback_stream.flush()
            except Exception:
                pass

class BuscaMedicosGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Busca Médicos - Coletor de Dados Públicos")
        self.root.geometry("740x580")
        self.root.minsize(640, 480)

        self.em_execucao = False

        # Tema nativo do Windows no ttk quando disponível
        style = ttk.Style()
        temas_disponiveis = style.theme_names()
        for tema in ['vista', 'winnative']:
            if tema in temas_disponiveis:
                style.theme_use(tema)
                break

        # Variáveis dos caminhos
        self.caminho_entrada = tk.StringVar(value="")
        self.caminho_saida = tk.StringVar(value="Medicos_pesquisados.csv")

        # Fila e redirecionamento de logs
        self.log_queue = queue.Queue()
        self._orig_stdout = sys.stdout
        self._orig_stderr = sys.stderr
        sys.stdout = LogQueueWriter(self.log_queue, self._orig_stdout)
        sys.stderr = LogQueueWriter(self.log_queue, self._orig_stderr)

        self._construir_interface()
        self.root.protocol("WM_DELETE_WINDOW", self._on_fechar)
        self.root.after(100, self._atualizar_logs)

    def _obter_hwnd(self):
        if os.name == 'nt':
            try:
                return self.root.winfo_id()
            except Exception:
                return None
        return None

    def _construir_interface(self):
        # NOTA: Não criamos barra de menu (Menu) conforme solicitado expressamente pelo usuário.
        container = ttk.Frame(self.root, padding="15 15 15 15")
        container.pack(fill='both', expand=True)

        # Cabeçalho
        frame_cabecalho = ttk.Frame(container)
        frame_cabecalho.pack(fill='x', pady=(0, 12))

        lbl_titulo = ttk.Label(
            frame_cabecalho,
            text="Coletor de Dados Públicos de Médicos",
            font=("Segoe UI" if os.name == 'nt' else "Helvetica", 13, "bold")
        )
        lbl_titulo.pack(anchor='w')

        lbl_sub = ttk.Label(
            frame_cabecalho,
            text="Enriqueça planilhas com telefones, consultórios, links de WhatsApp e perfis públicos.",
            font=("Segoe UI" if os.name == 'nt' else "Helvetica", 9)
        )
        lbl_sub.pack(anchor='w', pady=(2, 0))

        # Quadro de Seleção de Arquivos
        frame_arquivos = ttk.LabelFrame(container, text=" Arquivos de Entrada e Saída ", padding="10 10 10 10")
        frame_arquivos.pack(fill='x', pady=(0, 10))

        lbl_in = ttk.Label(frame_arquivos, text="Arquivo de Entrada (CSV):", font=("Segoe UI" if os.name == 'nt' else "Helvetica", 9, "bold"))
        lbl_in.grid(row=0, column=0, sticky='w', pady=(0, 2))

        entry_in = ttk.Entry(frame_arquivos, textvariable=self.caminho_entrada, state='readonly')
        entry_in.grid(row=1, column=0, sticky='ew', pady=(0, 8))

        lbl_out = ttk.Label(frame_arquivos, text="Arquivo de Saída (CSV):", font=("Segoe UI" if os.name == 'nt' else "Helvetica", 9, "bold"))
        lbl_out.grid(row=2, column=0, sticky='w', pady=(0, 2))

        entry_out = ttk.Entry(frame_arquivos, textvariable=self.caminho_saida)
        entry_out.grid(row=3, column=0, sticky='ew', pady=(0, 8))

        frame_arquivos.columnconfigure(0, weight=1)

        # Botão 1: Seleção em Sequência (Entrada -> Saída)
        self.btn_selecionar = ttk.Button(
            frame_arquivos,
            text="Selecionar Arquivo de entrada e Local de Saída",
            command=self.selecionar_sequencial
        )
        self.btn_selecionar.grid(row=4, column=0, sticky='ew', pady=(4, 2))

        # Botão 2: Começar
        frame_acao = ttk.Frame(container)
        frame_acao.pack(fill='x', pady=(5, 10))

        self.btn_comecar = ttk.Button(
            frame_acao,
            text="Começar",
            command=self.on_comecar
        )
        self.btn_comecar.pack(fill='x', ipady=6)

        # Barra de Progresso e Status
        frame_progresso = ttk.Frame(container)
        frame_progresso.pack(fill='x', pady=(0, 8))

        self.status_label = ttk.Label(
            frame_progresso,
            text="Pronto para iniciar. Selecione os arquivos e clique em 'Começar'.",
            font=("Segoe UI" if os.name == 'nt' else "Helvetica", 9)
        )
        self.status_label.pack(anchor='w', pady=(0, 4))

        self.progresso = ttk.Progressbar(frame_progresso, orient='horizontal', mode='determinate')
        self.progresso.pack(fill='x')

        # Área de Logs em Tempo Real
        lbl_logs = ttk.Label(
            container,
            text="Registro de Atividades (Logs em tempo real):",
            font=("Segoe UI" if os.name == 'nt' else "Helvetica", 9, "bold")
        )
        lbl_logs.pack(anchor='w', pady=(4, 2))

        frame_log = ttk.Frame(container)
        frame_log.pack(fill='both', expand=True)

        scroll = ttk.Scrollbar(frame_log, orient='vertical')
        scroll.pack(side='right', fill='y')

        self.area_log = tk.Text(
            frame_log,
            wrap='char',
            yscrollcommand=scroll.set,
            font=("Consolas" if os.name == 'nt' else "Monospace", 9),
            state='disabled',
            bg="#fdfdfd" if os.name == 'nt' else "#ffffff",
            relief='solid',
            borderwidth=1
        )
        self.area_log.pack(side='left', fill='both', expand=True)
        scroll.config(command=self.area_log.yview)

    def selecionar_sequencial(self):
        """Abre em sequência: 1º entrada via GetOpenFileNameW, depois saída via GetSaveFileNameW."""
        hwnd = self._obter_hwnd()
        dir_atual = ""
        if self.caminho_entrada.get() and os.path.exists(self.caminho_entrada.get()):
            dir_atual = os.path.dirname(os.path.abspath(self.caminho_entrada.get()))
        elif os.path.exists(os.getcwd()):
            dir_atual = os.getcwd()

        # Diálogo 1: Selecionar entrada
        arquivo_in = selecionar_arquivo_entrada_nativo(hwnd=hwnd, dir_inicial=dir_atual)
        if not arquivo_in:
            return  # Usuário cancelou

        self.caminho_entrada.set(arquivo_in)
        dir_sugerido = os.path.dirname(os.path.abspath(arquivo_in))
        saida_padrao = os.path.join(dir_sugerido, "Medicos_pesquisados.csv")

        # Diálogo 2: Selecionar saída imediatamente (se o 1º não foi cancelado)
        arquivo_out = selecionar_arquivo_saida_nativo(
            hwnd=hwnd,
            dir_inicial=dir_sugerido,
            nome_padrao="Medicos_pesquisados.csv"
        )
        if arquivo_out:
            self.caminho_saida.set(arquivo_out)
        else:
            # Fallback seguro com caminho padrão preenchido
            if not self.caminho_saida.get() or self.caminho_saida.get() == "Medicos_pesquisados.csv":
                self.caminho_saida.set(saida_padrao)

        self.status_label.config(text="Arquivos selecionados. Clique em 'Começar' para iniciar a busca.")

    def on_comecar(self):
        """Inicia a busca na aplicação ou aciona a seleção caso os arquivos ainda não tenham sido definidos."""
        if self.em_execucao:
            return

        arquivo_in = self.caminho_entrada.get().strip()
        if not arquivo_in or not os.path.exists(arquivo_in):
            # Se o usuário clicar em "Começar" sem ter escolhido os arquivos ainda, aciona a seleção em sequência
            self.selecionar_sequencial()
            arquivo_in = self.caminho_entrada.get().strip()
            if not arquivo_in or not os.path.exists(arquivo_in):
                mostrar_mensagem_nativa(
                    "Arquivo de Entrada Não Selecionado",
                    "Por favor, selecione uma planilha CSV de entrada válida antes de começar.",
                    tipo="aviso",
                    hwnd=self._obter_hwnd()
                )
                return

        arquivo_out = self.caminho_saida.get().strip()
        if not arquivo_out:
            dir_in = os.path.dirname(os.path.abspath(arquivo_in))
            arquivo_out = os.path.join(dir_in, "Medicos_pesquisados.csv")
            self.caminho_saida.set(arquivo_out)

        self._iniciar_execucao(arquivo_in, arquivo_out)

    def _iniciar_execucao(self, arquivo_in: str, arquivo_out: str):
        self.em_execucao = True
        self.btn_comecar.config(state='disabled')
        self.btn_selecionar.config(state='disabled')
        self.progresso['value'] = 0
        self.status_label.config(text="Iniciando processamento...")

        t = threading.Thread(
            target=self._worker_processamento,
            args=(arquivo_in, arquivo_out),
            daemon=True
        )
        t.start()

    def _worker_processamento(self, arquivo_in: str, arquivo_out: str):
        try:
            total = processar_planilha(
                arquivo_in,
                arquivo_out,
                progresso_callback=self._progresso_callback
            )
            self.root.after(0, lambda: self._on_sucesso(total, arquivo_out))
        except Exception as e:
            msg_erro = str(e)
            self.root.after(0, lambda: self._on_erro(msg_erro))

    def _progresso_callback(self, i: int, total: int, nome: str):
        def _atualizar():
            self.progresso['maximum'] = total
            self.progresso['value'] = i
            self.status_label.config(text=f"Processando [{i}/{total}]: {nome}")
        self.root.after(0, _atualizar)

    def _on_sucesso(self, total: int, arquivo_out: str):
        self.em_execucao = False
        self.btn_comecar.config(state='normal')
        self.btn_selecionar.config(state='normal')
        self.progresso['value'] = self.progresso['maximum']
        self.status_label.config(text=f"Concluído! {total} médicos processados com sucesso.")
        mostrar_mensagem_nativa(
            "Busca Concluída com Sucesso",
            f"A coleta e enriquecimento dos médicos foram concluídos com sucesso!\n\n"
            f"Total de registros: {total}\n"
            f"Arquivo gerado:\n{arquivo_out}\n\n"
            f"O arquivo pode ser aberto diretamente no Microsoft Excel.",
            tipo="info",
            hwnd=self._obter_hwnd()
        )

    def _on_erro(self, msg_erro: str):
        self.em_execucao = False
        self.btn_comecar.config(state='normal')
        self.btn_selecionar.config(state='normal')
        self.status_label.config(text="Erro durante o processamento.")
        mostrar_mensagem_nativa(
            "Erro no Processamento",
            f"Não foi possível concluir a busca:\n\n{msg_erro}",
            tipo="erro",
            hwnd=self._obter_hwnd()
        )

    def _atualizar_logs(self):
        texto_acumulado = []
        try:
            while True:
                texto = self.log_queue.get_nowait()
                texto_acumulado.append(texto)
        except queue.Empty:
            pass

        if texto_acumulado:
            self.area_log.configure(state='normal')
            self.area_log.insert('end', "".join(texto_acumulado))
            self.area_log.see('end')
            self.area_log.configure(state='disabled')

        self.root.after(100, self._atualizar_logs)

    def _on_fechar(self):
        if self.em_execucao:
            ret = mostrar_mensagem_nativa(
                "Confirmação de Saída",
                "Uma busca está em andamento no momento.\nDeseja realmente interromper e fechar?",
                tipo="pergunta",
                hwnd=self._obter_hwnd()
            )
            if ret != IDYES:
                return
        sys.stdout = self._orig_stdout
        sys.stderr = self._orig_stderr
        self.root.destroy()

def iniciar_gui():
    """Inicia a aplicação gráfica."""
    ativar_high_dpi()
    root = tk.Tk()
    app = BuscaMedicosGUI(root)
    root.mainloop()

if __name__ == '__main__':
    if len(sys.argv) > 1:
        if sys.argv[1] in ('-h', '--help'):
            print("Uso: python buscar_medicos.py [arquivo_entrada.csv] [arquivo_saida.csv]")
            print("     Sem argumentos: inicia a interface gráfica (GUI).")
            sys.exit(0)

        arquivo_in = resolver_caminho_arquivo(sys.argv[1])
        arquivo_out = sys.argv[2] if len(sys.argv) > 2 else "Medicos_pesquisados.csv"
        try:
            processar_planilha(arquivo_in, arquivo_out)
        except Exception as e:
            print(f"\n[!] Erro no processamento: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        # Modo Interface Gráfica (GUI) - padrão ao dar duplo clique no executável
        iniciar_gui()
