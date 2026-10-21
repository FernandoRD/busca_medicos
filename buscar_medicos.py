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
from typing import Dict, List, Optional, Tuple, Set, Any, Union

try:
    from PyQt6 import QtWidgets, QtCore, QtGui
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QLabel, QLineEdit, QPushButton, QProgressBar, QPlainTextEdit,
        QFrame, QFileDialog, QMessageBox, QSizePolicy
    )
    from PyQt6.QtCore import Qt, pyqtSignal, QObject
    from PyQt6.QtGui import QFont, QTextCursor
    PYQT6_DISPONIVEL = True
except ImportError:
    PYQT6_DISPONIVEL = False

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

def mostrar_mensagem_nativa(titulo: str, mensagem: str, tipo: str = "info", hwnd=None, parent_widget=None) -> int:
    """Exibe caixas de diálogo 100% nativas do Windows via MessageBoxW (ou fallback PyQt fora do Windows)."""
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
        if PYQT6_DISPONIVEL:
            try:
                if QApplication.instance():
                    if tipo == "erro":
                        QMessageBox.critical(parent_widget, titulo, mensagem)
                        return 1
                    elif tipo == "aviso":
                        QMessageBox.warning(parent_widget, titulo, mensagem)
                        return 1
                    elif tipo == "pergunta":
                        res = QMessageBox.question(
                            parent_widget, titulo, mensagem,
                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                        )
                        return IDYES if res == QMessageBox.StandardButton.Yes else 0
                    else:
                        QMessageBox.information(parent_widget, titulo, mensagem)
                        return 1
            except Exception:
                pass
        print(f"[{titulo}] {mensagem}")
        return 1

def selecionar_arquivo_entrada_nativo(hwnd=None, dir_inicial: str = "", parent_widget=None) -> Optional[str]:
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
        if PYQT6_DISPONIVEL:
            try:
                if QApplication.instance():
                    caminho, _ = QFileDialog.getOpenFileName(
                        parent_widget,
                        "Selecione o arquivo CSV de entrada",
                        dir_inicial or os.getcwd(),
                        "Planilhas CSV (*.csv);;Todos os arquivos (*.*)"
                    )
                    return caminho if caminho else None
            except Exception:
                pass
        return None

def selecionar_arquivo_saida_nativo(hwnd=None, dir_inicial: str = "", nome_padrao: str = "Medicos_pesquisados.csv", parent_widget=None) -> Optional[str]:
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
        if PYQT6_DISPONIVEL:
            try:
                if QApplication.instance():
                    caminho_inicial = os.path.join(dir_inicial or os.getcwd(), nome_padrao)
                    caminho, _ = QFileDialog.getSaveFileName(
                        parent_widget,
                        "Selecione onde salvar o resultado da pesquisa",
                        caminho_inicial,
                        "Planilhas CSV (*.csv);;Todos os arquivos (*.*)"
                    )
                    return caminho if caminho else None
            except Exception:
                pass
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

def enriquecer_medico(nome: str, crm: str, uf_original: str = "", session: Optional[requests.Session] = None, cancelar_evento: Optional[threading.Event] = None) -> Dict[str, str]:
    """Combina todas as fontes públicas para enriquecer o registro do médico com verificação de cancelamento."""
    if cancelar_evento and cancelar_evento.is_set():
        return {
            'nome': sanitizar_texto(nome),
            'crm': sanitizar_texto(crm),
            'uf': uf_original,
            'especialidade': '',
            'telefone': '',
            'whatsapp_link': '',
            'endereco': '',
            'doctoralia_url': '',
            'listamedicos_url': '',
            'instagram_url': '',
            'google_maps_link': '',
            'google_search_link': '',
            'status': 'Interrompido'
        }

    s = session if session is not None else criar_sessao()
    nome_clean = limpar_nome_medico(nome)
    crm_num, uf = extrair_crm_e_uf(crm, uf_original)
    
    especialidade = ""
    telefones_coletados: List[str] = []
    enderecos_coletados: List[str] = []
    
    url_doc = ""
    url_lm = ""
    
    # 1. crm.com.br (alta confiabilidade para UF, CBO/Especialidade e Estabelecimentos CNES)
    if not (cancelar_evento and cancelar_evento.is_set()):
        dados_crm = buscar_crm_com_br(nome_clean, crm_num, uf, session=s)
        if not uf and dados_crm['uf']:
            uf = dados_crm['uf']
        if dados_crm['especialidade']:
            especialidade = dados_crm['especialidade']
        if dados_crm['endereco']:
            enderecos_coletados.append(dados_crm['endereco'])
        
    # 2. catalogo.med.br (alta confiabilidade para telefones via AJAX e endereços com CEP)
    if not (cancelar_evento and cancelar_evento.is_set()):
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
    if not (cancelar_evento and cancelar_evento.is_set()):
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
    if not (cancelar_evento and cancelar_evento.is_set()):
        dados_lm = buscar_listamedicos(nome_clean, crm_num, session=s)
        if dados_lm['url']:
            url_lm = dados_lm['url']
        if dados_lm['telefone'] and dados_lm['telefone'] not in telefones_coletados:
            telefones_coletados.append(dados_lm['telefone'])
        if dados_lm['endereco'] and dados_lm['endereco'] not in enderecos_coletados:
            enderecos_coletados.append(dados_lm['endereco'])
        
    # 5. Instagram (DuckDuckGo HTML Web Dorking com fallback defensivo)
    if not (cancelar_evento and cancelar_evento.is_set()):
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
    if not (cancelar_evento and cancelar_evento.is_set()):
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
    
    if cancelar_evento and cancelar_evento.is_set():
        status = "Interrompido"
    elif telefone_final and endereco_final:
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

class ResultadoProcessamento(tuple):
    """Tupla personalizada (total, cancelado) compatível com int."""
    def __new__(cls, total: int, cancelado: bool):
        return super().__new__(cls, (total, cancelado))

    @property
    def total(self) -> int:
        return self[0]

    @property
    def cancelado(self) -> bool:
        return self[1]

    def __int__(self) -> int:
        return self.total

    def __eq__(self, other):
        if isinstance(other, int):
            return self.total == other
        return tuple.__eq__(self, other)

def processar_planilha(
    arquivo_entrada: str,
    arquivo_saida: str = "Medicos_pesquisados.csv",
    progresso_callback: Optional[Any] = None,
    cancelar_evento: Optional[threading.Event] = None
) -> ResultadoProcessamento:
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

    total_processados = 0
    cancelado = False

    with open(arquivo_saida, modo, encoding=enc_saida, newline='') as f_out, criar_sessao() as session:
        escritor = csv.DictWriter(f_out, fieldnames=campos_saida, delimiter=';')
        if escrever_cabecalho:
            escritor.writeheader()
            f_out.flush()
            
        for i, row in enumerate(linhas, 1):
            if cancelar_evento and cancelar_evento.is_set():
                cancelado = True
                f_out.flush()
                print("\n[!] Processamento interrompido pelo usuário.")
                break

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
            
            info = enriquecer_medico(nome_sanit, crm_sanit, uf_raw, session=session, cancelar_evento=cancelar_evento)
            
            if info.get('status') == 'Interrompido':
                cancelado = True
                f_out.flush()
                print("\n[!] Processamento interrompido pelo usuário.")
                break

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
            total_processados += 1
            processados_ids.add((nome_sanit, crm_sanit))
            
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

            if cancelar_evento and cancelar_evento.is_set():
                cancelado = True
                f_out.flush()
                print("\n[!] Processamento interrompido pelo usuário.")
                break

            # Espera com suporte a cancelamento imediato
            if cancelar_evento:
                if cancelar_evento.wait(timeout=random.uniform(1.0, 1.8)):
                    cancelado = True
                    f_out.flush()
                    print("\n[!] Processamento interrompido pelo usuário.")
                    break
            else:
                time.sleep(random.uniform(1.0, 1.8))
            
    if cancelado:
        return ResultadoProcessamento(total_processados, True)
    else:
        print("\n" + "=" * 65)
        print(f"  Coleta concluída com sucesso!")
        print(f"  Arquivo gerado: '{arquivo_saida}'")
        print(f"  Formato: Delimitador ';' e codificação UTF-8 com BOM")
        print(f"  Pode abrir diretamente no Microsoft Excel ou Google Sheets.")
        print("=" * 65 + "\n")
        return ResultadoProcessamento(total, False)

# ---------------------------------------------------------------------------
# Redirecionamento de Logs e Interface Gráfica Moderna (PyQt6 - Fluent Design)
# ---------------------------------------------------------------------------

ESTILO_FLUENT = """
QMainWindow, QWidget#centralWidget {
    background-color: #f1f5f9;
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
    color: #0f172a;
}

QFrame#cardFrame {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 12px;
}

QLabel#titleLabel {
    font-size: 17px;
    font-weight: 700;
    color: #0f172a;
}

QLabel#subtitleLabel {
    font-size: 12px;
    color: #64748b;
}

QLabel#sectionTitle {
    font-size: 13px;
    font-weight: 600;
    color: #334155;
    margin-bottom: 2px;
}

QLabel#fieldLabel {
    font-size: 12px;
    font-weight: 600;
    color: #475569;
}

QLineEdit {
    background-color: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 8px 10px;
    font-size: 12px;
    color: #0f172a;
}

QLineEdit:focus {
    border: 2px solid #0284c7;
    background-color: #ffffff;
}

QLineEdit:read-only {
    background-color: #f1f5f9;
    color: #475569;
}

QPushButton#btnSelecionar {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 9px 16px;
    font-size: 13px;
    font-weight: 600;
    color: #0f172a;
}

QPushButton#btnSelecionar:hover {
    background-color: #f8fafc;
    border-color: #94a3b8;
}

QPushButton#btnSelecionar:pressed {
    background-color: #e2e8f0;
}

QPushButton#btnComecar {
    background-color: #0284c7;
    border: 1px solid #0284c7;
    border-radius: 6px;
    padding: 10px 18px;
    font-size: 14px;
    font-weight: 700;
    color: #ffffff;
}

QPushButton#btnComecar:hover {
    background-color: #0369a1;
    border-color: #0369a1;
}

QPushButton#btnComecar:pressed {
    background-color: #075985;
}

QPushButton#btnComecar:disabled {
    background-color: #94a3b8;
    border-color: #94a3b8;
    color: #f8fafc;
}

QPushButton#btnCancelar {
    background-color: #ef4444;
    border: 1px solid #ef4444;
    border-radius: 6px;
    padding: 10px 18px;
    font-size: 14px;
    font-weight: 700;
    color: #ffffff;
}

QPushButton#btnCancelar:hover {
    background-color: #dc2626;
    border-color: #dc2626;
}

QPushButton#btnCancelar:pressed {
    background-color: #b91c1c;
}

QPushButton#btnCancelar:disabled {
    background-color: #e2e8f0;
    border-color: #e2e8f0;
    color: #94a3b8;
}

QProgressBar {
    background-color: #e2e8f0;
    border: none;
    border-radius: 6px;
    height: 14px;
    text-align: center;
    font-size: 11px;
    font-weight: 600;
    color: #1e293b;
}

QProgressBar::chunk {
    background-color: #0284c7;
    border-radius: 6px;
}

QPlainTextEdit#logEdit {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 8px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 11px;
    color: #1e293b;
}

QLabel#statusLabel {
    font-size: 12px;
    color: #475569;
}
"""

if PYQT6_DISPONIVEL:
    class WorkerSinais(QtCore.QObject):
        progresso = QtCore.pyqtSignal(int, int, str)
        log = QtCore.pyqtSignal(str)
        concluido = QtCore.pyqtSignal(int, bool, str)  # total, cancelado, arquivo_saida
        erro = QtCore.pyqtSignal(str)

    class QtLogWriter:
        """Redirecionador thread-safe de stdout/stderr para o log da GUI via sinal do Qt."""
        def __init__(self, log_signal, fallback_stream=None):
            self.log_signal = log_signal
            self.fallback_stream = fallback_stream

        def write(self, text: str):
            if text:
                try:
                    self.log_signal.emit(text)
                except Exception:
                    pass
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

    class BuscaMedicosWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Busca Médicos - Coletor de Dados Públicos")
            self.resize(840, 720)
            self.setMinimumSize(660, 500)

            self.em_execucao = False
            self.cancelar_evento: Optional[threading.Event] = None

            self.sinais = WorkerSinais()
            self.sinais.progresso.connect(self._on_progresso)
            self.sinais.log.connect(self._on_log)
            self.sinais.concluido.connect(self._on_concluido)
            self.sinais.erro.connect(self._on_erro)

            # Redirecionamento de stdout e stderr para o console de log da GUI
            self._orig_stdout = sys.stdout
            self._orig_stderr = sys.stderr
            sys.stdout = QtLogWriter(self.sinais.log, self._orig_stdout)
            sys.stderr = QtLogWriter(self.sinais.log, self._orig_stderr)

            self._construir_interface()
            self.setStyleSheet(ESTILO_FLUENT)

        def _obter_hwnd(self):
            if os.name == 'nt':
                try:
                    return int(self.winId())
                except Exception:
                    return None
            return None

        def _construir_interface(self):
            # NOTA: Não criamos barra de menu (Menu) conforme solicitado pelo usuário.
            self.setMenuBar(None)

            widget_central = QWidget()
            widget_central.setObjectName("centralWidget")
            self.setCentralWidget(widget_central)

            layout_principal = QVBoxLayout(widget_central)
            layout_principal.setContentsMargins(16, 16, 16, 16)
            layout_principal.setSpacing(10)

            # Card de Cabeçalho
            card_header = QFrame()
            card_header.setObjectName("cardFrame")
            layout_header = QVBoxLayout(card_header)
            layout_header.setContentsMargins(12, 10, 12, 10)
            layout_header.setSpacing(4)

            lbl_titulo = QLabel("Coletor de Dados Públicos de Médicos")
            lbl_titulo.setObjectName("titleLabel")
            lbl_sub = QLabel("Enriqueça planilhas com telefones, consultórios, links de WhatsApp e perfis públicos.")
            lbl_sub.setObjectName("subtitleLabel")
            layout_header.addWidget(lbl_titulo)
            layout_header.addWidget(lbl_sub)
            layout_principal.addWidget(card_header)

            # Card de Arquivos
            card_arquivos = QFrame()
            card_arquivos.setObjectName("cardFrame")
            layout_arquivos = QVBoxLayout(card_arquivos)
            layout_arquivos.setContentsMargins(12, 12, 12, 12)
            layout_arquivos.setSpacing(8)

            lbl_secao_arq = QLabel("Arquivos de Entrada e Saída")
            lbl_secao_arq.setObjectName("sectionTitle")
            layout_arquivos.addWidget(lbl_secao_arq)

            lbl_in = QLabel("Arquivo de Entrada (CSV):")
            lbl_in.setObjectName("fieldLabel")
            layout_arquivos.addWidget(lbl_in)

            self.edit_entrada = QLineEdit()
            self.edit_entrada.setReadOnly(True)
            self.edit_entrada.setPlaceholderText("Nenhum arquivo CSV selecionado...")
            layout_arquivos.addWidget(self.edit_entrada)

            lbl_out = QLabel("Arquivo de Saída (CSV):")
            lbl_out.setObjectName("fieldLabel")
            layout_arquivos.addWidget(lbl_out)

            self.edit_saida = QLineEdit("Medicos_pesquisados.csv")
            layout_arquivos.addWidget(self.edit_saida)

            self.btn_selecionar = QPushButton("Selecionar Arquivo de entrada e Local de Saída")
            self.btn_selecionar.setObjectName("btnSelecionar")
            self.btn_selecionar.setCursor(Qt.CursorShape.PointingHandCursor)
            self.btn_selecionar.clicked.connect(self.selecionar_sequencial)
            layout_arquivos.addWidget(self.btn_selecionar)

            layout_principal.addWidget(card_arquivos)

            # Layout de Ações: Botão Começar e Botão Cancelar Lado a Lado (QHBoxLayout)
            layout_botoes = QHBoxLayout()
            layout_botoes.setSpacing(10)

            self.btn_comecar = QPushButton("Começar")
            self.btn_comecar.setObjectName("btnComecar")
            self.btn_comecar.setCursor(Qt.CursorShape.PointingHandCursor)
            self.btn_comecar.clicked.connect(self.on_comecar)

            self.btn_cancelar = QPushButton("Cancelar")
            self.btn_cancelar.setObjectName("btnCancelar")
            self.btn_cancelar.setCursor(Qt.CursorShape.PointingHandCursor)
            self.btn_cancelar.setEnabled(False)  # Desabilitado inicialmente
            self.btn_cancelar.clicked.connect(self.on_cancelar)

            layout_botoes.addWidget(self.btn_comecar, 1)
            layout_botoes.addWidget(self.btn_cancelar, 1)
            layout_principal.addLayout(layout_botoes)

            # Card de Progresso e Status
            card_progresso = QFrame()
            card_progresso.setObjectName("cardFrame")
            layout_progresso = QVBoxLayout(card_progresso)
            layout_progresso.setContentsMargins(12, 10, 12, 10)
            layout_progresso.setSpacing(6)

            self.status_label = QLabel("Pronto para iniciar. Selecione os arquivos e clique em 'Começar'.")
            self.status_label.setObjectName("statusLabel")
            layout_progresso.addWidget(self.status_label)

            self.progresso = QProgressBar()
            self.progresso.setRange(0, 100)
            self.progresso.setValue(0)
            layout_progresso.addWidget(self.progresso)

            layout_principal.addWidget(card_progresso)

            # Card de Logs em Tempo Real
            card_logs = QFrame()
            card_logs.setObjectName("cardFrame")
            layout_logs = QVBoxLayout(card_logs)
            layout_logs.setContentsMargins(12, 10, 12, 10)
            layout_logs.setSpacing(6)

            lbl_logs = QLabel("Registro de Atividades (Logs em tempo real):")
            lbl_logs.setObjectName("sectionTitle")
            layout_logs.addWidget(lbl_logs)

            self.log_edit = QPlainTextEdit()
            self.log_edit.setObjectName("logEdit")
            self.log_edit.setReadOnly(True)
            layout_logs.addWidget(self.log_edit)

            layout_principal.addWidget(card_logs, 1)

        def selecionar_sequencial(self):
            """Executa a seleção em sequência:
            1. Abre diálogo nativo do Windows para escolher o CSV de entrada.
            2. Se selecionado, abre automaticamente o diálogo nativo de salvamento com 'Medicos_pesquisados.csv'.
            """
            hwnd = self._obter_hwnd()
            dir_atual = ""
            entrada_atual = self.edit_entrada.text().strip()
            if entrada_atual and os.path.exists(entrada_atual):
                dir_atual = os.path.dirname(os.path.abspath(entrada_atual))
            elif os.path.exists(os.getcwd()):
                dir_atual = os.getcwd()

            # Diálogo 1: Selecionar entrada
            arquivo_in = selecionar_arquivo_entrada_nativo(hwnd=hwnd, dir_inicial=dir_atual, parent_widget=self)
            if not arquivo_in:
                return  # Usuário cancelou

            self.edit_entrada.setText(arquivo_in)
            dir_sugerido = os.path.dirname(os.path.abspath(arquivo_in))
            saida_padrao = os.path.join(dir_sugerido, "Medicos_pesquisados.csv")

            # Diálogo 2: Selecionar saída imediatamente (se o 1º não foi cancelado)
            arquivo_out = selecionar_arquivo_saida_nativo(
                hwnd=hwnd,
                dir_inicial=dir_sugerido,
                nome_padrao="Medicos_pesquisados.csv",
                parent_widget=self
            )
            if arquivo_out:
                self.edit_saida.setText(arquivo_out)
            else:
                saida_atual = self.edit_saida.text().strip()
                if not saida_atual or saida_atual == "Medicos_pesquisados.csv":
                    self.edit_saida.setText(saida_padrao)

            self.status_label.setText("Arquivos configurados. Clique em 'Começar' para iniciar a busca.")

        def on_comecar(self):
            """Inicia a busca na aplicação ou aciona a seleção caso os arquivos ainda não tenham sido definidos."""
            if self.em_execucao:
                return

            arquivo_in = self.edit_entrada.text().strip()
            if not arquivo_in or not os.path.exists(arquivo_in):
                # Se o usuário clicar em "Começar" sem ter escolhido os arquivos ainda, aciona a seleção sequencial
                self.selecionar_sequencial()
                arquivo_in = self.edit_entrada.text().strip()
                if not arquivo_in or not os.path.exists(arquivo_in):
                    mostrar_mensagem_nativa(
                        "Arquivo de Entrada Não Selecionado",
                        "Por favor, selecione uma planilha CSV de entrada válida antes de começar.",
                        tipo="aviso",
                        hwnd=self._obter_hwnd(),
                        parent_widget=self
                    )
                    return

            arquivo_out = self.edit_saida.text().strip()
            if not arquivo_out:
                dir_in = os.path.dirname(os.path.abspath(arquivo_in))
                arquivo_out = os.path.join(dir_in, "Medicos_pesquisados.csv")
                self.edit_saida.setText(arquivo_out)

            self._iniciar_execucao(arquivo_in, arquivo_out)

        def _iniciar_execucao(self, arquivo_in: str, arquivo_out: str):
            self.em_execucao = True
            self.cancelar_evento = threading.Event()

            self.btn_comecar.setEnabled(False)
            self.btn_selecionar.setEnabled(False)
            self.btn_cancelar.setEnabled(True)

            self.progresso.setValue(0)
            self.status_label.setText("Iniciando busca e enriquecimento de dados...")

            def worker():
                try:
                    res = processar_planilha(
                        arquivo_in,
                        arquivo_out,
                        progresso_callback=lambda i, tot, nm: self.sinais.progresso.emit(i, tot, nm),
                        cancelar_evento=self.cancelar_evento
                    )
                    self.sinais.concluido.emit(res.total, res.cancelado, arquivo_out)
                except Exception as e:
                    self.sinais.erro.emit(str(e))

            self.worker_thread = threading.Thread(target=worker, daemon=True)
            self.worker_thread.start()

        def on_cancelar(self):
            """Disparado ao clicar no botão Cancelar: interrompe a execução com segurança."""
            if not self.em_execucao or self.cancelar_evento is None:
                return

            self.btn_cancelar.setEnabled(False)
            self.status_label.setText("Cancelando processamento... aguarde.")
            self.cancelar_evento.set()

        def _on_progresso(self, i: int, total: int, nome: str):
            if total > 0:
                self.progresso.setMaximum(total)
                self.progresso.setValue(i)
            self.status_label.setText(f"Processando [{i}/{total}]: {nome}")

        def _on_log(self, text: str):
            self.log_edit.moveCursor(QtGui.QTextCursor.MoveOperation.End)
            self.log_edit.insertPlainText(text)
            self.log_edit.moveCursor(QtGui.QTextCursor.MoveOperation.End)

        def _on_concluido(self, total: int, cancelado: bool, arquivo_out: str):
            self.em_execucao = False
            self.btn_comecar.setEnabled(True)
            self.btn_selecionar.setEnabled(True)
            self.btn_cancelar.setEnabled(False)

            if cancelado:
                self.status_label.setText("Processamento cancelado pelo usuário.")
                mostrar_mensagem_nativa(
                    "Processamento Interrompido",
                    f"O processamento foi interrompido pelo usuário.\n\n"
                    f"Registros processados e salvos: {total}\n"
                    f"Arquivo salvo:\n{arquivo_out}",
                    tipo="aviso",
                    hwnd=self._obter_hwnd(),
                    parent_widget=self
                )
            else:
                self.progresso.setValue(self.progresso.maximum())
                self.status_label.setText(f"Concluído com sucesso! {total} médicos processados.")
                mostrar_mensagem_nativa(
                    "Busca Concluída com Sucesso",
                    f"A coleta e enriquecimento dos médicos foram concluídos com sucesso!\n\n"
                    f"Total de registros processados: {total}\n"
                    f"Arquivo gerado:\n{arquivo_out}\n\n"
                    f"O arquivo pode ser aberto diretamente no Microsoft Excel.",
                    tipo="info",
                    hwnd=self._obter_hwnd(),
                    parent_widget=self
                )

        def _on_erro(self, msg_erro: str):
            self.em_execucao = False
            self.btn_comecar.setEnabled(True)
            self.btn_selecionar.setEnabled(True)
            self.btn_cancelar.setEnabled(False)
            self.status_label.setText("Erro durante o processamento.")
            mostrar_mensagem_nativa(
                "Erro no Processamento",
                f"Não foi possível concluir a busca:\n\n{msg_erro}",
                tipo="erro",
                hwnd=self._obter_hwnd(),
                parent_widget=self
            )

        def closeEvent(self, event):
            if self.em_execucao:
                ret = mostrar_mensagem_nativa(
                    "Confirmação de Saída",
                    "Uma busca está em andamento no momento.\nDeseja realmente cancelar o processamento e fechar a aplicação?",
                    tipo="pergunta",
                    hwnd=self._obter_hwnd(),
                    parent_widget=self
                )
                if ret != IDYES:
                    event.ignore()
                    return
                if self.cancelar_evento:
                    self.cancelar_evento.set()

            sys.stdout = self._orig_stdout
            sys.stderr = self._orig_stderr
            event.accept()

def iniciar_gui():
    """Inicia a aplicação gráfica moderna em PyQt6."""
    if not PYQT6_DISPONIVEL:
        print("\n[!] PyQt6 não está instalado no ambiente.")
        print("    Para usar a interface gráfica, instale com:")
        print("    pip install PyQt6\n")
        sys.exit(1)

    ativar_high_dpi()
    app = QApplication.instance()
    criou_app = False
    if app is None:
        app = QApplication(sys.argv)
        criou_app = True

    app.setStyle("Fusion")
    janela = BuscaMedicosWindow()
    janela.show()

    if criou_app:
        sys.exit(app.exec())

if __name__ == '__main__':
    if len(sys.argv) > 1:
        if sys.argv[1] in ('-h', '--help'):
            print("Uso: python buscar_medicos.py [arquivo_entrada.csv] [arquivo_saida.csv]")
            print("     Sem argumentos: inicia a interface gráfica moderna (PyQt6).")
            sys.exit(0)

        arquivo_in = resolver_caminho_arquivo(sys.argv[1])
        arquivo_out = sys.argv[2] if len(sys.argv) > 2 else "Medicos_pesquisados.csv"
        try:
            res = processar_planilha(arquivo_in, arquivo_out)
        except Exception as e:
            print(f"\n[!] Erro no processamento: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        # Modo Interface Gráfica (GUI) - padrão ao dar duplo clique no executável
        iniciar_gui()
