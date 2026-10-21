# Coletor de Dados Públicos de Médicos

Script em Python para enriquecer automaticamente listas de médicos brasileiros a partir do **Nome Completo** e **CRM**.

---

## 🚀 Como Funciona

O script consulta fontes médicas públicas e abertas em tempo real:
1. **ConsultaCRM / CFM**: Identifica e valida a **UF do CRM** e a **Especialidade Médica** oficial.
2. **ListaMédicos**: Extrai o **endereço físico do consultório** e o **telefone comercial**.
3. **Doctoralia**: Localiza o perfil profissional, nomes de **clínicas onde atende** e links de agendamento.
4. **Instagram**: Localiza o perfil profissional verificado com checagem anti-homônimos e extrai dados de contato (telefones comerciais, celulares WhatsApp) e consultórios/endereços diretamente do snippet da bio pública no buscador.
5. **Gerador de Links Rápidos**:
   - Link direto para o **Google Maps** com o nome do médico, especialidade e cidade.
   - Link direto para o **Google Busca** com operadores prontos para conferência rápida em 1 clique.
   - Link direto para o **WhatsApp** (caso o telefone identificado seja celular com 9 dígitos).

---

## 📦 1. Instalação das Dependências

Abra o terminal (Linux, macOS ou Prompt de Comando / PowerShell no Windows) e execute:

```bash
pip install requests beautifulsoup4
```

---

## 📄 2. Como Preparar sua Planilha

Coloque seu arquivo de médicos na mesma pasta com o nome `medicos.csv` (ou qualquer outro nome).

O arquivo só precisa ter no mínimo duas colunas:
- **`Nome`** (ou `Nome Completo`, `Médico`)
- **`CRM`** (ou `Registro`, `CRM/UF`)
- *(Opcional)* **`UF`** (se não tiver, o script tentará descobrir a UF automaticamente)

Exemplo de formato (`medicos.csv`):
```csv
Nome;CRM;UF
Dr. João da Silva;12345;SP
Dra. Maria Santos;54321;RJ
Carlos Alberto;98765;MG
```

*(O script detecta automaticamente se sua planilha usa vírgula `,` ou ponto e vírgula `;`, e suporta acentuação brasileira em UTF-8 ou Latin-1).*

---

## ▶️ 3. Como Executar

Para rodar usando o arquivo padrão `medicos.csv`:
```bash
python buscar_medicos.py
```

Ou especificando o nome do seu arquivo:
```bash
python buscar_medicos.py minha_lista.csv resultado_final.csv
```

---

## 📊 4. Resultado Gerado (`medicos_enriquecidos.csv`)

O arquivo gerado é configurado especificamente para o Microsoft Excel e Google Sheets do Brasil:
- Codificação **UTF-8 com BOM** (não desconfigura acentos ou cedilha).
- Delimitador **ponto e vírgula (`;`)** (abre direto com colunas divididas no Excel).

### Colunas Geradas:
1. **`Nome Original`**: Nome enviado na lista.
2. **`CRM Original`**: Número do CRM.
3. **`UF`**: Estado do CRM (identificado ou fornecido).
4. **`Especialidade`**: Especialidade médica registrada (ex: Cardiologia, Neurocirurgia, Pediatria).
5. **`Telefone Comercial`**: Telefone do consultório formatado com DDD: `(XX) XXXX-XXXX`.
6. **`Link WhatsApp`**: Link direto `https://wa.me/...` (quando disponível celular).
7. **`Endereço / Consultório`**: Logradouro, bairro, clínica ou CEP.
8. **`Perfil Doctoralia`**: Link do perfil público no Doctoralia.
9. **`Perfil ListaMédicos`**: Link do diretório ListaMédicos.
10. **`Perfil Instagram`**: Link do perfil profissional público no Instagram (verificado).
11. **`Link Google Maps`**: Link pronto para abrir a clínica no Google Maps com 1 clique.
12. **`Link Google Busca`**: Link com busca exata para conferência em 1 clique.
13. **`Status da Coleta`**: `Completo`, `Parcial` ou `Links Gerados`.

---

## 🛡️ Segurança e Resiliência

- **Salvamento em Tempo Real**: Cada médico pesquisado é imediatamente salvo no arquivo. Se faltar luz ou a conexão cair, você não perde nada do que já foi feito.
- **Retomada Automática (Resume)**: Se você reiniciar o script, ele lê o arquivo de saída e pula automaticamente os médicos já pesquisados.
- **Pausa Automática**: Inclui intervalo de 1.5 a 2.5 segundos entre as consultas para respeitar os servidores e evitar bloqueios temporários.
