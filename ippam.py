import datetime
import hashlib
import io
import ipaddress
import json
import os
import sqlite3
import urllib.request
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# ==========================================
# CONFIGURAÇÃO INICIAL
# ==========================================
st.set_page_config(
    page_title="IPAM Enterprise - Gestão de IPs, ASNs e Engenharia",
    layout="wide",
    page_icon="🌐",
)

DB_FILE = "ipam_database.db"
EXCEL_FILE = "ipam_database.xlsx"

COLUNAS_ASN = ["ASN", "Entidade", "Tipo", "Contato"]
COLUNAS_IP = [
    "Prefixo",
    "Versão",
    "ASN Vinculado",
    "Finalidade",
    "VLAN ID",
    "VRF",
    "POP / Localidade",
    "ID Circuito",
]


# ==========================================
# FUNÇÕES DE SEGURANÇA E HASH DE SENHA
# ==========================================
def make_hashes(password):
  """Gera o hash SHA-256 da senha."""
  return hashlib.sha256(str.encode(password)).hexdigest()


def check_hashes(password, hashed_text):
  """Verifica se a senha digitada corresponde ao hash armazenado."""
  if make_hashes(password) == hashed_text:
    return True
  return False


# ==========================================
# GESTÃO DE BANCO DE DADOS (SQLITE)
# ==========================================
def init_db():
  """Inicializa e atualiza a estrutura do banco de dados SQLite."""
  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS asns (
            asn TEXT PRIMARY KEY,
            entidade TEXT,
            tipo TEXT,
            contato TEXT
        )
    """)

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS blocos_ip (
            prefixo TEXT PRIMARY KEY,
            versao TEXT,
            asn_vinculado TEXT,
            finalidade TEXT,
            vlan_id TEXT,
            vrf TEXT,
            pop_localidade TEXT,
            id_circuito TEXT
        )
    """)

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TEXT,
            usuario TEXT,
            acao TEXT,
            tabela TEXT,
            detalhes TEXT
        )
    """)

  # Migração da coluna de usuário nos logs
  try:
    cursor.execute("ALTER TABLE audit_logs ADD COLUMN usuario TEXT")
  except sqlite3.OperationalError:
    pass

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT,
            role TEXT
        )
    """)

  # Criar usuário admin padrão se a tabela estiver vazia
  cursor.execute("SELECT COUNT(*) FROM users")
  if cursor.fetchone()[0] == 0:
    cursor.execute(
        "INSERT INTO users VALUES (?, ?, ?)",
        ("admin", make_hashes("admin"), "admin"),
    )

  conn.commit()
  conn.close()


def registrar_log(acao, tabela, detalhes):
  """Registra alterações na tabela de audit_logs."""
  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()
  agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  usuario_atual = st.session_state.get("username", "Sistema")
  cursor.execute(
      "INSERT INTO audit_logs (data_hora, usuario, acao, tabela, detalhes)"
      " VALUES (?, ?, ?, ?, ?)",
      (agora, usuario_atual, acao, tabela, detalhes),
  )
  conn.commit()
  conn.close()


def autenticar_usuario(username, password):
  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT password, role FROM users WHERE username=?", (username.strip(),)
  )
  result = cursor.fetchone()
  conn.close()

  if result and check_hashes(password, result[0]):
    return result[1]  # Retorna a role ('admin' ou 'operator')
  return None


def carregar_dados_asn():
  conn = sqlite3.connect(DB_FILE)
  df = pd.read_sql_query("SELECT * FROM asns", conn)
  conn.close()
  df.columns = COLUNAS_ASN
  return df


def carregar_dados_ip():
  conn = sqlite3.connect(DB_FILE)
  df = pd.read_sql_query("SELECT * FROM blocos_ip", conn)
  conn.close()
  df.columns = COLUNAS_IP
  return df


def carregar_logs():
  conn = sqlite3.connect(DB_FILE)
  df = pd.read_sql_query(
      "SELECT data_hora, usuario, acao, tabela, detalhes FROM audit_logs ORDER"
      " BY id DESC",
      conn,
  )
  conn.close()
  df.columns = ["Data/Hora", "Usuário", "Ação", "Entidade", "Detalhes"]
  return df


# Inicializa DB
init_db()

# ==========================================
# GERENCIAMENTO DE SESSÃO / TELA DE LOGIN
# ==========================================
if "logged_in" not in st.session_state:
  st.session_state["logged_in"] = False
if "username" not in st.session_state:
  st.session_state["username"] = ""
if "user_role" not in st.session_state:
  st.session_state["user_role"] = ""

if not st.session_state["logged_in"]:
  st.title("🔒 Acesso Restrito - IPAM Enterprise")
  st.markdown(
      "Entre com suas credenciais de usuário para acessar o painel de redes."
  )

  with st.form("login_form"):
    user_input = st.text_input("Usuário")
    pass_input = st.text_input("Senha", type="password")
    submit_login = st.form_submit_button("Entrar")

    if submit_login:
      role = autenticar_usuario(user_input, pass_input)
      if role:
        st.session_state["logged_in"] = True
        st.session_state["username"] = user_input.strip()
        st.session_state["user_role"] = role
        st.success(f"Bem-vindo, {user_input}!")
        st.rerun()
      else:
        st.error("Usuário ou senha incorretos.")

  st.info("💡 **Acesso Padrão Inicial**: Usuário: `admin` | Senha: `admin`")
  st.stop()


# ==========================================
# EXPORTAÇÃO E ESTILIZAÇÃO DO EXCEL
# ==========================================
def gerar_excel_formatado():
  df_asn = carregar_dados_asn()
  df_ip = carregar_dados_ip()

  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  header_fill = PatternFill(
      start_color="1F497D", end_color="1F497D", fill_type="solid"
  )
  header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
  data_font = Font(name="Segoe UI", size=10)
  center_align = Alignment(
      horizontal="center", vertical="center", wrap_text=False
  )
  left_align = Alignment(horizontal="left", vertical="center", wrap_text=False)
  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )

  ws_asn = wb.create_sheet(title="ASNs")
  ws_asn.views.sheetView[0].showGridLines = True
  ws_asn.append(COLUNAS_ASN)
  for r in df_asn.itertuples(index=False):
    ws_asn.append(list(r))

  ws_ip = wb.create_sheet(title="Blocos_IP")
  ws_ip.views.sheetView[0].showGridLines = True
  ws_ip.append(COLUNAS_IP)
  for r in df_ip.itertuples(index=False):
    ws_ip.append(list(r))

  for ws in wb.worksheets:
    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
      cell.fill = header_fill
      cell.font = header_font
      cell.alignment = center_align

    for col in ws.columns:
      max_len = 0
      col_letter = get_column_letter(col[0].column)
      for cell in col:
        val_str = str(cell.value) if cell.value is not None else ""
        if len(val_str) > max_len:
          max_len = len(val_str)

        if cell.row > 1:
          cell.font = data_font
          cell.border = thin_border
          ws.row_dimensions[cell.row].height = 22
          if col_letter in ["A", "B", "C", "E"]:
            cell.alignment = center_align
          else:
            cell.alignment = left_align

      ws.column_dimensions[col_letter].width = max(max_len + 6, 18)

  output = io.BytesIO()
  wb.save(output)
  return output.getvalue()


# ==========================================
# FUNÇÕES DE INTELIGÊNCIA DE REDE
# ==========================================
def verificar_overlap(novo_prefixo):
  try:
    nova_rede = ipaddress.ip_network(novo_prefixo.strip(), strict=False)
    df_ip = carregar_dados_ip()
    conflitos = []

    for _, row in df_ip.iterrows():
      try:
        rede_existente = ipaddress.ip_network(row["Prefixo"], strict=False)
        if (
            nova_rede.overlaps(rede_existente)
            and nova_rede.version == rede_existente.version
        ):
          conflitos.append(row["Prefixo"])
      except ValueError:
        continue
    return conflitos
  except ValueError:
    return []


def buscar_ip_global(ip_busca):
  try:
    ip_obj = ipaddress.ip_address(ip_busca.strip())
    df_ip = carregar_dados_ip()
    encontrados = []

    for _, row in df_ip.iterrows():
      try:
        rede = ipaddress.ip_network(row["Prefixo"], strict=False)
        if ip_obj in rede:
          encontrados.append(row.to_dict())
      except ValueError:
        continue
    return pd.DataFrame(encontrados)
  except ValueError:
    return pd.DataFrame()


def consultar_rdap_asn(asn_num):
  clean_asn = asn_num.upper().replace("AS", "").strip()
  url = f"https://rdap.registro.br/autnum/{clean_asn}"
  try:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=4) as response:
      if response.status == 200:
        data = json.loads(response.read().decode())
        entidade = data.get("entities", [{}])[0].get("vcardArray", [])
        nome = "N/D"
        if len(entidade) > 1:
          for item in entidade[1]:
            if item[0] == "fn":
              nome = item[3]
        return {
            "ASN": f"AS{clean_asn}",
            "Nome / Razão Social": nome,
            "Handle": data.get("handle", "N/D"),
            "Status": "Ativo / Encontrado",
        }
  except Exception:
    return {
        "ASN": f"AS{clean_asn}",
        "Mensagem": "Consulta RDAP indisponível ou ASN não localizado.",
    }


# ==========================================
# INTERFACE STREAMLIT
# ==========================================
st.title("🌐 IPAM Enterprise - Gestão Integrada de Redes")

# Painel Lateral de Informações do Usuário & Logout
st.sidebar.markdown(
    f"👤 **Usuário**: `{st.session_state['username']}`"
    f" ({st.session_state['user_role']})"
)
if st.sidebar.button("🚪 Sair (Logout)"):
  st.session_state["logged_in"] = False
  st.session_state["username"] = ""
  st.session_state["user_role"] = ""
  st.rerun()

st.sidebar.divider()

df_asn = carregar_dados_asn()
df_ip = carregar_dados_ip()

# Opções de Menu
opcoes_menu = [
    "📊 Dashboard & Busca Global",
    "📋 Cadastrar ASN",
    "📌 Cadastrar Bloco IP",
    "✏️ Editar / Excluir Registros",
    "🧮 Calculadora & Subnetting",
    "⚡ Gerador Prefix-List BGP",
    "🛠️ Utilidades NOC (CLI & PTR)",
    "🔎 Consulta WHOIS/RDAP",
    "📜 Histórico & Audit Log",
    "☁️ Exportação Excel / Nuvem",
]

# Exibe o menu de Gestão de Usuários apenas para administradores
if st.session_state["user_role"] == "admin":
  opcoes_menu.append("🔐 Gestão de Usuários")

aba = st.sidebar.radio("Navegação Principal", opcoes_menu)

# ==========================================
# ABA 1: DASHBOARD & BUSCA GLOBAL
# ==========================================
if aba == "📊 Dashboard & Busca Global":
  st.header("📊 Visão Geral e Busca Rápida de Hosts")

  col1, col2, col3, col4 = st.columns(4)
  col1.metric("ASNs Cadastrados", len(df_asn))
  col2.metric(
      "Blocos IPv4",
      len(df_ip[df_ip["Versão"] == "IPv4"]) if not df_ip.empty else 0,
  )
  col3.metric(
      "Blocos IPv6",
      len(df_ip[df_ip["Versão"] == "IPv6"]) if not df_ip.empty else 0,
  )
  col4.metric(
      "VLANs Distintas",
      df_ip["VLAN ID"].nunique() if not df_ip.empty else 0,
  )

  st.divider()

  st.subheader("🎯 Buscar um bloco IP")
  ip_query = st.text_input(
      "Digite um IP para identificar o bloco, Pode ser Network, Gatewy ou Válido (ex:"
      " 192.168.1.1):"
  )
  if ip_query:
    res = buscar_ip_global(ip_query)
    if not res.empty:
      st.success(f"O IP `{ip_query}` pertence ao bloco cadastrado:")
      st.dataframe(res, use_container_width=True)
    else:
      st.warning(f"O IP `{ip_query}` não pertence a nenhum bloco cadastrado.")

  st.divider()

  col_t1, col_t2 = st.columns(2)
  with col_t1:
    st.subheader("📋 Tabela de ASNs")
    st.dataframe(df_asn, use_container_width=True)
  with col_t2:
    st.subheader("📌 Tabela de Blocos IP e Engenharia")
    st.dataframe(df_ip, use_container_width=True)

# ==========================================
# ABA 2: CADASTRAR ASN
# ==========================================
elif aba == "📋 Cadastrar ASN":
  st.header("📋 Registrar Novo ASN")
  with st.form("form_asn", clear_on_submit=True):
    asn_num = st.text_input("Número do ASN (ex: AS264698)")
    entidade = st.text_input("Nome da Entidade / Cliente (ex: BITCORE)")
    tipo = st.selectbox("Tipo", ["Próprio", "Cliente", "Trânsito/Upstream"])
    contato = st.text_input("E-mail / Contato NOC")

    if st.form_submit_button("Salvar ASN"):
      if asn_num and entidade:
        clean_asn = asn_num.upper().strip()
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        try:
          cursor.execute(
              "INSERT INTO asns VALUES (?, ?, ?, ?)",
              (clean_asn, entidade.strip(), tipo, contato.strip()),
          )
          conn.commit()
          registrar_log("CADASTRAR", "ASNs", f"ASN {clean_asn} ({entidade})")
          st.success(f"ASN {clean_asn} salvo com sucesso!")
        except sqlite3.IntegrityError:
          st.error(f"O ASN {clean_asn} já está cadastrado no banco!")
        finally:
          conn.close()
        st.rerun()
      else:
        st.error("Preencha ASN e Entidade.")

# ==========================================
# ABA 3: CADASTRAR BLOCO IP
# ==========================================
elif aba == "📌 Cadastrar Bloco IP":
  st.header("📌 Registrar Novo Bloco IP")

  with st.form("form_ip", clear_on_submit=True):
    col_a, col_b = st.columns(2)
    with col_a:
      prefixo = st.text_input("Bloco CIDR (ex: X.X.X.X/30)")
      asn_vinc = st.text_input("ASN Vinculado (ex: xxxxx)")
      finalidade = st.text_input("Finalidade (ex: INFORMAÇÃO / DESCRIÇÃO)")
    with col_b:
      vlan_id = st.text_input("VLAN ID (ex: 100)")
      vrf = st.text_input("VRF (ex: VRF-INTERNET)")
      pop = st.text_input("POP / Localidade (ex: POP-SP-01)")
      id_circuito = st.text_input("ID Circuito (ex: L2VPN-(NOME DO CLIENTE)-88412)")

    if st.form_submit_button("Salvar Bloco IP"):
      try:
        rede = ipaddress.ip_network(prefixo.strip(), strict=False)
        str_prefixo = str(rede)

        conflitos = verificar_overlap(str_prefixo)
        if conflitos:
          st.error(
              f"⚠️ Conflito de Roteamento! O bloco `{str_prefixo}` faz"
              f" sobreposição com os blocos já existentes: {conflitos}"
          )
        else:
          conn = sqlite3.connect(DB_FILE)
          cursor = conn.cursor()
          cursor.execute(
              "INSERT INTO blocos_ip VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
              (
                  str_prefixo,
                  f"IPv{rede.version}",
                  asn_vinc.upper().strip(),
                  finalidade.strip(),
                  vlan_id.strip(),
                  vrf.strip(),
                  pop.strip(),
                  id_circuito.strip(),
              ),
          )
          conn.commit()
          conn.close()
          registrar_log(
              "CADASTRAR",
              "Blocos_IP",
              f"Bloco {str_prefixo} (VLAN {vlan_id}, POP {pop})",
          )
          st.success(f"Bloco IP {str_prefixo} cadastrado com sucesso!")
          st.rerun()
      except ValueError:
        st.error("Formato CIDR inválido. Exemplo correto: 192.168.1.0/30")

# ==========================================
# ABA 4: EDITAR / EXCLUIR
# ==========================================
elif aba == "✏️ Editar / Excluir Registros":
  st.header("✏️ Gerenciar Registros Existentes")
  t1, t2 = st.tabs(["Gerenciar ASNs", "Gerenciar Blocos IP"])

  with t1:
    if df_asn.empty:
      st.info("Nenhum ASN cadastrado.")
    else:
      asn_sel = st.selectbox("Selecione o ASN:", df_asn["ASN"].tolist())
      row = df_asn[df_asn["ASN"] == asn_sel].iloc[0]

      with st.form("edit_asn"):
        e_entidade = st.text_input("Entidade", value=row["Entidade"])
        e_tipo = st.selectbox(
            "Tipo",
            ["Próprio", "Cliente", "Trânsito/Upstream"],
            index=["Próprio", "Cliente", "Trânsito/Upstream"].index(
                row["Tipo"]
            ),
        )
        e_contato = st.text_input("Contato", value=row["Contato"])

        c_save, c_del = st.columns(2)
        if c_save.form_submit_button("💾 Salvar Alterações"):
          conn = sqlite3.connect(DB_FILE)
          cursor = conn.cursor()
          cursor.execute(
              "UPDATE asns SET entidade=?, tipo=?, contato=? WHERE asn=?",
              (e_entidade, e_tipo, e_contato, asn_sel),
          )
          conn.commit()
          conn.close()
          registrar_log("EDITAR", "ASNs", f"ASN {asn_sel} atualizado")
          st.success(f"ASN {asn_sel} atualizado!")
          st.rerun()

        if c_del.form_submit_button("🗑️ Excluir ASN"):
          conn = sqlite3.connect(DB_FILE)
          cursor = conn.cursor()
          cursor.execute("DELETE FROM asns WHERE asn=?", (asn_sel,))
          conn.commit()
          conn.close()
          registrar_log("EXCLUIR", "ASNs", f"ASN {asn_sel} removido")
          st.warning(f"ASN {asn_sel} removido.")
          st.rerun()

  with t2:
    if df_ip.empty:
      st.info("Nenhum Bloco IP cadastrado.")
    else:
      ip_sel = st.selectbox("Selecione o Bloco IP:", df_ip["Prefixo"].tolist())
      row = df_ip[df_ip["Prefixo"] == ip_sel].iloc[0]

      with st.form("edit_ip"):
        col_e1, col_e2 = st.columns(2)
        with col_e1:
          e_asn_v = st.text_input("ASN Vinculado", value=row["ASN Vinculado"])
          e_fin = st.text_input("Finalidade", value=row["Finalidade"])
          e_vlan = st.text_input("VLAN ID", value=row["VLAN ID"])
        with col_e2:
          e_vrf = st.text_input("VRF", value=row["VRF"])
          e_pop = st.text_input(
              "POP / Localidade", value=row["POP / Localidade"]
          )
          e_circ = st.text_input("ID Circuito", value=row["ID Circuito"])

        c_s, c_d = st.columns(2)
        if c_s.form_submit_button("💾 Salvar Alterações"):
          conn = sqlite3.connect(DB_FILE)
          cursor = conn.cursor()
          cursor.execute(
              """
                        UPDATE blocos_ip 
                        SET asn_vinculado=?, finalidade=?, vlan_id=?, vrf=?, pop_localidade=?, id_circuito=?
                        WHERE prefixo=?
                    """,
              (e_asn_v, e_fin, e_vlan, e_vrf, e_pop, e_circ, ip_sel),
          )
          conn.commit()
          conn.close()
          registrar_log("EDITAR", "Blocos_IP", f"Bloco {ip_sel} atualizado")
          st.success(f"Bloco {ip_sel} atualizado!")
          st.rerun()

        if c_d.form_submit_button("🗑️ Excluir Bloco"):
          conn = sqlite3.connect(DB_FILE)
          cursor = conn.cursor()
          cursor.execute("DELETE FROM blocos_ip WHERE prefixo=?", (ip_sel,))
          conn.commit()
          conn.close()
          registrar_log("EXCLUIR", "Blocos_IP", f"Bloco {ip_sel} removido")
          st.warning(f"Bloco {ip_sel} removido.")
          st.rerun()

# ==========================================
# ABA 5: CALCULADORA & SUBNETTING
# ==========================================
elif aba == "🧮 Calculadora & Subnetting":
  st.header("🧮 Calculadora e Expansor de Sub-redes")
  pref = st.text_input(
      "Digite o bloco/sub-rede CIDR:", value="X.X.X.X/30"
  )

  if pref:
    try:
      rede = ipaddress.ip_network(pref.strip(), strict=False)
      c1, c2, c3, c4 = st.columns(4)
      c1.metric("Versão", f"IPv{rede.version}")
      c2.metric("Máscara CIDR", f"/{rede.prefixlen}")
      c3.metric("Total de IPs", f"{rede.num_addresses:,}")
      if rede.version == 4:
        c4.metric("Máscara Decimal", str(rede.netmask))

      st.divider()
      nova_mascara = st.slider(
          "Dividir em sub-redes menores:",
          min_value=rede.prefixlen + 1,
          max_value=32 if rede.version == 4 else 64,
          value=(
              rede.prefixlen + 2
              if rede.prefixlen < 30
              else rede.prefixlen
          ),
      )
      subredes = list(rede.subnets(new_prefix=nova_mascara))
      st.write(
          f"Bloco `{rede}` fatiado em **/{nova_mascara}** = **{len(subredes)}**"
          " sub-redes:"
      )

      dados_sub = []
      for s in subredes[:128]:
        if rede.version == 4 and s.prefixlen <= 30:
          dados_sub.append({
              "Sub-rede CIDR": str(s),
              "Network ID": str(s.network_address),
              "Primeiro IP Útil": str(s.network_address + 1),
              "Último IP Útil": str(s.broadcast_address - 1),
              "Broadcast": str(s.broadcast_address),
              "IPs Úteis": s.num_addresses - 2,
          })
        else:
          dados_sub.append({
              "Sub-rede CIDR": str(s),
              "Network ID": str(s.network_address),
              "Total de IPs": s.num_addresses,
          })

      st.dataframe(pd.DataFrame(dados_sub), use_container_width=True)
    except ValueError:
      st.error("Bloco CIDR inválido.")

# ==========================================
# ABA 6: GERADOR PREFIX-LIST
# ==========================================
elif aba == "⚡ Gerador Prefix-List BGP":
  st.header("⚡ Gerador de Prefix-Lists BGP")
  if df_ip.empty:
    st.info("Nenhum bloco cadastrado.")
  else:
    vendor = st.selectbox(
        "Vendor do Roteador:",
        [
            "Cisco (IOS/IOS-XE)",
            "Huawei (VRP)",
            "Juniper (Junos)",
            "MikroTik RouterOS v7",
        ],
    )
    nome_pl = st.text_input("Nome da Prefix-List:", "PL-BGP-ANNOUNCE")
    blocos_sel = st.multiselect(
        "Blocos para Anúncio:",
        options=df_ip["Prefixo"].tolist(),
        default=df_ip["Prefixo"].tolist(),
    )

    if blocos_sel and nome_pl:
      script_out = ""
      for idx, p in enumerate(blocos_sel, start=10):
        r = ipaddress.ip_network(p, strict=False)
        if vendor == "Cisco (IOS/IOS-XE)":
          script_out += f"{'ip' if r.version == 4 else 'ipv6'} prefix-list {nome_pl} seq {idx} permit {p}\n"
        elif vendor == "Huawei (VRP)":
          script_out += f"{'ip' if r.version == 4 else 'ipv6'} ip-prefix {nome_pl} index {idx} permit {r.network_address} {r.prefixlen}\n"
        elif vendor == "Juniper (Junos)":
          script_out += f"set policy-options prefix-list {nome_pl} {p}\n"
        elif vendor == "MikroTik RouterOS v7":
          script_out += f'/routing filter rule add chain={nome_pl} rule="if (dst in {p}) {{ accept }}"\n'

      st.code(script_out, language="text")

# ==========================================
# ABA 7: UTILIDADES NOC (CLI & PTR)
# ==========================================
elif aba == "🛠️ Utilidades NOC (CLI & PTR)":
  st.header("🛠️ Automação NOC: Configuração de Interfaces & DNS Reverso")
  if df_ip.empty:
    st.info("Cadastre blocos IP para utilizar a ferramenta NOC.")
  else:
    tab_cli, tab_ptr = st.tabs(
        ["Gerador de Configuração CLI", "Gerador de DNS Reverso (PTR)"]
    )

    with tab_cli:
      bloco_cli = st.selectbox(
          "Selecione o Bloco IP:", df_ip["Prefixo"].tolist()
      )
      row_cli = df_ip[df_ip["Prefixo"] == bloco_cli].iloc[0]
      vendor_cli = st.selectbox(
          "Selecione o Vendor:",
          [
              "Cisco (IOS-XE)",
              "Huawei (VRP)",
              "Juniper (Junos)",
              "MikroTik RouterOS v7",
          ],
      )

      r = ipaddress.ip_network(bloco_cli, strict=False)
      ip_gw = r.network_address + 1
      vlan = row_cli["VLAN ID"] or "100"
      desc = f"{row_cli['Finalidade']} - {row_cli['ID Circuito']}"

      st.subheader("Script de Interface / Sub-interface:")
      cli_code = ""

      if vendor_cli == "Cisco (IOS-XE)":
        cli_code = f"""interface GigabitEthernet0/0/0.{vlan}
 encapsulation dot1Q {vlan}
 description {desc}
 ip address {ip_gw} {r.netmask}
 no shutdown"""

      elif vendor_cli == "Huawei (VRP)":
        cli_code = f"""interface GigabitEthernet0/1/0.{vlan}
 dot1q termination vid {vlan}
 description {desc}
 ip address {ip_gw} {r.prefixlen}
 arp broadcast enable"""

      elif vendor_cli == "Juniper (Junos)":
        cli_code = f"""set interfaces ge-0/0/0 unit {vlan} vlan-id {vlan}
set interfaces ge-0/0/0 unit {vlan} description "{desc}"
set interfaces ge-0/0/0 unit {vlan} family inet address {ip_gw}/{r.prefixlen}"""

      elif vendor_cli == "MikroTik RouterOS v7":
        cli_code = f"""/interface vlan add name=vlan{vlan} vlan-id={vlan} interface=ether1 comment="{desc}"
/ip address add address={ip_gw}/{r.prefixlen} interface=vlan{vlan}"""

      st.code(cli_code, language="text")

    with tab_ptr:
      bloco_ptr = st.selectbox(
          "Selecione o Bloco para Zona Reverso:", df_ip["Prefixo"].tolist()
      )
      dominio_ptr = st.text_input("Domínio de Host (ex: noc.provedor.com.br)")

      if dominio_ptr:
        r_ptr = ipaddress.ip_network(bloco_ptr, strict=False)
        st.subheader("Entradas de Zona BIND9 / DNS Reverso:")
        ptr_out = ""

        if r_ptr.version == 4 and r_ptr.prefixlen <= 30:
          for ip in list(r_ptr.hosts())[:32]:
            octetos = str(ip).split(".")
            ptr_out += (
                f"{octetos[3]}\tIN\tPTR\thost-{octetos[3]}.{dominio_ptr}.\n"
            )
        else:
          ptr_out = "; Exemplo para bloco IPv6 ou sub-rede reduzida:\n"
          ptr_out += f"; Prefixo: {r_ptr.network_address} PTR {dominio_ptr}"

        st.code(ptr_out, language="text")

# ==========================================
# ABA 8: CONSULTA WHOIS/RDAP
# ==========================================
elif aba == "🔎 Consulta WHOIS/RDAP":
  st.header("🔎 Consulta RDAP/WHOIS de ASN")
  asn_query = st.text_input(
      "Digite o número do ASN para consulta (ex: 264698)"
  )

  if st.button("Consultar no Registro.br / RDAP"):
    if asn_query:
      dados_rdap = consultar_rdap_asn(asn_query)
      st.json(dados_rdap)

# ==========================================
# ABA 9: HISTÓRICO & AUDIT LOG
# ==========================================
elif aba == "📜 Histórico & Audit Log":
  st.header("📜 Histórico de Alterações (Audit Log)")
  df_logs = carregar_logs()
  if df_logs.empty:
    st.info("Nenhum registro de log até o momento.")
  else:
    st.dataframe(df_logs, use_container_width=True)

# ==========================================
# ABA 10: EXPORTAÇÃO EXCEL / NUVEM
# ==========================================
elif aba == "☁️ Exportação Excel / Nuvem":
  st.header("☁️ Exportar Planilha Formatada (.xlsx)")
  st.write("""
    Baixe a planilha multi-aba totalmente estilizada em Excel (`.xlsx`). 
    Ao abri-la no **Excel** ou no **Google Sheets**, as abas **ASNs** e **Blocos_IP** estarão perfeitamente separadas com cabeçalho azul e larguras ajustadas.
    """)

  excel_data = gerar_excel_formatado()
  st.download_button(
      label="📥 Baixar IPAM_Database.xlsx",
      data=excel_data,
      file_name="ipam_database.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

# ==========================================
# ABA 11: GESTÃO DE USUÁRIOS (APENAS ADMIN)
# ==========================================
elif aba == "🔐 Gestão de Usuários":
  # Validação de segurança estrita para Admin
  if st.session_state["user_role"] != "admin":
    st.error(
        "⚠️ Acesso Negado! Apenas administradores têm permissão para acessar a"
        " gestão de usuários."
    )
    st.stop()

  st.header("🔐 Controle de Usuários e Permissões de Acesso")

  t_list, t_add, t_edit = st.tabs([
      "📋 Usuários Cadastrados",
      "➕ Cadastrar Novo Usuário",
      "✏️ Editar / Excluir Usuário",
  ])

  # Sub-aba 1: Listar Usuários
  with t_list:
    conn = sqlite3.connect(DB_FILE)
    df_users = pd.read_sql_query("SELECT username, role FROM users", conn)
    conn.close()
    df_users.columns = ["Usuário (Login)", "Perfil de Acesso (Role)"]
    st.dataframe(df_users, use_container_width=True)

  # Sub-aba 2: Cadastrar Usuário
  with t_add:
    with st.form("form_novo_usuario", clear_on_submit=True):
      new_user = st.text_input("Novo Usuário (Login)")
      new_pass = st.text_input("Senha", type="password")
      new_role = st.selectbox(
          "Perfil de Permissão",
          ["operator", "admin"],
          help=(
              "Admin: Acesso total e Gestão de Usuários. Operator: Acesso"
              " completo às ferramentas de IPAM."
          ),
      )

      if st.form_submit_button("Cadastrar Usuário"):
        if new_user and new_pass:
          conn = sqlite3.connect(DB_FILE)
          cursor = conn.cursor()
          try:
            cursor.execute(
                "INSERT INTO users VALUES (?, ?, ?)",
                (new_user.strip(), make_hashes(new_pass), new_role),
            )
            conn.commit()
            registrar_log(
                "CADASTRAR",
                "Usuários",
                f"Usuário {new_user.strip()} cadastrado como {new_role}",
            )
            st.success(
                f"Usuário `{new_user.strip()}` registrado com sucesso!"
            )
            st.rerun()
          except sqlite3.IntegrityError:
            st.error("Este nome de usuário já existe no sistema.")
          finally:
            conn.close()
        else:
          st.error("Preencha o nome de usuário e a senha.")

  # Sub-aba 3: Editar Perfil / Excluir Usuário
  with t_edit:
    conn = sqlite3.connect(DB_FILE)
    df_u_edit = pd.read_sql_query("SELECT username, role FROM users", conn)
    conn.close()

    if df_u_edit.empty:
      st.info("Nenhum usuário cadastrado.")
    else:
      lista_usuarios = df_u_edit["username"].tolist()
      user_selecionado = st.selectbox(
          "Selecione o Usuário para Gerenciar:", lista_usuarios
      )
      role_atual = df_u_edit[df_u_edit["username"] == user_selecionado][
          "role"
      ].values[0]

      with st.form("form_editar_usuario"):
        st.subheader(f"Modificar Conta: `{user_selecionado}`")

        novo_perfil = st.selectbox(
            "Perfil de Acesso",
            ["operator", "admin"],
            index=0 if role_atual == "operator" else 1,
            help="Altere a função e as permissões deste usuário.",
        )

        nova_senha = st.text_input(
            "Nova Senha (deixe em branco caso não queira alterar a senha atual)",
            type="password",
        )

        btn_salvar, btn_excluir = st.columns(2)

        if btn_salvar.form_submit_button("💾 Salvar Alterações"):
          conn = sqlite3.connect(DB_FILE)
          cursor = conn.cursor()

          if nova_senha.strip():
            # Atualiza perfil e senha
            hash_nova_senha = make_hashes(nova_senha.strip())
            cursor.execute(
                "UPDATE users SET role=?, password=? WHERE username=?",
                (novo_perfil, hash_nova_senha, user_selecionado),
            )
            detalhe_log = (
                f"Perfil do usuário {user_selecionado} alterado para"
                f" {novo_perfil} e senha redefinida"
            )
          else:
            # Atualiza apenas perfil
            cursor.execute(
                "UPDATE users SET role=? WHERE username=?",
                (novo_perfil, user_selecionado),
            )
            detalhe_log = (
                f"Perfil do usuário {user_selecionado} alterado para"
                f" {novo_perfil}"
            )

          conn.commit()
          conn.close()
          registrar_log("EDITAR", "Usuários", detalhe_log)
          st.success(f"Usuário `{user_selecionado}` atualizado com sucesso!")
          st.rerun()

        if btn_excluir.form_submit_button("🗑️ Excluir Usuário"):
          if user_selecionado == st.session_state["username"]:
            st.error(
                "⚠️ Operação não permitida! Você não pode excluir a sua própria"
                " conta enquanto estiver logado."
            )
          else:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute(
                "DELETE FROM users WHERE username=?", (user_selecionado,)
            )
            conn.commit()
            conn.close()
            registrar_log(
                "EXCLUIR",
                "Usuários",
                f"Usuário {user_selecionado} excluído do sistema",
            )
            st.warning(f"Usuário `{user_selecionado}` removido com sucesso!")
            st.rerun()