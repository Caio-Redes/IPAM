# 🌐 IPAM Enterprise - Gestão de IPs, ASNs e Engenharia de Rede

O **IPAM Enterprise** é uma plataforma web construída com **Python** e **Streamlit** para gerenciamento de endereçamento IP (IPv4/IPv6), Autonomous System Numbers (ASNs), geração de scripts de rede e automação NOC para provedores de internet (ISPs) e engenharia de redes enterprise.

---

## 🚀 Funcionalidades

- **📊 Dashboard & Busca Global**: Métricas em tempo real e ponteiro de busca rápida para rastrear IPs, VLANs e VRFs vinculadas.
- **📋 Gestão de ASNs & Blocos IP**: Cadastro, edição e remoção de prefixos com verificação automática de sobreposição de rotas (*overlap*).
- **🧮 Calculadora de Sub-redes**: Fatiamento dinâmico CIDR e detalhamento de Network ID, Broadcast e IPs úteis.
- **⚡ Gerador de Prefix-List BGP**: Geração automática de sintaxe para **Cisco (IOS/IOS-XE)**, **Huawei (VRP)**, **Juniper (Junos)** e **MikroTik (RouterOS v7)**.
- **🛠️ Automação NOC**: Gerador de configuração de sub-interfaces e gerador de zonas para DNS Reverso (PTR / BIND9).
- **🔎 Consulta WHOIS/RDAP**: Integração via API RDAP com o Registro.br para verificação de ASNs ativos.
- **☁️ Exportação Estilizada**: Download de relatórios em Excel (`.xlsx`) multi-aba formatados com estilos visuais e larguras auto-ajustáveis.
- **📜 Audit Log**: Registro em banco de dados de todas as operações de criação, edição e exclusão de recursos.
- **🔐 Controle de Acesso (RBAC)**: Autenticação segura com senhas em hash SHA-256 e gestão de usuários (perfis `admin` e `operator`).

---

## 🛠️ Tecnologias Utilizadas

- **Linguagem**: Python 3.10+
- **Interface Web**: Streamlit
- **Banco de Dados**: SQLite3
- **Processamento de Dados**: Pandas, OpenPyXL
- **Manipulação de Redes**: `ipaddress` (built-in)

---

## 📋 Pré-requisitos e Instalação

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/Caio-Redes/IPAM.git](https://github.com/Caio-Redes/IPAM.git)
   cd IPAM
