# Agente IA VR/VA

## 🧾 Descrição

Este projeto implementa um Agente de Inteligência Artificial para automatizar o cálculo e compra mensal de VR/VA (Vale Refeição e Vale Alimentação).

Ele combina:

- Planilhas corporativas (ZIP descompactado automaticamente)
- Regras dos acordos coletivos (PDFs dos sindicatos + documento de descrição do desafio via RAG)
- Modelo de IA Gemini para responder perguntas, explicar cálculos e validar regras

Entrega final:

- Resumo interativo dos dados (via frontend em Streamlit)
- Chat com o agente IA
- Planilha consolidada no layout exigido (VR Mensal) com cálculos automáticos e validações.

---

## 📋 Requisitos

- Python 3.10 ou superior
- Conta na OpenAI com uma API Key do Gemini
- Dependências listadas em requirements.txt

---

## 🛠️ Instalação

```bash
1. Clone o repositório:

git clone agente-ia-vrva
cd agente-ia-vrva

2. Crie um ambiente virtual:

python -m venv venv

3. Ative o ambiente virtual:

Windows (PowerShell):
.\venv\Scripts\activate

Linux / Mac:
source venv/bin/activate

4. Instale as dependências:

bash
pip install -r requirements.txt

5. Editar o arquivo chamado .env na raiz do projeto com o seguinte conteúdo:

OPENAI_API_KEY="sua-chave-da-openai-aqui"

6. Rodar o script principal e Front:

bash
streamlit run agente.py
streamlit run front.py
```

---

## 🚀 Uso

- Abra o frontend em: http://localhost:8501
- O sistema exibirá:
    - Resumo dos dados das planilhas
    - Status do agente (carregando / pronto)
    - Chat interativo com o agente IA
- É possível:
    - Perguntar sobre colaboradores, sindicatos, férias, desligamentos, admissões etc.
    - Solicitar cálculos automáticos (dias úteis, VR proporcional, afastamentos).
    - Gerar e baixar a planilha final consolidada no layout exigido (VR Mensal).

---

✅ Exemplo de execução:

Digite sua pergunta: Quantos colaboradores ativos temos?