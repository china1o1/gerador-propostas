import streamlit as st
from docx import Document
import io
from datetime import date, timedelta
import random

# Função para gerar as datas automaticamente
def gerar_datas():
    hoje = date.today()
    data_emissao = hoje.strftime("%d/%m/%Y")
    ultimo_dia_ano = date(hoje.year, 12, 31)
    dias_restantes = (ultimo_dia_ano - hoje).days
    dias_aleatorios = random.randint(15, dias_restantes) if dias_restantes >= 15 else 15
    data_fim = hoje + timedelta(days=dias_aleatorios)
    data_vigencia = f"{data_emissao} - {data_fim.strftime('%d/%m/%Y')}"
    return data_emissao, data_vigencia

# Função que procura as TAGS no Word (nas linhas e nas tabelas)

# Função que procura as TAGS no Word (nas linhas, tabelas e cabeçalhos)
def substituir_texto(doc, dicionario_dados):
    # 1. Substituir no corpo do texto (parágrafos e tabelas)
    for paragrafo in doc.paragraphs:
        for tag, valor in dicionario_dados.items():
            if tag in paragrafo.text:
                paragrafo.text = paragrafo.text.replace(tag, str(valor))
                
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    for tag, valor in dicionario_dados.items():
                        if tag in paragrafo.text:
                            paragrafo.text = paragrafo.text.replace(tag, str(valor))

    # 2. Varredura completa em TODOS os tipos de cabeçalho
    for secao in doc.sections:
        # Lista todos os cabeçalhos possíveis (Principal, Primeira Página, Páginas Pares)
        headers = [secao.header, secao.first_page_header, secao.even_page_header]
        
        for header in headers:
            # Procura em parágrafos do cabeçalho
            for paragrafo in header.paragraphs:
                for tag, valor in dicionario_dados.items():
                    if tag in paragrafo.text:
                        paragrafo.text = paragrafo.text.replace(tag, str(valor))
            
            # Procura em tabelas dentro do cabeçalho
            for tabela in header.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        for paragrafo in celula.paragraphs:
                            for tag, valor in dicionario_dados.items():
                                if tag in paragrafo.text:
                                    paragrafo.text = paragrafo.text.replace(tag, str(valor))
# ==========================================
# Função ESPECÍFICA para buscar dentro de Caixas de Texto (Shapes)
def substituir_em_shapes(doc, dicionario_dados):
    for shape in doc.shapes:
        if shape.has_text_frame:
            for paragrafo in shape.text_frame.paragraphs:
                for tag, valor in dicionario_dados.items():
                    if tag in paragrafo.text:
                        paragrafo.text = paragrafo.text.replace(tag, str(valor))
# INTERFACE DO SITE
# ==========================================
st.set_page_config(page_title="Gerador de Propostas SENAI", page_icon="📄", layout="wide")

st.title("📄 Gerador de Propostas - SENAI")
st.write("Preencha os dados abaixo para gerar o documento Word formatado automaticamente.")

st.subheader("1. Dados do Cliente")
col1, col2 = st.columns(2)

with col1:
    numero_proposta = st.text_input("Número da Proposta (Ex: PRO-210/2026)")
    nome_empresa = st.text_input("Nome da Empresa Solicitante")
    cnpj = st.text_input("CNPJ")
    endereco = st.text_input("Endereço de Execução")

with col2:
    telefone = st.text_input("Telefone de Contato")
    email = st.text_input("E-mail")
    num_pessoas = st.text_input("Nº de pessoas atendidas")

st.subheader("2. Dados do Serviço e Valores")
col3, col4 = st.columns(2)

with col3:
    servico = st.text_input("Serviço (Ex: Curso de Excel Avançado e Power BI)")
    descricao = st.text_area("Descrição do Serviço")
    unidade = st.text_input("Unidade Executora (Ex: SENAI - Santo Amaro)")

with col4:
    qtd = st.text_input("Quantidade")
    valor_un = st.text_input("Valor Unitário (Ex: R$ 400,00)")
    valor_total = st.text_input("Valor Total (Ex: R$ 6.000,00)")

st.write("---")

# Botão principal
if st.button("Gerar Proposta 🚀"):
    if not nome_empresa:
        st.warning("⚠️ Por favor, preencha pelo menos o nome da empresa.")
    else:
        try:
            # Carrega o documento Word que serve de molde
            doc = Document("Template_Proposta.docx")
            
            # Gera as datas
            data_emissao, data_vigencia = gerar_datas()
            
            # DICIONÁRIO COMPLETO COM TODOS OS CAMPOS DO SITE
            dados_para_trocar = {
                "{{NUMERO_DA_PROPOSTA}}": numero_proposta,
                "{{EMPRESA}}": nome_empresa,
                "{{CNPJ}}": cnpj,
                "{{ENDERECO}}": endereco,
                "{{TELEFONE}}": telefone,
                "{{EMAIL}}": email,
                "{{NUM_PESSOAS}}": num_pessoas,
                "{{SERVICO}}": servico,
                "{{DESCRICAO}}": descricao,
                "{{UNIDADE}}": unidade,
                "{{QTD}}": qtd,
                "{{VALOR_UN}}": valor_un,
                "{{VALOR_TOTAL}}": valor_total,
                "{{DATA_EMISSAO}}": data_emissao,
                "{{DATA_VIGENCIA}}": data_vigencia,
                "{{RESPONSAVEL}}": "ELITON GABRIEL SILVA CORDEIRO" # Mantive o seu nome fixo
            }
            
            # Executa a varredura e substituição no documento
            substituir_texto(doc, dados_para_trocar)
            
            # Salva o documento pronto na "memória" do site para download
            arquivo_memoria = io.BytesIO()
            doc.save(arquivo_memoria)
            arquivo_memoria.seek(0)
            
            st.success("✨ Proposta processada com sucesso!")
            
            # Cria o botão de download para o usuário baixar
            st.download_button(
                label="📥 Baixar Documento Word Pronto",
                data=arquivo_memoria,
                file_name=f"Proposta_{nome_empresa}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error("Erro ao gerar. Verifique se o arquivo 'Template_Proposta.docx' está na mesma pasta e fechado.")