# 📄 Gerador Automático de Propostas Comerciais

Um sistema web interno desenvolvido para automatizar, padronizar e dar agilidade à criação de propostas comerciais em formato Word (.docx). 

Este projeto foi desenhado para eliminar o retrabalho manual, evitar erros de cálculo e garantir a integridade visual do layout oficial dos documentos, gerando propostas prontas para envio em poucos segundos.

##  O Problema que Resolvemos
O fluxo tradicional de criação de propostas envolvia a edição manual de templates no Word, cópia de dados, deleção de linhas de tabelas vazias e o cálculo humano de subtotais e descontos. Isso consumia tempo valioso da equipe comercial e abria margem para erros operacionais.

**A Solução:** Uma interface limpa onde o usuário apenas insere os dados do cliente e os serviços. O sistema processa a matemática, injeta as informações no template e entrega o documento final formatado.

##  Principais Funcionalidades

*   **Interface Web Intuitiva:** Formulário simples criado com Streamlit, acessível para qualquer membro da equipe sem necessidade de conhecimentos técnicos.
*   **Motor Matemático Integrado:** Soma automática de todos os serviços listados e subtração de descontos em tempo real, garantindo 100% de precisão financeira.
*   **Ajuste Dinâmico de Tabelas (Manipulação de XML):** O código identifica os serviços não utilizados e acessa a estrutura profunda do Word para deletar as linhas excedentes fisicamente, mantendo o documento sem espaços em branco.
*   **Geração de Datas:** Preenchimento automático da data de emissão.
*   **Segurança e Privacidade:** O processamento ocorre localmente. Nenhum dado financeiro ou de cliente é armazenado ou enviado para nuvens externas.

##  Tecnologias Utilizadas

*   **Python:** Lógica principal, cálculos e estruturação de dados.
*   **Streamlit:** Criação da interface web (Front-end) de forma ágil e responsiva.
*   **python-docx:** Biblioteca para leitura, manipulação de tags e reescrita de arquivos `.docx`.
