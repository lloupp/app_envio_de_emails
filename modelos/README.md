# Modelos de planilha

Planilhas de exemplo com dados fictícios para testar os scripts do repositório. Nenhum e-mail ou nome real é usado.

- **modelo_gerador_email.xlsx** — planilha genérica para o [gerador_email.py](../gerador_email.py) (app Streamlit). Colunas: `Nome`, `Email`, `CC`, `BCC`, `Curso`, `Anexo`.
- **modelo_certificados_lote.xlsx** — estrutura esperada por [gerar_certificados_lote.py](../gerar_certificados_lote.py) (aba `Sheet1`). Colunas: `NOME`, `E-mail`, `Modalidade`, `TEXTO`, `ESP`, `DE`, `ATE`, `CH`, `anexo`.
- **modelo_certificados_simples.xlsx** — estrutura esperada por [gerar_certificados.py](../gerar_certificados.py) (aba `Envio`). Colunas: `Nome do Destinatário`, `Email`, `Anexo`, `Nome do Curso`, `Data do Curso`, `Tipo Participação`.

Os nomes de arquivo em `Anexo`/`anexo` são apenas ilustrativos; para testar o envio de anexos, crie arquivos com esses nomes na pasta indicada ao rodar o script.
