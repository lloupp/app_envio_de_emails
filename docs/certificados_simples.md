# Certificados simples — conferência antes dos rascunhos

O script `gerar_certificados.py` lê `planilha_envio_certificados.xlsx`, aba `Envio`, na mesma pasta dos PDFs. Ele **cria rascunhos no Outlook Desktop, sem enviar e-mails**.

## Utilização no Windows

```powershell
pip install openpyxl pywin32
python gerar_certificados.py "C:\Certificados" --dry-run
python gerar_certificados.py "C:\Certificados" --relatorio-csv "C:\Certificados\resultado.csv"
```

O comando com `--dry-run` verifica toda a planilha, destinatários e anexos **sem abrir o Outlook**. Revise o resultado antes de executar o segundo comando (sem `--dry-run`). O CSV de relatório é opcional e contém nomes e e-mails: mantenha-o apenas em local autorizado.

As colunas exigidas são: `Nome do Destinatário`, `Email`, `Anexo`, `Nome do Curso`, `Data do Curso` e `Tipo Participação`.

## Regras de proteção

- Um rascunho só é criado quando o endereço é válido, os campos essenciais estão preenchidos e o **PDF existe dentro da pasta selecionada**.
- Caminhos absolutos, referências fora da pasta, anexos que não sejam PDF e pares repetidos de endereço+arquivo são ignorados e registrados na saída.
- A assinatura configurada no Outlook é mantida. Erros por destinatário são relatados; o programa não envia mensagens automaticamente.
- O processamento sinaliza erro (código de saída `1`) quando houver linha ignorada, arquivo ausente ou falha no Outlook.

Para testar sem Windows: `python -m unittest discover -s tests -v` (requer `openpyxl`).
