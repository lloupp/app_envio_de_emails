# 📧 Alfredo do Email

**Versão 2.0** - Sistema inteligente de automação de envio de e-mails em massa via Microsoft Outlook

Alfredo do Email é uma ferramenta desenvolvida com Python e Streamlit que facilita o envio de e-mails personalizados em massa através do Microsoft Outlook. Com interface intuitiva e recursos avançados, permite automação completa do processo de envio de e-mails com anexos.

## ✨ Funcionalidades

### 📊 Principais Recursos

- **Upload de Planilhas**: Suporte completo para arquivos `.xlsx` (Excel) e `.csv`
- **Mapeamento Dinâmico**: Escolha quais colunas representam:
  - Destinatário (Para)
  - Cópia (CC)
  - Cópia Oculta (BCC)
  - Nome do arquivo anexo
- **Editor de Texto Rico**: Formatação avançada com negrito, listas, links e mais
- **Tags Personalizadas**: Use `{NomeDaColuna}` no assunto ou corpo para personalizar cada mensagem
- **Integração com Outlook**: E-mails gerados como rascunhos, preservando sua assinatura padrão
- **Modo Teste**: Gere apenas o primeiro registro para validar antes de processar toda a lista
- **Validação de E-mails**: Verificação automática de formato de e-mails
- **Relatório de Erros**: Identificação detalhada de problemas durante o processamento
- **Sistema de Logs**: Registro completo de todas as operações para auditoria

### 🆕 Novidades da Versão 2.0

- ✅ Código completamente refatorado e modular
- ✅ Validação robusta de e-mails e caminhos de arquivo
- ✅ Sistema de logging para debugging e auditoria
- ✅ Type hints para melhor manutenção do código
- ✅ Tratamento de erros aprimorado
- ✅ Testes unitários para funções críticas
- ✅ Configurações centralizadas
- ✅ Documentação completa com docstrings
- ✅ Interface mais intuitiva e informativa
- ✅ Melhor feedback visual de progresso e erros

## 🛠️ Pré-requisitos

Para rodar este projeto, você precisará de:

1. **Sistema Operacional**: Windows (necessário para a integração com win32com)
2. **Microsoft Outlook**: Instalado e configurado com uma conta ativa
3. **Python 3.8+**: Instalado no sistema

## 🚀 Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/lloupp/app_envio_de_emails.git
cd app_envio_de_emails
```

### 2. Crie um ambiente virtual (recomendado)

```bash
python -m venv venv
venv\Scripts\activate  # No Windows
# ou
source venv/bin/activate  # No Linux/Mac
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

### 4. Execute a aplicação

```bash
streamlit run gerador_email.py
```

A aplicação abrirá automaticamente no seu navegador padrão.

## 📖 Como Usar

### 1. Prepare sua planilha

Certifique-se de que sua planilha contém:
- Uma coluna com endereços de e-mail dos destinatários
- (Opcional) Colunas para CC e BCC
- (Opcional) Uma coluna com o nome exato dos arquivos de anexo

**Exemplo de estrutura:**

| Nome  | Email              | Documento       |
|-------|-------------------|-----------------|
| João  | joao@empresa.com  | certificado_joao.pdf |
| Maria | maria@empresa.com | certificado_maria.pdf |

### 2. Faça upload da planilha

- Clique em "Suba sua Planilha" na barra lateral
- Selecione seu arquivo `.xlsx` ou `.csv`
- Se for Excel com múltiplas abas, selecione a aba desejada

### 3. Configure a pasta de anexos

1. Abra a pasta onde estão seus arquivos no Windows Explorer
2. Clique na barra de endereços no topo da pasta
3. Copie o caminho (ex: `C:\MeusDocumentos\Certificados`)
4. Cole no campo "Caminho da pasta de anexos"

### 4. Mapeie as colunas

- **Coluna 'Para'**: Selecione a coluna com os e-mails dos destinatários (obrigatório)
- **Coluna 'CC'**: Selecione a coluna para cópia (opcional)
- **Coluna 'BCC'**: Selecione a coluna para cópia oculta (opcional)
- **Coluna do Anexo**: Selecione a coluna com os nomes dos arquivos (opcional)

### 5. Redija seu e-mail

**Assunto:**
```
Certificado de Conclusão - {Nome}
```

**Corpo:**
```
Olá {Nome},

Segue em anexo seu certificado do curso.

Atenciosamente,
Equipe de Treinamento
```

As tags `{Nome}` serão substituídas pelos valores da planilha automaticamente.

### 6. Gere os e-mails

- **Modo Teste**: Clique em "🧪 Gerar 1 TESTE" para criar apenas o primeiro e-mail
  - Verifique o rascunho no Outlook
  - Confira se tudo está correto

- **Envio em Massa**: Clique em "🚀 Gerar TODOS" para processar toda a planilha
  - Acompanhe o progresso na barra
  - Verifique o relatório de conclusão

### 7. Envie os e-mails

1. Abra o Microsoft Outlook
2. Vá até a pasta "Rascunhos"
3. Revise os e-mails gerados
4. Envie manualmente ou use a opção de envio em massa do Outlook

## 🔧 Configurações Avançadas

### Arquivo de Configuração

O arquivo `config.py` permite customizar:

- Extensões de arquivo permitidas para anexos
- Tamanho máximo de anexos
- Configurações de logging
- Mensagens de erro personalizadas

### Sistema de Logs

Todos os eventos são registrados no arquivo `email_generator.log`:
- Sucessos e erros no processamento
- Validações de e-mail
- Arquivos não encontrados
- Conexões com o Outlook

## 🧪 Testes

Execute os testes unitários com:

```bash
python test_gerador_email.py
```

Os testes cobrem:
- Validação de e-mails
- Sanitização de caminhos
- Validação de anexos
- Substituição de tags
- Carregamento de DataFrames

## 📝 Estrutura do Projeto

```
app_envio_de_emails/
├── gerador_email.py          # Aplicação principal
├── config.py                 # Configurações
├── test_gerador_email.py     # Testes unitários
├── requirements.txt          # Dependências
├── README.md                 # Documentação
├── .gitignore               # Arquivos ignorados
└── email_generator.log      # Log de operações (gerado automaticamente)
```

## 🔒 Segurança

- ✅ Validação rigorosa de endereços de e-mail
- ✅ Sanitização de caminhos de arquivo
- ✅ Validação de extensões de arquivo para anexos
- ✅ Logs detalhados para auditoria
- ✅ Sem armazenamento de credenciais (usa Outlook local)

## 🐛 Solução de Problemas

### Erro ao conectar com o Outlook

**Problema**: "Erro ao conectar com o Outlook"

**Soluções**:
1. Certifique-se de que o Outlook está instalado
2. Abra o Outlook pelo menos uma vez antes de usar a ferramenta
3. Configure uma conta no Outlook
4. Execute o script como administrador (se necessário)

### Arquivos não encontrados

**Problema**: "Arquivo não encontrado" nos anexos

**Soluções**:
1. Verifique se o caminho da pasta está correto
2. Certifique-se de que os nomes na planilha correspondem exatamente aos nomes dos arquivos
3. Verifique se os arquivos existem na pasta especificada
4. Use o caminho completo, não relativo

### E-mails inválidos

**Problema**: "E-mail inválido" durante o processamento

**Soluções**:
1. Verifique se a coluna selecionada contém e-mails válidos
2. Remova espaços extras nas células
3. Certifique-se de que o formato é `usuario@dominio.com`

## 🤝 Contribuindo

Contribuições são bem-vindas! Para contribuir:

1. Faça um Fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 👥 Autor

Desenvolvido com ❤️ para automatizar processos de envio de e-mails

## 📞 Suporte

Se encontrar problemas ou tiver sugestões:
- Abra uma [Issue](https://github.com/lloupp/app_envio_de_emails/issues)
- Consulte a documentação acima
- Verifique os logs em `email_generator.log`

## 🗺️ Roadmap

Funcionalidades planejadas para versões futuras:

- [ ] Suporte para templates de e-mail salvos
- [ ] Agendamento de envios
- [ ] Relatórios de envio em PDF/Excel
- [ ] Suporte para múltiplos anexos por e-mail
- [ ] Interface web para visualização de estatísticas
- [ ] Suporte para outros clientes de e-mail (Gmail, etc.)
- [ ] Validação avançada de anexos (tamanho, tipo MIME)
- [ ] Sistema de filas para grandes volumes

---

**Nota**: Este software é fornecido "como está", sem garantias de qualquer tipo. Use por sua conta e risco.
