# Monitor de Preços - Mercado Livre

Um script Python automatizado para monitorar preços de produtos no Mercado Livre, registrar consultas em uma planilha Excel e enviar alertas por e-mail quando o preço cai abaixo de um limite definido.

## Funcionalidades

- **Monitoramento Automatizado**: Baixa a página do produto e extrai o preço atual usando técnicas robustas de scraping (meta tags, JSON-LD, componentes visuais).
- **Registro em Excel**: Salva cada consulta com data/hora, preço e URL em uma planilha Excel.
- **Alertas por E-mail**: Envia notificações automáticas via SMTP quando o preço fica abaixo do limite configurado.
- **Configuração Segura**: Usa variáveis de ambiente para credenciais e configurações, mantendo a segurança.
- **Flexibilidade**: Suporte a diferentes formatos de preço (vírgula ou ponto) e configurações personalizáveis.

## Pré-requisitos

- Python 3.8 ou superior
- Conta de e-mail com suporte a SMTP (ex.: Gmail com senha de app)
- Acesso à internet para scraping

## Instalação

1. Clone este repositório:
   ```bash
   git clone https://github.com/snowz004/monitor-precos-ml.git
   cd monitor-precos-ml
   ```

2. Crie um ambiente virtual (recomendado):
   ```bash
   python -m venv venv
   source venv/bin/activate  # No Windows: venv\Scripts\activate
   ```

3. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

## Configuração

1. Copie o arquivo de exemplo para criar seu `.env`:
   ```bash
   cp .env.example .env
   ```

2. Edite o arquivo `.env` com suas configurações:
   - `EMAIL_TO`: E-mail destinatário dos alertas.
   - `PRICE_THRESHOLD`: Limite de preço em reais (use ponto ou vírgula, ex.: 2500 ou 2500,99).
   - `EXCEL_PATH`: Caminho do arquivo Excel (opcional; padrão: `consultas_preco.xlsx`).
   - `SMTP_HOST`: Servidor SMTP (ex.: `smtp.gmail.com`).
   - `SMTP_PORT`: Porta SMTP (ex.: 587).
   - `SMTP_USER`: Seu e-mail para login SMTP.
   - `SMTP_PASSWORD`: Senha ou token de app para SMTP.
   - `EMAIL_FROM`: E-mail remetente (geralmente o mesmo que SMTP_USER).

   **Exemplo de configuração para Gmail:**
   ```
   EMAIL_TO=seu@email.com
   PRICE_THRESHOLD=2500
   SMTP_HOST=smtp.gmail.com
   SMTP_PORT=587
   SMTP_USER=seu@gmail.com
   SMTP_PASSWORD=sua_senha_de_app
   EMAIL_FROM=seu@gmail.com
   ```

   > **Nota de Segurança**: Nunca commite o arquivo `.env` no Git. Ele está listado no `.gitignore`.

## Como Rodar

Execute o script principal:
```bash
python main.py
```

O script irá:
- Baixar a página do produto configurado.
- Extrair o preço atual.
- Registrar a consulta no Excel.
- Verificar se o preço está abaixo do limite e enviar e-mail se necessário.

Para automação, configure um cron job ou tarefa agendada (ex.: no Linux com `crontab -e`):
```
# Executa a cada hora
0 * * * * /caminho/para/venv/bin/python /caminho/para/monitor-precos-ml/main.py
```

## Estrutura do Projeto

- `main.py`: Script principal com lógica de scraping, parsing e envio de e-mail.
- `requirements.txt`: Dependências Python.
- `.env.example`: Modelo de configuração.
- `.gitignore`: Arquivos ignorados pelo Git.

## Tecnologias Utilizadas

- **Python**: Linguagem principal.
- **Requests**: Para downloads HTTP.
- **BeautifulSoup**: Para parsing HTML.
- **OpenPyXL**: Para manipulação de Excel.
- **python-dotenv**: Para carregamento de variáveis de ambiente.
- **smtplib**: Para envio de e-mails.

## Contribuição

Contribuições são bem-vindas! Abra uma issue ou pull request no GitHub.

## Licença

Este projeto é distribuído sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.
