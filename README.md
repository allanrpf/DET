# DET
PT-BR: DET - Disparador de E-mails em massa (com anexos diferentes para destinatários diferentes.) EN-US: DET - Mass Email Trigger (with different attachments for different recipients.)

### *Versão em Português* 🇧🇷

# 📨 Disparador de E-mails em Massa (DET)

Esta é uma aplicação de desktop para Windows, desenvolvida em AutoIt, que automatiza o envio de e-mails em massa com anexos individualizados. A ferramenta foi projetada para otimizar processos repetitivos, como o envio de recibos de pagamento, relatórios ou comunicados, utilizando uma simples planilha CSV como fonte de dados.

O script utiliza o *PowerShell* em segundo plano para realizar o envio dos e-mails via SMTP do Gmail , garantindo uma comunicação segura e robusta.

### Principais Funcionalidades

 *Envio Baseado em CSV*: Carrega uma lista de destinatários e dados de um arquivo `.csv`, onde cada linha corresponde a um envio.  
 *Personalização Dinâmica*: Utiliza placeholders como `{nome}` e `{mes}` no assunto e corpo do e-mail para criar mensagens personalizadas.  
 *Agrupamento Inteligente de Anexos*: Permite agrupar múltiplos arquivos para um mesmo destinatário em um único e-mail, otimizando a comunicação.  
 *Validação Pré-Envio*: Uma ferramenta para validar a existência dos anexos, calcular o tamanho total, gerar hashes MD5 e identificar possíveis divergências entre o nome do destinatário e o nome do arquivo antes de iniciar os envios.  
* *Criação de CSV a partir de Pasta*: Gera automaticamente um arquivo CSV modelo a partir de uma pasta de arquivos, preenchendo o nome do destinatário e o caminho do anexo, agilizando a configuração inicial.
 *Controle Total sobre o Envio*: Oferece funcionalidades para Pausar, Continuar e Cancelar o processo de envio em tempo real.  
 *Tentativas de Reenvio*: Configuração para tentar reenviar e-mails que falharam, com definição do número de tentativas e do atraso entre elas.  
 *Logging e Relatórios*: Gera um log detalhado (`log_envios.csv`) de todas as transações (sucessos e falhas) e permite exportar relatórios filtrados.  
 *Backup de Sucesso*: Cria um backup organizado em pastas dos arquivos que foram enviados com sucesso.  
 *Interface Gráfica Intuitiva*: Interface com abas para configurações básicas e avançadas, com feedback visual de progresso e estatísticas de envio.  

---

### *English Version*

# 📨 Bulk Email Sender (DET)

This is a Windows desktop application developed in AutoIt to automate sending bulk emails with individualized attachments. The tool is designed to streamline repetitive tasks such as sending payment slips, reports, or newsletters, using a simple CSV file as a data source.

The script leverages *PowerShell* in the background to handle email dispatch via Gmail's SMTP , ensuring secure and robust communication.

### Key Features

 *CSV-Based Dispatch*: Loads a list of recipients and related data from a `.csv` file, where each row corresponds to an email.  
 *Dynamic Personalization*: Uses placeholders like `{nome}` (name) and `{mes}` (month) in the subject and body to create personalized messages.  
 *Smart Attachment Grouping*: Optionally groups multiple files for the same recipient into a single email, optimizing communication.  
 *Pre-Send Validation*: A tool to validate attachment existence, calculate total size, generate MD5 hashes, and identify potential discrepancies between recipient and file names before starting the dispatch.  
* *CSV Creation from Folder*: Automatically generates a template CSV file from a folder of files, populating the recipient's name and attachment path to speed up the initial setup.
 *Full Dispatch Control*: Provides real-time controls to Pause, Resume, and Cancel the sending process.  
 *Retry Mechanism*: Option to configure retry attempts for failed emails, including the number of retries and the delay between them.  
 *Logging & Reporting*: Creates a detailed log file (`log_envios.csv`) for all transactions (successes and failures) and allows exporting filtered reports.  
 *Success Backup*: Creates an organized, folder-based backup of all successfully sent files.  
 *Intuitive GUI*: Features a tabbed interface for basic and advanced settings, with visual feedback on progress and sending statistics.
