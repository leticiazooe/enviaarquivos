🧠 Tecnologias Utilizadas
Python: Linguagem principal do projeto, escolhida pela sua versatilidade e poder de automação.
Tkinter: Usada para criar uma interface gráfica interativa (GUI) para facilitar o uso por colaboradores não técnicos.
pandas: Para leitura e manipulação dos dados de uma planilha Excel (.xlsx).
smtplib e email.mime: Bibliotecas nativas para envio de e-mails via protocolo SMTP, com suporte a corpo de texto e anexos.
openpyxl: Motor de leitura de planilhas Excel.
datetime e os: Manipulação de tempo e diretórios para geração e armazenamento de logs personalizados.

⚙️ Funcionalidades Desenvolvidas
Leitura dinâmica de planilhas Excel com nomes e e-mails dos destinatários.
Envio de e-mails personalizados com substituição dinâmica do nome {nome} no assunto e corpo do e-mail.
Suporte a múltiplos anexos, com extensão restrita a tipos de documentos úteis no ambiente corporativo (.pdf, .docx, .pptx, .xls, .pbix etc).
Barra de progresso para feedback visual durante o envio em lote.
Log de execução salvo automaticamente com:
Status por destinatário (sucesso ou erro),
Horário de envio,
Pasta customizável para armazenar os logs.
Mensagens em tempo real na interface informando o status de cada envio.
Interface intuitiva e limpa, permitindo:
Seleção de planilha,
Edição do texto e assunto do e-mail,
Inserção e remoção de anexos,
Escolha e abertura da pasta de log.

🤝 Soft Skills e Boas Práticas Demonstradas
Empatia com o usuário: Escutou a necessidade real dos colaboradores e adaptou a solução.
Escalabilidade: Pensou em uma solução que pode ser reaproveitada por outros times.
Documentação visual: Feedback claro na interface e log externo para auditoria.
Automação inteligente: Redução de retrabalho manual com foco na produtividade.
Segurança e boas práticas: Controle dos tipos de anexos permitidos e registros de log.
