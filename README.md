# Piloto

Este projeto é um robô de automação desenvolvido em Python que realiza agendamentos automáticos no sistema SIRESP (Sistema de Regulação Eletrônica de Saúde Pública de São Paulo). Ele foi criado para otimizar o processo de marcação de consultas e exames, reduzindo erros manuais e economizando tempo.
------------------------------------------------------------------------------------------------------------------------------------
🚀 Funcionalidades Principais

✔ Automação de Agendamentos

Realiza agendamentos de consultas e exames no SI RESP.

Seleciona automaticamente especialidades, profissionais e horários disponíveis.

✔ Tratamento de Dados

Importa planilhas Excel com dados de pacientes (código, especialidade, profissional, observações).

Gera relatórios em Excel com histórico de agendamentos.

✔ Gestão de Erros

Identifica pacientes já agendados.

Detecta ausência de vagas e notifica o usuário.

✔ Automação Inteligente

Para pacientes com "Coleta" na observação, agenda também o exame laboratorial com intervalo de 15-20 dias após a consulta.
----------------------------------------------------------------------------------------------------------------------------------
⚙️ Tecnologias Utilizadas
Python (Linguagem principal)

Selenium (Automação web)

Pandas (Manipulação de dados)

CustomTkinter (Interface gráfica moderna)

JSON (Armazenamento de histórico)
-------------------------------------------------------------------------------------------------------------------------------------
📊 Fluxo do Sistema
Login no SIRESP (via automação Selenium).

Importação da planilha (seleção via interface gráfica).

Processamento dos agendamentos:

Seleciona especialidade, profissional e data disponível.

Se houver "Coleta" na observação, agenda exame laboratorial.

Geração de relatórios (Excel + JSON com histórico).
---------------------------------------------------------------------------------------------------------------------------------------

📧 duetum@hotmail.com
🔗 LinkedIn: https://www.linkedin.com/in/matheus-augusto-3a1152289/
