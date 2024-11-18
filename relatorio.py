import pyodbc
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from datetime import datetime
import config  # Importa as configurações

# Conexão ao banco de dados usando as configurações do arquivo config.py
conn_string = config.DB_CONNECTION_STRING

try:
    conn = pyodbc.connect(conn_string)
    print("Conectado ao banco de dados!")
except pyodbc.Error as ex:
    print(f"Erro ao conectar: {ex}")

# Criar um cursor para realizar as consultas
cursor = conn.cursor()

# Consultas SQL
query1 = """
SELECT
    CONVERT(VARCHAR(10), ultimaMovimentacao, 103) AS DATA,
    nomeProcesso AS TIPO
FROM redoc_tramitacao
WHERE progresso = 100
    AND (usuarioOrigem = '9076' OR usuarioOrigem = '9061' OR usuarioOrigem = '9077')
    AND YEAR(ultimaMovimentacao) = YEAR(GETDATE())
    AND MONTH(ultimaMovimentacao) = MONTH(GETDATE())
    AND DAY(ultimaMovimentacao) = DAY(GETDATE());
"""

# Executando a consulta e armazenando os resultados em um DataFrame
df1 = pd.read_sql(query1, conn)
print("\nFINALIZADAS HOJE:")
print(df1)

query2 = """
SELECT
    cliente AS CPF,
    nome_cliente AS PARTICIPANTE,
    nomeProcesso AS TIPO,
    ultimaMovimentacao
FROM redoc_tramitacao AS rt
INNER JOIN redoc_documento AS rd
    ON rt.cdDoc = rd.id
WHERE progresso = 10
    AND rt.cdProcesso = cd_processo
    AND (nomeProcesso IN (
         'Informações complementares a inscrição automática',
         'Mudança de Patrocinador',
         'Alteração Cadastral',
         'Ficha Cadastral',
         'Opção pelo regime de Tributação',
         'Termo Especial de Cadastro de Pessoa Politica Exposta',
         'Alteração de Percentual e Base de Contribuição',
         'Cancelamento de Inscrição',
         'Devolução de Contribuição'))
    AND YEAR(ultimaMovimentacao) = YEAR(GETDATE())
ORDER BY tipo;
"""

# Executando a consulta e armazenando os resultados em um DataFrame
df2 = pd.read_sql(query2, conn)
print("\nPENDENTES:")
print(df2)

# Fechar o cursor e a conexão
cursor.close()
conn.close()

# Contagem de registros
finalizadas_count = len(df1) if not df1.empty else 0
pendentes_count = len(df2) if not df2.empty else 0

# Processamento dos atrasos
df2['ultimaMovimentacao'] = pd.to_datetime(df2['ultimaMovimentacao'])
df2['dias_atraso'] = (datetime.now() - df2['ultimaMovimentacao']).dt.days

# Filtrando e criando intervalos
df2_atraso = df2[df2['dias_atraso'] > 5]

# Gerar gráfico
sns.set_theme(style="darkgrid")
if not df2_atraso.empty:
    atraso_por_dia = df2_atraso['dias_atraso'].value_counts().sort_index()

    def definir_cor(dias):
        if dias > 20:
            return 'red'
        elif dias > 15:
            return 'yellow'
        elif dias > 10:
            return 'blue'
        else:
            return 'green'

    cores = [definir_cor(dias) for dias in atraso_por_dia.index]

    plt.figure(figsize=(12, 7))
    ax = atraso_por_dia.plot(kind='bar', color=cores, edgecolor='black')
    plt.title('Farol - Dias de atraso', fontsize=16, fontweight='bold')
    plt.xlabel(' ', fontsize=14)

    dias_labels = [f"{dias} dia" if dias == 1 else f"{dias} dias" for dias in atraso_por_dia.index]
    ax.set_xticks(range(len(atraso_por_dia)))
    ax.set_xticklabels(dias_labels, rotation=0, fontsize=12)
    ax.set_yticks([])

    for i, v in enumerate(atraso_por_dia):
        ax.text(i, v + 0.5, f"{v}", ha='center', fontsize=12, fontweight='bold', color='black')
else:
    plt.figure(figsize=(12, 7))
    plt.text(0.5, 0.5, 'Nenhuma ficha em atraso', horizontalalignment='center', verticalalignment='center', fontsize=14)
    plt.title('Quantidade de Fichas em Atraso', fontsize=16, fontweight='bold')

plt.tight_layout()
image_buffer = BytesIO()
plt.savefig(image_buffer, format='png', dpi=300)
plt.close()
image_buffer.seek(0)

# E-mail
data_atual = datetime.now().strftime('%d/%m/%Y')
smtp_server = config.SMTP_SERVER
smtp_port = config.SMTP_PORT
email = config.SMTP_EMAIL
senha = config.SMTP_PASSWORD
destinatario = config.SMTP_RECIPIENT
assunto = f'Formulários processados - Portal do patrocinador {data_atual}'

corpo = f"""<html>
<body>
<p>Prezados, boa tarde! <br><br>

Encaminho abaixo o relatório referente aos formulários processados no Portal do Patrocinador, para fins de ajustes e atualizações cadastrais, realizados em {data_atual}.</p>

<p><strong>FINALIZADAS HOJE (TOTAL - {finalizadas_count}):</strong></p>
<pre>{df1.to_string(index=False)}</pre>

<p><strong>PENDENTES (TOTAL - {pendentes_count}):</strong></p>
<pre>{df2.to_string(index=False)}</pre>

<p><strong>GRÁFICO DE FICHAS EM ATRASO</strong></p>
<img src="cid:imagem_atraso" alt="Gráfico de Atraso">
</body>
</html>
"""

mensagem = MIMEMultipart()
mensagem['From'] = email
mensagem['To'] = destinatario
mensagem['Subject'] = assunto
mensagem.attach(MIMEText(corpo, 'html'))

imagem = MIMEImage(image_buffer.read())
imagem.add_header('Content-ID', '<imagem_atraso>')
mensagem.attach(imagem)

# Enviando o e-mail
try:
    servidor = smtplib.SMTP(smtp_server, smtp_port)
    servidor.starttls()
    servidor.login(email, senha)
    servidor.send_message(mensagem)
    print('E-mail enviado com sucesso!')
except Exception as e:
    print(f'Erro ao enviar e-mail: {e}')
finally:
    servidor.quit()
