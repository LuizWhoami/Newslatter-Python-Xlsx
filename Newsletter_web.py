import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
import pandas as pd

print('Bem vindo ao Tech-Loom')

# Email and web scraping settings
web = requests.get("https://portogente.com.br/noticias-corporativas/115348-melhor-curso-python")
content = web.content
site = BeautifulSoup(content, "html.parser")

# Corpo do email
corpo = """ <h2><strong>Quais as vantagens de fazer um curso de python?</strong></h2>
<p>Fazer um curso de Python oferece várias vantagens. Os cursos fornecem um plano de estudos organizado e abrangente, ajudando você a entender pra que serve o python, como aprender de maneira sequencial e progressiva.</p>
<p>A orientação de instrutores experientes facilita a compreensão de conceitos complexos e o esclarecimento de dúvidas.</p>
<p>Além disso, o ambiente de aprendizado colaborativo proporciona a troca de ideias e experiências com outros alunos. O suporte oferecido pelos cursos também ajuda a manter a motivação e o comprometimento, melhorando a retenção do conhecimento.</p>
<p>Por fim, muitos cursos oferecem certificados, o que pode ser útil para comprovar suas habilidades e aumentar a empregabilidade.</p>


<h2><strong>5 Melhores cursos de Python (Gratuitos e Pagos)</strong></h2>

<ol>
    <li><a href="https://cursos.dankicode.com/curso-python-completo?ref=Y80526592L?src=orderedlist_portogente" target="_blank" rel="noopener noreferrer">Curso de Python Mais Completo e Melhor Avaliado</a> (Danki Code)</li>
    <li><a href="https://go.hotmart.com/T80393383K?src=orderedlist_portogente" target="_blank" rel="noopener noreferrer">Curso de Python Com o Melhor Custo Benefício</a> (Expert Cursos)</li>
    <li><a href="https://go.hotmart.com/C80526597D?src=orderedlist_portogente" target="_blank" rel="noopener noreferrer">Python para Iniciantes do Zero ao Primeiro Software</a> (Wagner Cardoso)</li>
    <li><a href="https://www.udemy.com/course/python-3-do-zero-ao-avancado/" target="_blank" rel="noopener noreferrer">Curso de Python 3 do Básico Ao Avançado Com Projetos Reais</a> (Udemy)</li>
    <li><a href="https://didatica.tech/curso-de-python-online-para-iniciantes/" target="_blank" rel="noopener noreferrer">Curso de Python Grátis e Online Para Iniciantes</a> (Didática Tech)</li>
    </ol>

<h2>TechLoom</h2>
 Desenvolvedor:
<a href="https://github.com/Dvisonlui">Github</a>

<h3>Informações</h3>
<p>Sempre lançadas as 6:00 horas da manhã</p>
<p>fiquem atentos a seu E-mail</p>"""  # Insira aqui o conteúdo do corpo do seu email

# SMTP server settings
server_smtp = "smtp.gmail.com"
porta = 587
quem_envia = "dvison.ferreir@gmail.com"
senha = "gurwjhsfuzqzfvza"

# Leitura da planilha Excel
planilha = pd.read_excel('clientes.xlsx')

# Conectando ao servidor SMTP
try:
    server = smtplib.SMTP(server_smtp, porta)
    server.starttls()
    server.login(quem_envia, senha)

    # Iterar sobre os destinatários e enviar e-mails personalizados
    for index, row in planilha.iterrows():
        recebe = row['E-mail']
        nome = row['nome']
        assunto = "Newsletter - TechLoom"

        mensagem = MIMEMultipart()
        mensagem["From"] = quem_envia
        mensagem["To"] = recebe
        mensagem["Subject"] = assunto
        mensagem.attach(MIMEText(corpo, "html"))

        server.sendmail(quem_envia, recebe, mensagem.as_string())
        print(f"Email enviado para {nome} ({recebe})")

except Exception as e:
    print(f"Deu erro: {e}")
finally:
    server.quit()
