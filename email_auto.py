import pandas as pd
import datetime as DT
import email
import locale
import smtplib

### Configurar o local, para o uso futuro da função currency.
locale.setlocale(locale.LC_ALL, "en_US.UTF-8")
#locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8") ## Caso tenha algum valor monetario em real, por favor descomente essa linha e comente a linha anterior

### Criação de uma função para a tranformação de algumas colunas para a formatação de moeda.
def moeda(df, colunas: list):
    try:
        for i in colunas:
            df[i] = df[i].astype(float)
            df[i] = df.apply(
                lambda x: locale.currency(x[i], grouping=True, symbol=True), axis=1
            )
        return df
    except:
        print(f"A COLUNA {i} APRESENTA ALGUM ERRO")

### Leitura do excel
df = pd.read_excel('sampledatafoodsales.xlsx',sheet_name='FoodSales')

### Formatação dos dados
df = moeda(df=df,colunas=['TotalPrice','UnitPrice'])
df['Qty'] = df['Qty'].map('{:n}'.format)

### Tranforção da coluna para data
df["Date"] = df["Date"].apply(lambda s: DT.datetime.strptime(s, "%Y-%m-%d"))

### Filtragem do dataframe para o dia desejado
#df = df[df["Date"] == DT.datetime.today().strftime('%Y-%m-%d')] ## Caso sua base de dados tenha valores futuros, descomente essa linha e comente a linha de baixo, antes de colocar seu codigo no agendador de tarefas.
df = df[df["Date"] == '2022-01-04']

df.drop(columns=['ID', 'Date'])

### Tornando o dataframe filtrado em uma versão html  
df_html = df.to_html(index=False,justify='center')
#print(df_html)

### Formnatndo a tabela em html
df_html = df_html.replace('<table border="1" class="dataframe">','<table border="1" class="dataframe" style="font-size: 10pt; border: 1px solid black; border-collapse: collapse; text-align: center; width: 100%;">')
################# EMAIL  #################

### Criação das duas versões de email que serão utilizados
mensagem_erro = f"""
            <center>
                <p>Bom dia,</p>
                <p>Não consta nenhuma entega para o dia de hoje.</p>

                <p>Caso seja encontrado algum erro, favor informar a equipe de dados. </p>

                <p>Atenção, essa é uma mensagem automática.Por favor, Não responda.</p>
                <div style="margin-bottom: 20px;">
                <img src="https://seeklogo.com/images/G/generic-logo-ECC6ED04F3-seeklogo.com.png" width="300" style="margin-right: 20px;">
                </div>
            </center>
            """
mensagem_sucesso = f"""
            <center>
                <p>Bom dia,</p>
                <p>Venho por meio deste, deixa-los ciente que as seguintes entregas estão marcadas para hoje.</p>

                <p>Aqui está a tabela:</p>
                
                <style> 
                    table, th, td {{font-size:10pt; border:1px solid black; border-collapse:collapse; text-align:left;}}
                    th, td {{padding: 5px;}}
                </style>
                {df_html}

                <p>Caso seja encontrado algum erro, favor informar a equipe de dados. </p>

                <p>Atenção, essa é uma mensagem automática.Por favor, Não responda.</p>
                <div style="margin-bottom: 20px;">
                <img src="https://seeklogo.com/images/G/generic-logo-ECC6ED04F3-seeklogo.com.png" width="300" style="margin-right: 20px;">
                </div>
            </center>
            """

subject = f"As entregas do dia {DT.datetime.today().strftime('%Y-%m-%d')}"
#print(mensagem_sucesso)

enderecos_email = ['seu_email1@email.com','seu_email2@email.com']

### Logica utilizada para o envio de email
if df.empty:
    for emails in enderecos_email:
        msg = email.message.Message()
        msg['Subject'] = subject
        msg['From'] = 'seu_email1@email.com'
        msg['To'] = emails
        password = "sua_senha1"
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(mensagem_erro)

        s = smtplib.SMTP('smtp.gmail.com',587)
        s.ehlo()
        s.starttls()

        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('ISO-8859-1'))
        s.close()
else:
    for emails in enderecos_email:
        msg = email.message.Message()
        msg['Subject'] = subject
        msg['From'] = 'seu_email1@email.com'
        msg['To'] = emails
        password = "sua_senha1"
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(mensagem_sucesso)

        s = smtplib.SMTP('smtp.gmail.com',587)
        s.ehlo()
        s.starttls()

        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('ISO-8859-1'))
        s.close()
