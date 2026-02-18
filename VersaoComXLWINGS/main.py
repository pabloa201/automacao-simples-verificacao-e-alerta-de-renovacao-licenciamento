import smtplib
from email.message import EmailMessage
import xlwings as xw

wbxl = xw.Book('controleRenovacao.xlsx')
# Celulas e Aba da Planilha de Licenças
nomeColuna = wbxl.sheets['DiasParaVencer'].range('C1').value
licencaIndividualTeams = int(wbxl.sheets['DiasParaVencer'].range('C2').value)
licenca16Teams = int(wbxl.sheets['DiasParaVencer'].range('C3').value)
licenca18Teams = int(wbxl.sheets['DiasParaVencer'].range('C4').value)
licencaBackupExec = int(wbxl.sheets['DiasParaVencer'].range('C5').value)
licencaTeamViewer = int(wbxl.sheets['DiasParaVencer'].range('C6').value)
licencaOracle = int(wbxl.sheets['DiasParaVencer'].range('C8').value)
# Celulas e Aba da Planilha de DiasParaVencer
nomeColuna2 = wbxl.sheets['DiasParaVencer'].range('C1').value
CertificadoA1 = int(wbxl.sheets['DiasParaVencer'].range('C9').value)
CertificadoRemoteApp = int(wbxl.sheets['DiasParaVencer'].range('C10').value)
# Celulas e Aba da Planilha de DiasParaVencer
nomeColuna3 = wbxl.sheets['DiasParaVencer'].range('C1').value
Domain1 = int(wbxl.sheets['DiasParaVencer'].range('C11').value)
Domain2 = int(wbxl.sheets['DiasParaVencer'].range('C12').value)
Domain3 = int(wbxl.sheets['DiasParaVencer'].range('C13').value)
Domain4 = int(wbxl.sheets['DiasParaVencer'].range('C14').value)

# Impressões de mensagem dos resultados
print('Licenciamento/Assinaturas: ' + nomeColuna)
print('Faltam ' + str(licencaIndividualTeams) + ' dias para vencer a licença individual do Teams')
print('Faltam ' + str(licenca16Teams) + ' dias para vencer as 16 licenças do Teams ')
print('Faltam ' + str(licenca18Teams) + ' dias para vencer as 18 licenças do Teams ')
print('Faltam ' + str(licencaBackupExec) + ' dias para vencer a licença do BackupExec ')
print('Faltam ' + str(licencaTeamViewer) + ' dias para vencer a licença do Team Viewer')
print('Faltam ' + str(licencaOracle) + ' dias para vencer a licença do Oracle ')
print()
print('DiasParaVencer: ' + nomeColuna2)
print('Faltam ' + str(CertificadoA1) + ' dias para vencer o certificado A1-empresa')
print('Faltam ' + str(CertificadoRemoteApp) + ' dias para vencer o Certificado do Remote App TS4 e TS2')
print()
print('DiasParaVencer: ' + nomeColuna3 + ' no RegistroBR')
print('Faltam ' + str(Domain1) + ' dias para vencer o dominio domain1.com.br')
print('Faltam ' + str(Domain2) + ' dias para vencer o dominio Domain2.com.br')
print('Faltam ' + str(Domain3) + ' dias para vencer o dominio Domain3.com.br')
print('Faltam ' + str(Domain4) + ' dias para vencer o dominio Domain4.com.br')

file = open('LogDiarioStatusLicenciamento.txt', 'w')
file.write('\nLicenciamento/Assinaturas: ' + nomeColuna + '\n')
file.write('Faltam ' + str(licencaIndividualTeams) + ' dias para vencer a licença individual do Teams\n')
file.write('Faltam ' + str(licenca16Teams) + ' dias para vencer as 16 licenças do Teams \n')
file.write('Faltam ' + str(licenca18Teams) + ' dias para vencer as 18 licenças do Teams \n')
file.write('Faltam ' + str(licencaBackupExec) + ' dias para vencer a licença do BackupExec \n')
file.write('Faltam ' + str(licencaTeamViewer) + ' dias para vencer a licença do Team Viewer\n')
file.write('Faltam ' + str(licencaOracle) + ' dias para vencer a licença do Oracle')
file.write('\n')
file.write('\nDiasParaVencer: ' + nomeColuna2 + '\n')
file.write('Faltam ' + str(CertificadoA1) + ' dias para vencer o certificado A1-empresa\n')
file.write('Faltam ' + str(CertificadoRemoteApp) + ' dias para vencer o Certificado do Remote App TS4 e TS2\n')
file.write('\n')
file.write('DiasParaVencer: ' + nomeColuna3 + ' no RegistroBR\n')
file.write('Faltam ' + str(Domain1) + ' dias para vencer o dominio domain1.com.br\n')
file.write('Faltam ' + str(Domain2) + ' dias para vencer o dominio Domain2.com.br\n')
file.write('Faltam ' + str(Domain3) + ' dias para vencer o dominio Domain3.com.br\n')
file.write('Faltam ' + str(Domain4) + ' dias para vencer o dominio Domain4.com.br\n')
file.close()

# AutenticaçãoEmail
email_from = 'emailremetente@email.com.br'
email_to = 'emaildestino@email.com.br'
smtp = 'mail.dominio.com.br'
senha = 'password'
# licenças
if (licencaIndividualTeams == 40 or licencaIndividualTeams == 30 or licencaIndividualTeams == 20 or
        licencaIndividualTeams == 15 or licencaIndividualTeams == 5 or licencaIndividualTeams == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Faltam ' + str(licencaIndividualTeams) + ' dias para vencer a licença individual do Teams')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (licenca16Teams == 40 or licenca16Teams == 30 or licenca16Teams == 20 or
        licenca16Teams == 15 or licenca16Teams == 5 or licenca16Teams == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Em ' + str(licenca16Teams) + ' dia(as) expiram as 16 licenças do Teams ')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (licenca18Teams == 40 or licenca18Teams == 30 or licenca18Teams == 20 or
        licenca18Teams == 15 or licenca18Teams == 5 or licenca18Teams == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Em ' + str(licenca18Teams) + ' dia(as) expiram as 18 licenças do Teams ')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (licencaBackupExec == 40 or licencaBackupExec == 30 or licencaBackupExec == 20 or
        licencaBackupExec == 15 or licencaBackupExec == 5 or licencaBackupExec == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Em ' + str(licencaBackupExec) + ' dia(as) expira a licença do BackupExec ')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (licencaTeamViewer == 40 or licencaTeamViewer == 30 or licencaTeamViewer == 20 or
        licencaTeamViewer == 15 or licencaTeamViewer == 5 or licencaTeamViewer == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Sua licença de Team Viewer expira em ' + str(licencaTeamViewer) + ' dia(as)')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (licencaOracle == 40 or licencaOracle == 30 or licencaOracle == 20 or
        licencaOracle == 15 or licencaOracle == 5 or licencaOracle == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Em ' + str(licencaOracle) + ' dia(as) vence o suporte do Oracle')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

# DiasParaVencer
if (CertificadoA1 == 40 or CertificadoA1 == 30 or CertificadoA1 == 20 or
        CertificadoA1 == 15 or CertificadoA1 == 5 or CertificadoA1 == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Em ' + str(CertificadoA1) + ' dias vence o certificado A1-empresa')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (CertificadoRemoteApp == 40 or CertificadoRemoteApp == 30 or CertificadoRemoteApp == 20 or
        CertificadoRemoteApp == 15 or CertificadoRemoteApp == 5 or CertificadoRemoteApp == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Em ' + str(CertificadoRemoteApp) + ' dias vence o certificado do Remote App TS4 e TS2 ')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

# DiasParaVencer no Registro BR
if (Domain1 == 40 or Domain1 == 30 or Domain1 == 20 or
        Domain1 == 15 or Domain1 == 5 or Domain1 == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content(
        'Em ' + str(Domain1) + ' dias vence o registro do dominio domain1.com.br no RegistroBR')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (Domain2 == 40 or Domain2 == 30 or Domain2 == 20 or
        Domain2 == 15 or Domain2 == 5 or Domain2 == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content('Em ' + str(Domain2) + ' dias vence registro do dominio Domain2.com.br no RegistroBR')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (Domain3 == 40 or Domain3 == 30 or Domain3 == 20 or
        Domain3 == 15 or Domain3 == 5 or Domain3 == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content(
        'Em ' + str(Domain3) + ' dias vence o registro do dominio Domain3.com.br no RegistroBR')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

if (Domain4 == 40 or Domain4 == 30 or Domain4 == 20 or
        Domain4 == 15 or Domain4 == 5 or Domain4 == 1):
    server = smtplib.SMTP(smtp, 587)
    server.starttls()
    server.login(email_from, senha)
    msg = EmailMessage()
    msg['Subject'] = 'Alerta Status Licenciamentos'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content(
        'Em ' + str(Domain4) + ' dias vence o registro do dominio Domain4.com.br no RegistroBR')
    server.send_message(msg)
    server.quit()
    print('Email Enviado')

wbxl.close()
