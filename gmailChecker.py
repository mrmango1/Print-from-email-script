import os, sys, imaplib, email, getopt, time
from subprocess import Popen, DEVNULL
from imbox import Imbox  
import traceback

 # If you want yo use gmail
host = 'imap.gmail.com'
# GMail username/password
username = 'email'
password = 'password'
# download_path
download_folder = "attachments_folder"
class MailCheker():

    def __init__(self, quiet=False):
        
        # Intervalo de busqueda
        self.mail_check_interval = 15
        
        self.inbox_read_only = True
        
        # The mail connection handler
        self.mailconn = None
        
        # Should we print any output?
        self.quiet = quiet
    
    ########################################################################
    
    # Conectando a GMail
    def connect(self):
        self.mailconn = imaplib.IMAP4_SSL(host)
        (retcode, capabilities) = self.mailconn.login(username, password)
        self.mailconn.select('inbox', readonly=self.inbox_read_only)
    
    # Check for mail
    def check_mail(self):
        # Connect to GMail (fresh connection needed every time)
        self.connect()
    
        num_new_messages=0
        (retcode, messages) = self.mailconn.search(None, '(UNSEEN Subject "Imprimir")')
        if retcode == 'OK':
    
           for num in messages[0].split() :
              num_new_messages += 1
              if not self.quiet:
                  print('  Processing #{}'.format(num_new_messages))
              typ, data = self.mailconn.fetch(num,'(RFC822)')
              for response_part in data:
                 if isinstance(response_part, tuple):
                     original = email.message_from_bytes(response_part[1])
                     if not self.quiet:
                         print("   ", original['From'])
                         print("   ", original['Subject'])
                     typ, data = self.mailconn.store(num,'+FLAGS','\\Seen')
    
        return num_new_messages

def gmailPrint():
    if not os.path.isdir(download_folder):
        os.makedirs(download_folder, exist_ok=True)

    mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
    messages = mail.messages(unread=True, subject='Imprimir') # por defecto en inbox
    print("Buscando archivos adjuntos para imprimir ")
    
    for (uid, message) in messages:
        mail.mark_seen(uid) # marcar mensaje como leido
    
        for idx, attachment in enumerate(message.attachments):
            try:
                att_fn = attachment.get('filename')
                download_path = f"{download_folder}/{att_fn}"
                print(f'Imprimiendo {att_fn}')
                # os.system(f'lpr {download_path}')
                # print(download_path)
                with open(download_path, "wb") as fp:
                    fp.write(attachment.get('content').read())
            except:
                print(traceback.print_exc())
    
    else:
        print("\nNo se encontraron archivos para imprimir ")
    mail.logout()

    
    ########################################################################
    
if __name__ == '__main__':

    # Specify -q for quiet operation (no output)
    opts, args = getopt.getopt(sys.argv[1:], 'q')
    quiet = False
    for opt, arg in opts:
        if opt == '-q':
            quiet = True

    while True:
        if not quiet:
            print("Checking... ", end='')
        num_new_messages = mb.check_mail()
        if not quiet:
            print("{} new messages".format(num_new_messages))
        if num_new_messages > 0:
            gmailPrint()
        time.sleep(mb.mail_check_interval)
