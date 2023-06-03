#from urllib.request import urlopen, Request  #importation module gestion url alternatif
import win32com.client as win32  #importation module gestion boite mail
import time  #importation module temps
from inputimeout import inputimeout, TimeoutOccurred
from time import strftime
from tqdm import tqdm

#alternative avec urlib
#headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.3'}     #authentification du browser (eviter erreur 403)
#reg_url = 'https://xxxxxxxx.com' #url a rechercher
#req = Request(url=reg_url, headers=headers) #requete
#html = urlopen(req).read()  #premiere ouverture de l'url pour trouver le dernier chapitre

import requests
page = requests.get("https://xxxxxxxx.com") #requête sur la page
html = page.text    #conversion de la page html en text

#initialisation de variables
j = 0
k = 0
der_chap=1

print("\n"+"Recherche du dernier chapitre, veuillez patienter..."+"\n")

while j != -1:
    der_chap=der_chap+1
    j = (html).find("Chapter "+str(der_chap))

for i in tqdm(range(der_chap)):
    time.sleep(0.01)
    pass

time.sleep(0.2)
print("\n"+"\033[32m"+"Trouvé \u2713")
time.sleep(1)
print("\033[0m")
print("Le dernier chapitre disponible est le " + str(der_chap-1) + (". Nous vous préviendrons quand le chapitre " + str(der_chap) + " sera disponible."+"\n")) 


trouve = 0  #initialisation de la condition de sortie a 0 pour rester dans la boucle


while trouve == 0:  #debut de la boucle

    #l'ecriture dans un fichier n'ai absolument pas necessaire ici c'est juste pour s'amuser

    with open('data.txt','w+') as fichier:  #ouverture d'un fichier pour stocker la page avec with pour pas oublier de le refermer apres utilisation

        debut = (html).find('class="page-content-listing single-page"')  #condition de debut decriture dans notre fichier, on veut limiter les donnees a verifier
        fin = (html).find('class="c-chapter-readmore"')  #condition de fin d'ecriture          

        for i in range(debut,fin):  #boucle entre les deux conditions
            fichier.write((html)[i]) #on ecrit dans le fichier sous forme de text le contenu de notre page web entre le debut et la fin 

        fichier.seek(0) #on replace notre pointeur dans notre fichier au debut
        
        if fichier.read().find('Chapter ' + str(der_chap)) != -1: #lecture du fichier et recherche de l'emplacement de 'Chapter $$' dedans, si pas de $$ la fonction find retourne -1

            print(strftime("%d/%m/%Y - %Hh%M") + ' :' + '\033[35m' + ' Nouveau chapitre disponible !' + '\033[0m')    #si presence de 'Chapter $$' le programme indique qu'il l'a trouve

            outlook = win32.Dispatch('outlook.application') #acces a l'application outlook du systeme
            mail = outlook.CreateItem(0)    #creation d'un objet mail
            mail.To = 'xxxxxxxx@gmail.com'  #destinataire du mail
            mail.Subject = 'Nouveau chapitre de ?????'  #objet du mail
            mail.Body = 'Un nouveau chapitre de ????? est disponible !'   #contenu du mail
            mail.Send() #envoi du mail

            time.sleep(180)

            trouve = 1  #mise a jour de la condition de sortie de la boucle

        else:   
            print(strftime("%d/%m/%Y - %Hh%M") + ' : Pas de nouveau chapitre'+"\n") #si pas de "Chapter $$" le programme l'indique
            
    for u in range(44):
        print('\033[33m'+".",end='')
        time.sleep(1.1)

    print("\n"+'\033[0m')

    if k == 0 :
        print('\n'+"Si vous voulez arreter la boucle, vous avez 10 secondes pour taper 'stop' à la suite de '>>>'.")
        print("La possibilité d'arrêter la boucle vous sera reproposée à chaque tour de boucle."+'\n')
        time.sleep(2)
        k=k+1 

    try:
        fin = inputimeout(prompt='>>>', timeout=10)
    except TimeoutOccurred:
        fin = "Vous n'avez pas tapé 'stop', la boucle va continuer."
    if fin == "stop" : 
        print('\n'+"Le programme va s'arrêter, merci d'avoir utilisé ce script!")
        time.sleep(4)
        exit()
    else :
        print('\n'+fin+'\n')

