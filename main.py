import csv, ebooklib, pypandoc
from os import path, get_terminal_size
from colorama import init, Fore
from unidecode import unidecode
from ebooklib import epub
def main():
    """
    Just all the program
    """
    width = get_terminal_size()[0]
    barlenght = width-75
    try:
        nomfichier = input("Quel fichier ouvrir ? >> ")
    except KeyboardInterrupt:
        return None
    extension = nomfichier.split('.',-1)[-1]
    if extension == 'txt':
        try:
            file = open(nomfichier)
        except FileNotFoundError:
            print("/!\\ Fichier non trouvé")
            return None
        print("Ouverture du fichier "+nomfichier+"\nExtraction des données")
        rawtext = file.read()
        file.close()
        print("Extraction faite")
    elif extension == 'epub':
        rawtext = ""
        try:
            book = epub.read_epub(nomfichier,{'ignore_ncx':True})
        except FileNotFoundError:
            print("/!\\ Fichier non trouvé")
            return None
        print("Ouverture du fichier "+nomfichier+"\nExtraction des données")
        length=0
        for doc in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
            length+=len(unidecode(doc.content.decode('utf8')))
        actuallength=0
        barlenght=width-((len(str(length))*2)+38)
        for doc in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
            newstring=""
            passerchar = False
            toparse = unidecode(doc.content.decode('utf8'))
            for char in toparse:
                actuallength+=1
                if not passerchar:
                    if char == "<":
                        passerchar = True
                    else:
                        newstring+=char
                elif char == ">":
                    passerchar = False
                print(str('\r'+Fore.WHITE+'['+Fore.GREEN+int(round(((actuallength+1)/length*barlenght),0)-1)*'='+'>'+Fore.RESET+int(abs(round(((actuallength+1)/length*barlenght),0)-barlenght))*' '+Fore.WHITE+']'+Fore.RESET+' >> Caractère n'+Fore.YELLOW+str(actuallength+1)+Fore.RESET+' sur '+Fore.YELLOW+str(length)+Fore.RESET+', '+Fore.CYAN+str(round(((actuallength+1)/length*100),1))+Fore.RESET+'% fait'),end='',flush=True)
            rawtext+="\n"+newstring
        print("\nExtraction faite")
    elif extension == 'docx' or extension == 'odt':
        newfile=nomfichier.split('.',1)[0]+'.txt'
        if not path.isfile(nomfichier):
            print("/!\\ Fichier non trouvé")
            return None
        print("Ouverture du fichier "+nomfichier+"\nExtraction des données")
        try:
            output = pypandoc.convert_file(nomfichier, 'plain', outputfile=newfile)
            file = open(newfile)
        except FileNotFoundError:
            print("/!\\ Fichier non trouvé")
            return None
        rawtext = file.read()
        file.close()
        print("Extraction faite")
    text = rawtext.split("\n",-1)
    print("Il y a "+Fore.YELLOW+str(len(text))+Fore.RESET+" lignes")
    words = []
    for row in text:
        for word in row.split(" ",-1):
            words.append(word)
    print("Filtrage en cours")
    for i in range(len(words)):
        words[i]=unidecode(words[i]).lower()
        newstring=""
        for char in words[i]:
            if not(char in '/[^£$%&*()}{@~?><>,|=_+¬-]/.!;'):
                newstring+=char
        words[i]=newstring
        print(str('\r'+Fore.WHITE+'['+Fore.GREEN+int(round(((i+1)/len(words)*barlenght),0)-1)*'='+'>'+Fore.RESET+int(abs(round(((i+1)/len(words)*barlenght),0)-barlenght))*' '+Fore.WHITE+']'+Fore.RESET+' >> Mot numero '+Fore.YELLOW+str(i+1)+Fore.RESET+' sur '+Fore.YELLOW+str(len(words))+Fore.RESET+', '+Fore.CYAN+str(round(((i+1)/len(words)*100),1))+Fore.RESET+'% fait'),end='',flush=True)
        #print('\rMot numero '+str(i+1)+' sur '+str(len(words))+', '+str(round(((i+1)/len(words)*100),1))+'% fait', end='', flush=True)
    print("\nFiltrage fait\nIl y a "+Fore.YELLOW+str(len(words))+Fore.RESET+" mots")
    dico = {}
    print("Classification en cours")
    for iword in range(len(words)):
        if len(words[iword]) != 1:
            trove = False
            if words[iword] in dico.keys():
                    trove = True
                    dico[words[iword]]+=1
            if not trove:
                dico[words[iword]]=1
            print(str('\r['+Fore.GREEN+int(round(((iword+1)/len(words)*barlenght),0)-1)*'='+'>'+Fore.RESET+int(abs(round(((iword+1)/len(words)*barlenght),0)-barlenght))*' '+'] >> Mot numero '+Fore.YELLOW+str(iword+1)+Fore.RESET+' sur '+Fore.YELLOW+str(len(words))+Fore.RESET+', '+Fore.CYAN+str(round(((iword+1)/len(words)*100),1))+Fore.RESET+'% fait'),end='',flush=True)
            #print('\rMot numero '+str(iword+1)+' sur '+str(len(words))+', '+str(round(((iword+1)/len(words)*100),1))+'% fait', end='', flush=True)
    print("\nClassification faite")
    resultatfile = str('words-'+nomfichier.split(".",1)[0]+'.csv')
    if not path.isfile(resultatfile):
        print('Fichier inexistant, création de '+resultatfile)
    else:
        print("Fichier existant, création d'un duplicata")
        i = 1
        resultatfile = str('words-'+nomfichier.split(".",1)[0]+str(i)+'.csv')
        while path.isfile(resultatfile):
            i+=1
            resultatfile = str('words-'+nomfichier.split(".",1)[0]+str(i)+'.csv')
    print('Création de '+resultatfile+' contenant les résultats')
    with open(resultatfile, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Mot","Nb occurences"])
        for element in dico:
            writer.writerow([element,dico[element]])
init(autoreset=True)
main()
input("\nAppuyer sur ENTREE pour quiter")