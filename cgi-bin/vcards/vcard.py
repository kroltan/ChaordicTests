# -*- coding: latin-1 -*-
__author__ = 'kroltan1@gmail.com (Leonardo Giovanni Scur)'


try:
    from xml.etree import ElementTree
except ImportError:
    from elementtree import ElementTree
import gdata.spreadsheet.service
import gdata.service
import atom.service
import gdata.spreadsheet
import atom
import sys
import string
import time
import os
import os.path

def zipfolder(path):
    print "aa"

class VCard:
    info = {}
    def __init__(self, version, tags):
        for i, entry in enumerate(tags):
            if i%2==0 and i+1 < len(tags):
                self.info[entry] = tags[i+1] #pega cada par na array e coloca no dict
        if "FN" not in tags:
            self.info["FN"] = "Anonymous"
        self.info["VERSION"] = version #garante a versão especificada no construtor
        self.info["REV"] = int(round(time.time()*1000)) #tempo atual em milisegundos
    
    def save(self, path):
        valid = True
        for key, val in self.info.iteritems():
            if key == None or val == None:
                valid = False
                print "Tag %s:%s is invalid, skipping save!"%(key, val)
        if valid:
            if not os.path.exists(os.path.dirname(path)): os.makedirs(os.path.dirname(path))
            contents = "BEGIN:VCARD"
            i = 0
            for key, val in self.info.iteritems():
                contents = "%s\n%s:%s" % (contents, key, val)
                i = i + 1
            contents = "%s\nEND:VCARD" % contents
            f = open(path, "w")
            f.write(contents)
            f.flush()
            f.close()
        return valid

class SpreadsheetAcessor:
    service = None
    def __init__(self, email, passw, spreadsheet, worksheet):
        # Inicializa o serviço que acessa as spreadsheets
        self.service = gdata.spreadsheet.service.SpreadsheetsService() #instancia o serviço que realiza as operações
        self.service.email = email #email da conta google
        self.service.password = passw #senha da conta google
        self.service.source = "vCard Generator" #identificação ("última modificação", etc)
        try:
            self.service.ProgrammaticLogin() #autentica ao servidor do google
        except:
            print "Authentication error."
        # Procura a spreadsheet (planilha) especificada através do nome
        try:
            feed = self.service.GetSpreadsheetsFeed() #obtém o feed atom representando a planilha
            self.spreadsheet = self._findKey(feed, spreadsheet) #encontra key da planilha
            # Procura a worksheet (página) especificada dentro da spreadsheet encontrada
            feed = self.service.GetWorksheetsFeed(self.spreadsheet) #obtém o feed atom da página
            self.worksheet = self._findKey(feed, worksheet) #encontra key da página
        except:
            print "Error when locating data"

    # Função básica para filtrar um feed e achar a key de um valor correspondente
    def _findKey(self, feed, name):
        # Caso não ache o valor correto, irá retornar o primeiro
        r = feed.entry[0].id.text.split("/") #separa o endereço do feed...
        r = r[len(r)-1] #... e pega a última parte
        for i in range(len(feed.entry)-1):
            if feed.entry[i].title.text == name:
                r = feed.entry[i].id.text.split("/") #separa o endereço do feed...
                r = r[len(r)-1] #... e pega a última parte
                break
        return r
    # Encontra o valor da célula no endereço especificado
    def getCell(self, adderess):
        self.cells = self.service.GetCellsFeed(self.spreadsheet, self.worksheet) #obtém feed de células
        for i, entry in enumerate(self.cells.entry): 
            if isinstance(self.cells, gdata.spreadsheet.SpreadsheetsCellsFeed) and entry.title.text == adderess:
                return entry.content.text

    # modifica o conteudo de uma celula no endereço especificado
    def setCell(self, row, col, val):
        #Notas: X e Y começam em 0
        self.service.UpdateCell(row, col, val, self.spreadsheet, self.worksheet)
    # Transforma coordenadas X,Y em endereços de célula
    #Notas: X e Y começam em 0
    def cellAdderess(self, x, y):
        xv = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        return "%s%s" % (xv[x], y+1)


def generate(args):
    FIELDS = 3
    if len(args) < 7:
        print "Incorrect number of arguments! \nMake sure you put them in this order: \nvcard.py user pass spreadsheet worksheet amount destination"
        return
    googleUser = args[1]
    googlePass = args[2]
    googleSpread = args[3]
    googleWork = args[4]
    vcardAmount = string.atoi(args[5])
    vcardPath = args[6]
    acessor = SpreadsheetAcessor(googleUser, googlePass, googleSpread, googleWork) #cria um acessor para a planilha especificada
    vcards = []
    for y in range(vcardAmount):
        i = [None for i in range(FIELDS*2)] #temp para guardar informações ate a criação do vCard
        for x in range(FIELDS):
            i[x*2] = acessor.getCell(acessor.cellAdderess(x, 0)) #pega o header de cada coluna
            i[x*2+1] = acessor.getCell(acessor.cellAdderess(x, y+1)) #pega a informação referente à linha
        vcards.append(VCard(4.0, i))
        vcards[y].save("%s%d.vcf"%(vcardPath,y))
    

# Só executa se for chamado diretamente.
if __name__ == "__main__":
    generate(sys.argv)
