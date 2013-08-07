# -*- coding: latin-1 -*-
__author__ = 'kroltan1@gmail.com (Leonardo Giovanni Scur)'


try:
    from xml.etree import ElementTree
except ImportError:
    from elementtree import ElementTree
import gdata.spreadsheet.service
import gdata.service
import atom.service
import gdata.spreadsheet.text_db
import atom
import sys
import string
import time
import os
import os.path

class VCard:
    info = {}
    def __init__(self, version, tags):
        for key, val in tags.iteritems():
            self.info[key.upper()] = val #pega cada par na array e coloca no dict
        if "fn" not in tags.keys():
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
                contents = "%s\n%s:%s" % (contents, key.upper(), val)
                i = i + 1
            contents = "%s\nEND:VCARD" % contents
            f = open(path, "w")
            f.write(contents)
            f.flush()
            f.close()
        return valid


class SpreadsheetAcessor:
    service = None
    table = None
    def __init__(self, email, passw, spreadsheet, worksheet):
        try:
            # Inicializa o serviço que acessa as spreadsheets
            self.service = gdata.spreadsheet.text_db.DatabaseClient(email, passw)
            # Procura a spreadsheet (planilha) especificada através do nome
            self.table = self.service.GetDatabases(None, spreadsheet)[0].GetTables(None, worksheet)[0]
        except Exception as e:
            print e
    
    def getRecordData(self, id):
        return self.table.GetRecords(id, id+1)[0].content
    


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
    for i in range(vcardAmount):
        info = acessor.getRecordData(i+1)
        vcards.append(VCard(4.0, info))
        vcards[i].save("%s%d.vcf"%(vcardPath,i))
    

# Só executa se for chamado diretamente.
if __name__ == "__main__":
    generate(sys.argv)