#!python
__author__ = 'kroltan1@gmail.com (Leonardo Giovanni Scur)'


import cgitb
import cgi
import io
import sys
import vcard
import gdata.spreadsheet.text_db

def printHtml(path):
    f = open(path)
    line = f.readline()
    while line:
        sys.stdout.write(line)
        line = f.readline() 
def formHasKeys(form, keys):
    has = True
    for k in keys:
        if not form.getfirst(k, ""):
            has = False
            break
    return has

cgitb.enable()
form = cgi.FieldStorage()
print "Content-type: text/html"
print ""
printHtml("form.html")
if formHasKeys(form, ["gUser", "gPass", "gSpread", "gWork", "FN", "EMAIL", "TEL"]):
    vci = {"fn":form.getfirst("FN", "Anonymous"), 
           "email":form.getfirst("EMAIL", "N/A"), 
           "tel":form.getfirst("TEL", "N/A")}
    try:
        server = gdata.spreadsheet.text_db.DatabaseClient(username=form.getfirst("gUser"), password=form.getfirst("gPass"))
        table = server.GetDatabases(None, form.getfirst("gSpread"))[0].GetTables(None, form.getfirst("gWork"))[0]
        record = table.AddRecord(vci)
        record.Push()
    except Exception as e:
        print e