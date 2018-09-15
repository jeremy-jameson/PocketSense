# rlib1.py :: General library module
# http://sites.google.com/site/pocketsense/
# Intial version: rlc: Dec-2010
#   - Initially contains QuoteHTMwriter() and supporting functions

# 10-Jan-2011*rlc
#    Moved _header, _field, _tag, _genuid, and _date functios from quotes.py.  Renamed and edited.
# 23Aug2012*rlc
#   - Added combineOFX()
# 19Jan2014*rlc:  
#   -Added support for Google Finance quotes
# 12Mar2017*rlc
#   - Add support for OFX 2.x (xml) field tags (while responding to Discover issue)
#   - Add support for user-specific clientUID pairs (by url+username)
# 09Nov2017*rlc
#   - prefix positive change w/ '+' symbol

import os, glob, site_cfg, time, uuid, re, random
import sys, pyDes, md5, pickle, locale, urllib2
from datetime import datetime
from control2 import *

if Debug:
    import traceback

def clientUID(url, username, delKey=False):
    #get clientUID for urlHost+username.  if not exists, create
    #delete key if delKey=True

    dTable = {}
    found=False
    dfile = 'connect.key'
    uuid = None

    #get urlHost:  example: url='https://test.ofx.com/my/script'
    prefix, path = urllib2.splittype(url)
    #path='//test.ofx.com/my/script';  Host= 'test.ofx.com' ; Selector= '/my/script'
    urlHost, urlSelector = urllib2.splithost(path)
    key = md5.md5(urlHost+username).digest()

    if glob.glob(dfile) <> []:
        #lookup
        f = open(dfile,'rb')
        dTable = pickle.load(f) 
        uuid = dTable.get(key, None)
        f.close()
        
    if uuid==None or (delKey and uuid<>None):
        f = open(dfile,'wb')
        if delKey:
            #remove existing key
            dTable.pop(key, None)
        else:
            #add new key
            uuid=str(ofxUUID())
            dTable[key] = uuid

        pickle.dump(dTable, f)
        f.close()
        
    return uuid
    
def get_int(prompt):
    #get number entry
    prompt = prompt.rstrip() + ' '
    done = False
    while not done:
        istr = raw_input(prompt)
        if istr == '':
            a = 0
            done = True
        else:
            try:
                a=int(istr)
                done = True
            except:
                print 'Please enter a valid integer'
    return a

def FieldVal(dic, fieldname):
    #get field value from a dict list
    #return value for fieldname (returns as type defined in dict)
    val = ''
    fieldname = fieldname.upper()
    if fieldname in dic:
        val = dic[fieldname]
    return val

def decrypt_pw(pwkey):
    #validate password if pwkey isn't null
    if pwkey <> '':
        #file encrypted... need password
        pw = pyDes.getDESpw()   #ask for password
        k = pyDes.des(pw)       #create encryption object using key
        pws = k.decrypt(pwkey,' ')  #decrypt
        if pws <> pw:               #comp to saved password
            print 'Invalid password.  Exiting.'
            sys.exit()
        else:
            #decrypt the encrypted fields
            pwkey = pws
    return pwkey
       
def acctEncrypt(AcctArray, pwkey):
    #encrypt accounts
    d = pyDes.des(pwkey)
    for acct in AcctArray:
       acct[1] = d.encrypt(acct[1],' ')
       acct[3] = d.encrypt(acct[3],' ')
       acct[4] = d.encrypt(acct[4],' ')
    return AcctArray
    
def acctDecrypt(AcctArray, pwkey):
    #decrypt accounts
    d = pyDes.des(pwkey)
    for acct in AcctArray:
       acct[1] = d.decrypt(acct[1],' ')
       acct[3] = d.decrypt(acct[3],' ')
       acct[4] = d.decrypt(acct[4],' ')
    return AcctArray
    
def get_cfg():
    #read in user configuration
    
    c_AcctArray = []        #AcctArray = [['SiteName', 'Account#', 'AcctType', 'UserName', 'PassWord'], ...]
    c_pwkey=''              #default = no encryption
    c_getquotes = False     #default = no quotes
    if glob.glob(cfgFile) <> []:
        cfg = open(cfgFile,'rb')
        try:
            c_pwkey = pickle.load(cfg)            #encrypted pw key
            c_getquotes = pickle.load(cfg)        #get stock/fund quotes?
            c_AcctArray = pickle.load(cfg)        #
        except:
            pass    #nothing to do... must not be any data in the file
        cfg.close()
    return c_pwkey, c_getquotes, c_AcctArray

def QuoteHTMwriter(qList):
    # Write quotes.htm containing quote data contained in quote list (qList)
    # Supports Yahoo! finance links
    # See quotes.py for qList structure
    global userdat
    
    userdat = site_cfg.site_cfg()
    
    # CREATE FILE
    filename = xfrdir + "quotes.htm"
    fullpath = '"' + os.path.realpath(filename) + '"'   #encapsulate spaces
    
    f = open(filename,"w")
    print "Writing", filename
    
    # Write HEADER
    _QHTMheader(f)
    
    # Write BODY
    shade = False
    for quote in qList:
        _QHTMrow(f, quote, shade)
        shade = not shade
    
    # Write FOOTER
    _QHTMfooter(f)
    
    f.close()
    
    return fullpath
    
def _QHTMheader(f):
    #header for quotes.htm
    
    header = """
        <! 
        Generated using PocketSense Python scripts for Microsoft Money
        http://sites.google.com/site/pocketsense
        Ver 1.0: Dec-2010: RLC
        /!>

        <style type="text/css">

        body {font-family: Verdana;
              font-size: 13;
              background-color: white}

        table {
            text-align: center;
            font-family: Verdana;
            font-weight: normal;
            font-size: 16px;
            border: 2px solid gray;
            border-collapse: collapse;
            empty-cells: show
        }

        /* header */
        th {
            border: 2px solid gray;
            font-weight: bold;
            padding: 5;
        }

        /* cell */
        td {
            font-family: monospace;
            border: 1px;
            border-style: dotted solid;
            padding-left: 6;
            padding-right: 6;
        }

        td.s0 {background-color: FFFFFF}  /* background */
        td.s1 {background-color: F5F5F5}  /* shaded cell */
        td.s2 {background-color: 98FB98}  /* green bkg */
        td.s3 {background-color: EE5C42}  /* red bkg */
        </style>
        
        <body topmargin=30 leftmargin=50>
        <h1 align=left><font color=blue>My Quotes</font>
        <a href='""" + userdat.YahooURL + """'><img width=140 border="0" src="http://l.yimg.com/a/i/brand/purplelogo/uh/us/fin.gif"></a>
        <a href='""" + userdat.GoogleURL + """'><img width=70 border="0" 
        src="http://www.google.com/images/logos/google_logo_41.png"></a>
        </h1>
        <p>

        <table>
        <tr><th>Source</th><th>Symbol</th><th>Name</th><th>Price</th><th>Time</th>
        <th >%Change<sup>(1)</sup></th></tr>
        """
        
    f.write(header)
    return

def _QHTMrow(f, quote, shade):
    #write table row for quote to file f
    #see quote.py for quote data structure
    #shade = shade row?
    
    #add + for non-negatives.  pchange is a formatted % string
    pchangeV = float(quote.pchange.strip('%'))
    pchange = ('+' if pchangeV>0 else '') + quote.pchange
    
    if shade: 
        td1   = '<td class="s1">'
        td1L = '<td align="left" class="s1">'
        td1R = '<td align="right" class="s1">'
    else:
        td1   = '<td class="s0">'
        td1L = '<td align="left" class="s0">'
        td1R = '<td align="right" class="s0">'
    
    if '-' in quote.pchange:
        td2 = '<td class="s3">'
    elif 'N/A' in quote.pchange or '?' in quote.pchange or pchangeV==0:
        td2 = td1       # no change given, or zero change... no shade
    else:
        td2 = '<td class="s2">'
    
    tspace = '&nbsp;' * (8-len(quote.time)) #right just time
    lspace = '&nbsp;' * (10-len(quote.date)) #leave space for double-digit month
    
    row = td1 + quote.source + '</td>' + \
          td1 + '<a href=' + quote.quoteURL +'>'+quote.symbol+'</a></td>' + \
          td1L+ quote.name+'</td>' + \
          td1R+ quote.price + '</td>' + \
          td1 + lspace + quote.date + tspace + quote.time +'</td>' + \
          td2 + pchange + '</td></tr>'
    
    f.write(row)
    return
    
def _QHTMfooter(f):
    #footer for quotes.htm
    qTime = datetime.today().strftime("%m-%d-%Y at %H:%M:%S")
    
    sitepath = os.path.realpath('sites.dat')
    sitepath = '<a href="file:///' + sitepath + '"><b>sites.dat</b></a>'
    
    footer = """
        </table>

        1.  %Change = Percent change in price since the last close.
        <p><br>
        This page was generated by the <a href="http://sites.google.com/site/pocketsense/home/msmoneyfixp1">PocketSense Python scripts</a> for  Money (on """ + qTime + """)</br>
        Quotes provided by <a href='""" + userdat.YahooURL + """'>Yahoo! Finance</a>
        and <a href='""" + userdat.GoogleURL + """'>Google Finance</a>.<br>
        Stock and fund symbols are defined in your """ + sitepath + """  file.
        </p>
        </body>
        """
    f.write(footer)
    return

def OfxSGMLHeader():
    #Standard OFX SGML Header
    return """OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:TYPE1
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

"""
        
def OfxField(tag,value, ofxver='102'):
    field = ''
    #skip empty values
    if tag <> '' and value <> '':
        field = '<'+tag+'>'+value
        #terminate as xml if ofx 2.x
        if ofxver[0]=='2': field = field + '</'+tag+'>'
    return field

def OfxTag(tag,*contents):
    tag1 = '<' + tag + '>'
    tag2 = '</' + tag + '>'
    return '\r\n'.join([tag1]+list(contents)+[tag2])

def OfxDate():
    return time.strftime("%Y%m%d%H%M%S",time.localtime())

def ofxUUID():
    return str(uuid.uuid4())

def validOFX(content):
    #does content appear to be a valid ofx statement?  returns message indicating reason (null if valid)
    msg=''
    content = content.upper().rstrip()
    
    if content == '': msg = 'No (null) statement received'
    
    elif content.find('OFXHEADER:') < 0 and content.find('<OFX>') < 0 and content.find('</OFX>') < 0:
        msg = 'Invalid OFX statement received'
        
    elif content.find('<SEVERITY>ERROR') > 0:
        msg = 'OFX message contains ERROR condition'
    
    #note:  spaces get stripped before this function is called
    elif content.find('ACCESSDENIED') > 0:
        msg = 'Access denied'
    
    return msg
    
def int2(str):
    #convert str to int, without throwing exception.  If str is not a "number", returns zero.
    try:
        f = int(str)
    except:
        f = 0
    return f
    
def float2(str):
    #convert str to float, without throwing exception.  If str is not a "number", returns zero.
    try:
        f = float(str)
    except:
        f = 0.0
    return f
    
def runFile(filename):
    #encapsulate call to os.system in quotes
    os.system('"'+filename+'"')
    return
    
def copy_txt_file(infile, outfile):
    #copy text file
    #for some reason, the shutil.copy() routines don't handle windows cr/lf correctly
    inp = open(infile,'r')
    out = open(outfile,'w')
    out.write(inp.read())
    return
   
def combineOfx(ofxList):
    #combine ofx statements into a single file in a manner that Money seems to accept
    
    dtnow = time.strftime("%Y%m%d%H%M%S",time.localtime())
    signon =  "\r".join([
              "<SIGNONMSGSRSV1><SONRS>",
              "<STATUS><CODE>0<SEVERITY>INFO<MESSAGE>Successful Sign On</STATUS>",
              "<DTSERVER>" + dtnow,
              "<LANGUAGE>ENG<DTPROFUP>20010101010000",
              "<FI><ORG>PocketSense</FI></SONRS></SIGNONMSGSRSV1>"])
    
    #these regexes capture everything between the tags, but not the tags
    bRe = re.compile('(?:<BANKMSGSRSV1>)(.*?)(?:</BANKMSGSRSV1>)', re.IGNORECASE)
    cRe = re.compile('(?:<CREDITCARDMSGSRSV1>)(.*?)(?:</CREDITCARDMSGSRSV1>)', re.IGNORECASE)
    iRe = re.compile('(?:<INVSTMTMSGSRSV1>)(.*?)(?:</INVSTMTMSGSRSV1>)', re.IGNORECASE)
    sRe = re.compile('(?:<SECLIST>)(.*?)(?:</SECLIST>)', re.IGNORECASE)
    
    bantrn=''
    crdtrn=''
    invtrn=''
    sectrn=''

    for file in ofxList:
        if glob.glob(file[2]):
            f=open(file[2])
            ofx = f.read()
            f.close()
            
            ofx = ofx.replace(chr(13),'')   #remove CRs
            ofx = ofx.replace(chr(10),'')   #remove LFs
            
            #create a string for each section found in the file
            #re.findall() returns a list of all matching sections
            b = '\r'.join(bRe.findall(ofx))
            c = '\r'.join(cRe.findall(ofx))
            i = '\r'.join(iRe.findall(ofx))
            s = '\r'.join(sRe.findall(ofx))
            
            #add statements to each section
            if b: bantrn = bantrn + '\r' + b
            if c: crdtrn = crdtrn + '\r' + c
            if i: invtrn = invtrn + '\r' + i
            if s: sectrn = sectrn + '\r' + s

    if bantrn: bantrn = OfxTag('BANKMSGSRSV1', bantrn)
    if crdtrn: crdtrn = OfxTag('CREDITCARDMSGSRSV1', crdtrn)
    if invtrn: invtrn = OfxTag('INVSTMTMSGSRSV1', invtrn)
    if sectrn: sectrn = OfxTag('SECLISTMSGSRSV1', OfxTag('SECLIST', sectrn))
    
    combOfx = '\r'.join(['<OFX>', signon, bantrn, crdtrn, invtrn, sectrn, '</OFX>'])
    
    #remove blank lines (not required... just to clean it up)
    combOfx2=''
    for line in combOfx.splitlines():
        if line: combOfx2 = combOfx2 + line + '\r'
    
    combOfx = OfxSGMLHeader() + combOfx2
    
    #there should never be two combined*.ofx files here, but we'll use a unique name just in case
    cfile = xfrdir + 'combined' + str(random.randrange(1e5,1e6)) + '.ofx'
    f=open(cfile,'w')
    f.write(combOfx)
    f.close()
    print "Combined OFX created: " + cfile
    return cfile