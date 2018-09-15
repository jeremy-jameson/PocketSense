#ofx.py
# http://sites.google.com/site/pocketsense/

# Original version: by Steve Dunham

# Revisions
# ---------
# 2009: TFB @ "TheFinanceBuff.com"

# Feb-2010*rlc (pocketsense) 
#       - Modified use of the code (call methods, etc.). Use getdata.py to call this routine for specific accounts.
#       - Added scrubber module to clean up known issues with statements. (currently only Discover)
#       - Modified sites structure to include minimum download period
#       - Moved site, stock, fund and user-defined parameters to sites.dat and implemented a parser
#       - Perform a bit of validation on output files before sending to Money
#       - Substantial script edits, so that users shouldn't have a need to debug/edit code.

# 07-May-2010*rlc
#   - Try not to bomb if the server connection fails or times out and set return STATUS accordinging

# 10-Sep-2010*rlc
#   - Added timeout to https call

# 12-Oct-2010*rlc
#   - Fixed bug w/ '&' character in SiteName entries

# 30-Nov-2010*rlc
#   - Catch missing <SECLIST> in statements when a <INVPOSLIST> section exists.  This appears to be a required
#     pairing, but sometimes Vanguard to omits the SECLIST when there are no transactions for the period.
#     Money bombs royally when it happens...

# 01-May-2011*rlc
#   - Replaced check for (<INVPOSLIST> & <SECLIST>) pair with a check for (<INVPOS> & <SECLIST>)

# 18Aug2012*rlc
#   - Added support for OFX version 103 and ClientUID parameter.  The version is defined in sites.dat for a 
#     specific site entry, and the ClientUID is auto-generated for the client and saved in sites.dat

# 20Aug2012*rlc
#   - Changed method used for interval selection (default passed by getdata)

# 15Mar2013*rlc
#   - Added sanity check for <SEVERITY>ERROR code in server reply

# 11Feb2015*rlc
#   - Added support for mapping bank accounts to multiple Money accounts
#     Account "versions" are defined by adding a ":xx" suffix in Setup.py.  
#     The appended "version" is stripped from the account# before passing 
#     to the bank, but is used when sending the results to Money.  

# 11Mar2017*rlc
#   - Changed httplib requests to manually populate headers, due to Discover quackery
#   - Add support for OFX 2.x (xml) exchange while responding to Discover issue
#   - See discussions circa Mar-2017, and contributions from Andrew Dingwall
#   - Add support for site and user specific ClientUID

# 17May2017*rlc
#   - Add V1 POST method as a fallback, for servers that don't like the newer header

# 13Jul2017*rlc
#   - Add support for session cookies in response to change @ Vanguard 

# 24Aug2018*rlc
#   - Remove CLTCOOKIE from request.  Not supported by Money 2005+ or Quicken

import time, os, sys, httplib, urllib2, glob, random, re
import getpass, scrubber, site_cfg, uuid
from control2 import *
from rlib1 import *

if Debug:
    import traceback

#define some function pointers
join = str.join
argv = sys.argv

#define some globals
userdat = site_cfg.site_cfg()
                                               
class OFXClient:
    #Encapsulate an ofx client, site is a dict containg site configuration
    def __init__(self, site, user, password):
        self.password = password
        self.status = True
        self.user = user
        self.site = site
        self.ofxver = FieldVal(site,"ofxver")
        self.url = FieldVal(self.site,"url")
        
        #example: url='https://test.ofx.com/my/script'
        prefix, path = urllib2.splittype(self.url)
        #path='//test.ofx.com/my/script';  Host= 'test.ofx.com' ; Selector= '/my/script'
        self.urlHost, self.urlSelector = urllib2.splithost(path)
        if Debug: 
            print 'urlHost    :', self.urlHost
            print 'urlSelector:', self.urlSelector, '\n'
        self.cookie = 3

    def _cookie(self):
        self.cookie += 1
        return str(self.cookie)

    #Generate signon message
    def _signOn(self):
        site = self.site
        ver  = self.ofxver
        
        clientuid=''
        if int(ver) > 102: 
            #include clientuid if version=103+, otherwise the server may reject the request
            clientuid = OfxField("CLIENTUID", clientUID(self.url, self.user), ver)
        
        fidata = [OfxField("ORG",FieldVal(site,"fiorg"), ver)]
        fidata += [OfxField("FID",FieldVal(site,"fid"), ver)]
        rtn = OfxTag("SIGNONMSGSRQV1",
                OfxTag("SONRQ",
                OfxField("DTCLIENT",OfxDate(), ver),
                OfxField("USERID",self.user, ver),
                OfxField("USERPASS",self.password, ver),
                OfxField("LANGUAGE","ENG", ver),
                OfxTag("FI", *fidata),
                OfxField("APPID",FieldVal(site,"APPID"), ver),
                OfxField("APPVER", FieldVal(site,"APPVER"), ver),
                clientuid
                ))
        return rtn

    def _acctreq(self, dtstart):
        req = OfxTag("ACCTINFORQ",OfxField("DTACCTUP",dtstart))
        return self._message("SIGNUP","ACCTINFO",req)

    def _bareq(self, bankid, acctid, dtstart, acct_type):
        site=self.site
        ver=self.ofxver
        req = OfxTag("STMTRQ",
                OfxTag("BANKACCTFROM",
                OfxField("BANKID",bankid, ver),
                OfxField("ACCTID",acctid, ver),
                OfxField("ACCTTYPE",acct_type, ver)),
                OfxTag("INCTRAN",
                OfxField("DTSTART",dtstart, ver),
                OfxField("INCLUDE","Y", ver))
                )
        return self._message("BANK","STMT",req)
    
    def _ccreq(self, acctid, dtstart):
        site=self.site
        ver  = self.ofxver
        req = OfxTag("CCSTMTRQ",
              OfxTag("CCACCTFROM",OfxField("ACCTID",acctid, ver)),
              OfxTag("INCTRAN",
              OfxField("DTSTART",dtstart, ver),
              OfxField("INCLUDE","Y", ver)))
        return self._message("CREDITCARD","CCSTMT",req)

    def _invstreq(self, brokerid, acctid, dtstart):
        dtnow = time.strftime("%Y%m%d%H%M%S",time.localtime())
        ver  = self.ofxver
        req = OfxTag("INVSTMTRQ",
                OfxTag("INVACCTFROM",
                    OfxField("BROKERID", brokerid, ver),
                    OfxField("ACCTID",acctid, ver)),
                OfxTag("INCTRAN",
                    OfxField("DTSTART",dtstart, ver),
                    OfxField("INCLUDE","Y", ver)),
                OfxField("INCOO","Y", ver),
                OfxTag("INCPOS",
                    OfxField("DTASOF", dtnow, ver),
                    OfxField("INCLUDE","Y", ver)),
                OfxField("INCBAL","Y", ver))
        return self._message("INVSTMT","INVSTMT",req)

    def _message(self,msgType,trnType,request):
        site = self.site
        ver  = self.ofxver
        return OfxTag(msgType+"MSGSRQV1",
               OfxTag(trnType+"TRNRQ",
               OfxField("TRNUID",ofxUUID(), ver),
               request))
    
    def _header(self):
        site = self.site
        if self.ofxver[0]=='2':
            rtn = """<?xml version="1.0" encoding="utf-8" ?>
                     <?OFX OFXHEADER="200" VERSION="%ofxver%" SECURITY="NONE" OLDFILEUID="NONE" NEWFILEUID="NONE"?>"""
            rtn = rtn.replace('%ofxver%', self.ofxver)

        else:
            rtn = join("\r\n",[ "OFXHEADER:100",
                           "DATA:OFXSGML",
                           "VERSION:" + self.ofxver,
                           "SECURITY:NONE",
                           "ENCODING:USASCII",
                           "CHARSET:1252",
                           "COMPRESSION:NONE",
                           "OLDFILEUID:NONE",
                           "NEWFILEUID:NONE",
                           ""])
                           
        return rtn

    def baQuery(self, bankid, acctid, dtstart, acct_type):
        #Bank account statement request
        return join("\r\n",
                    [self._header(),
                     OfxTag("OFX",
                          self._signOn(),
                          self._bareq(bankid, acctid, dtstart, acct_type)
                          )
                    ]
                )
                        
    def ccQuery(self, acctid, dtstart):
        #CC Statement request
        return join("\r\n",[self._header(),
                    OfxTag("OFX",
                    self._signOn(),
                    self._ccreq(acctid, dtstart))])

    def acctQuery(self,dtstart='19700101000000'):
        return join("\r\n",[self._header(),
                    OfxTag("OFX",
                    self._signOn(),
                    self._acctreq(dtstart))])

    def invstQuery(self, brokerid, acctid, dtstart):
        return join("\r\n",[self._header(),
                    OfxTag("OFX",
                    self._signOn(),
                    self._invstreq(brokerid, acctid, dtstart))])

    def doQuery(self,query,name):
        # urllib doesn't honor user Content-type, use urllib2

        response=False
        try:
            errmsg= "** An ERROR occurred attempting HTTPS connection to"
            h = httplib.HTTPSConnection(self.urlHost, timeout=5)
            if Debug: h.set_debuglevel(1)
            
            #proxy config for fiddler tests
            #h = httplib.HTTPSConnection('localhost:8888', timeout=5)
            #h.set_tunnel(self.urlHost)
            
            errmsg= "** An ERROR occurred sending POST request to"
            
            #try without a user-agent or cookie, and retry if the first one fails
            #if both fail, revert to V1 request method
            
            response = None
            
            for i in [0,1,2]:
                if i in [0,1]:
                    #V2 request supports latest Discover
                    h.putrequest('POST', self.urlSelector, skip_host=1, skip_accept_encoding=1)
                    
                    h.putheader('Content-Type', 'application/x-ofx')
                    h.putheader('Host', self.urlHost)
                    h.putheader('Content-Length', str(len(query)))
                    h.putheader('Connection', 'Keep-Alive')
                    
                    #optional parameters are appended only when failure on first pass
                    #   + cookies are added if provided by server on first pass
                    #   + user-agent is always added on second pass
                    if i==1:
                        h.putheader('User-Agent', 'PocketSense')
                        #this is our second pass, so add session cookies if found in first response
                        cookie = response.getheader('set-cookie')   #server cookie(s) provided in last response
                        if cookie <> None: 
                            if Debug: print '<Response Cookies>', cookie
                            h.putheader('cookie', cookie)
                    h.endheaders(query)

                else:
                   #i=2: try V1 request (deprecated).  Shouldn't get here... keeping "just in case"
                   h.request('POST', self.urlSelector, query, 
                             {"Content-type": "application/x-ofx",
                              "Accept": "application/x-ofx"})
                    
                errmsg= "** An ERROR occurred retrieving POST response from"
                #allow up to 30 secs for the server response (if it takes longer, something's wrong)
                h.sock.settimeout(30) 
                response = h.getresponse()
                respDat  = response.read()
            
                #if this is a OFX 2.x response, replace the header w/ OFX 1.x
                if self.ofxver[0] == '2':
                    respDat = re.sub(r'<\?.*\?>', '', respDat)      #remove xml header lines like <? content...content ?>
                    respDat = OfxSGMLHeader() + respDat.lstrip()
            
                #did we get a valid response?  if not, try again w/ different request header
                if validOFX(respDat)=='': break
            
            f = file(name,"w")
            f.write(respDat)
            f.close()
            
        except Exception as e:
            self.status = False
            print errmsg, self.urlHost
            print "   Exception type  :", type(e)
            print "   Exception val   :", e

            if response:
                print "   HTTPS ResponseCode  :", response.status
                print "   HTTPS ResponseReason:", response.reason

        if h: h.close()   
        
#------------------------------------------------------------------------------

def getOFX(account, interval):

    sitename   = account[0]
    _acct_num  = account[1]             #account value defined in sites.dat
    acct_type  = account[2]
    user       = account[3]
    password   = account[4]
    acct_num = _acct_num.split(':')[0]  #bank account# (stripped of :xxx version)
    
    #get site and other user-defined data
    site = userdat.sites[sitename]
    
    #set the interval (days)
    minInterval = FieldVal(site,'mininterval')    #minimum interval (days) defined for this site (optional)
    if minInterval:
         interval = max(minInterval, interval)    #use the longer of the two
    
    #set the start date/time
    dtstart = time.strftime("%Y%m%d",time.localtime(time.time()-interval*86400))
    dtnow = time.strftime("%Y%m%d%H%M%S",time.localtime())
  
    client = OFXClient(site, user, password)
    print sitename,':',acct_num,": Getting records since: ",dtstart
    
    status = True
    #we'll place ofx data transfers in xfrdir (defined in control2.py).  
    #check to see if we have this directory.  if not, create it
    if not os.path.exists(xfrdir):
        try:
            os.mkdir(xfrdir)
        except:
            print '** Error.  Could not create', xfrdir
            system.exit()
    
    #remove illegal WinFile characters from the file name (in case someone included them in the sitename)
    #Also, the os.system() call doesn't allow the '&' char, so we'll replace it too
    sitename = ''.join(a for a in sitename if a not in ' &\/:*?"<>|()')  #first char is a space
    
    ofxFileSuffix = str(random.randrange(1e5,1e6)) + ".ofx"
    ofxFileName = xfrdir + sitename + dtnow + ofxFileSuffix
    
    try:
        if acct_num == '':
            query = client.acctQuery()
        else:
            caps = FieldVal(site, "CAPS")
            if "CCSTMT" in caps:
                query = client.ccQuery(acct_num, dtstart)
            elif "INVSTMT" in caps:
                #if we have a brokerid, use it.  Otherwise, try the fiorg value.
                orgID = FieldVal(site, 'BROKERID')
                if orgID == '': orgID = FieldVal(site, 'FIORG')
                if orgID == '':
                    msg = '** Error: Site', sitename, 'does not have a (REQUIRED) BrokerID or FIORG value defined.'
                    raise Exception(msg)
                query = client.invstQuery(orgID, acct_num, dtstart)

            elif "BASTMT" in caps:
                bankid = FieldVal(site, "BANKID")
                if bankid == '':
                    msg='** Error: Site', sitename, 'does not have a (REQUIRED) BANKID value defined.'
                    raise Exception(msg)
                query = client.baQuery(bankid, acct_num, dtstart, acct_type)

        SendRequest = True
        if Debug: 
            print query
            print
            ask = raw_input('DEBUG:  Send request to bank server (y/n)?').upper()
            if ask=='N': return True, ''
        
        #do the deed
        client.doQuery(query, ofxFileName)
        if not client.status: return False, ''
        
        #check the ofx file and make sure it looks valid (contains header and <ofx>...</ofx> blocks)
        if glob.glob(ofxFileName) == []:
            status = False  #no ofx file?
        else: 
            f = open(ofxFileName,'r')
            content = f.read().upper()
            f.close

            if acct_num <> _acct_num:
                #replace bank account number w/ value defined in sites.dat
                content = content.replace('<ACCTID>'+acct_num, '<ACCTID>'+ _acct_num)
                f = open(ofxFileName,'w')
                f.write(content)
                f.close()
                
            content = ''.join(a for a in content if a not in '\r\n ')  #strip newlines & spaces
            msg = validOFX(content)  #checks for valid format and error messages
            
            if msg<>'':
                #throw exception and exit
                raise Exception(msg)
                
            #attempted debug of a Vanguard issue... rlc*2010
            if content.find('<INVPOS>') > -1 and content.find('<SECLIST>') < 0:
                #An investment statement must contain a <SECLIST> section when a <INVPOSLIST> section exists
                #Some Vanguard statements have been missing this when there are no transactions, causing Money to crash
                #It may be necessary to match every investment position with a security entry, but we'll try to just
                #verify the existence of these section pairs. rlc*9/2010
                raise Exception("OFX statement is missing required <SECLIST> section.")
                
            #cleanup the file if needed
            scrubber.scrub(ofxFileName, site)
        
    except Exception as inst:
        status = False
        print inst
        if glob.glob(ofxFileName) <> []:
           print '**  Review', ofxFileName, 'for possible clues...'
        if Debug:
            traceback.print_exc()
        
    return status, ofxFileName
