# scrubber.py
# http://sites.google.com/site/pocketsense/
# fix known issues w/ OFX downloads
# rlc: 2010

# 05-Aug-2010*rlc
# - Added _scrubTime() function to fix NULL time stamps so that transactions record on the correct date
#   regardless of time zone.

# 28-Jan-2011*rlc
#   - Added _scrubDTSTART() to fix missing <DTEND> fields when a <DTSTART> exists.
#   - Recoded scrub routines to use regex substitutions

# 28-Aug-2012*rlc
#   - Added quietScrub option to sites.dat (suppresses scrub messages)

# 02-Feb-2013*rlc
#   - Bug fix in _scrubDTSTART
#   - Added scrub routine to verify Investment buy and sell transactions

# 17-Feb-2013*rlc
#   - Added scrub routine for CORRECTACTION and CORRECTFITID tags (not supported by Money)

# 06-Jul-2013*rlc
#   - Bug fix in _scrubDTSTART()

# 20-Feb-2014*rlc
#   - Bug fix in _scrubINVsign() for SELL transactions

#14-Aug-2016*rlc
#   - Update to handle new Discover Bank FITID format.

#11-Mar-2017*rlc
#   - Added REINVEST transactions to _scrubINVsign()
#   - Added _scrubRemoveZeroTrans().  Removes $0.00 transactions when enabled in sites.dat

#08-Apr-2017*cgn / rlc
#   - Revert _scrubINVsign() to previous version
#   - Add _scrubREINVESTsign() to handle REINVEST transactions separately 
#     Note the differnt field order vs what's used for BUY/SELL transactions in _scrubINVsign()

#27-Aug-2017*rlc
#   - Bug patch to fix timedelta call

#24-Oct-2017*ad / rlc
#   - Update to address recent change by Discover Bank re transaction ids
#   - Insert check# field for discover bank transactions such that Money recognizes it

#01-Jan-2018*rlc
#   - Revert Discover Bank fitid substitution to the same as it was before the 24-Oct update

#14-Apr-2018*rlc
#   - Replace ampsersand "&" symbol w/ &amp; code when not part of valid escape code. see _scrubGeneral()

#27-Jul-2018*dbc
#   - Add TRowePrice scrub function to fix paid-out dividends/cap gains that are marked as reinvested

import os, sys, re
import site_cfg
from datetime import datetime, timedelta
from control2 import *
from rlib1 import *

userdat = site_cfg.site_cfg()
stat = False    #global used between re lambda subs to track status

def scrubPrint(line):
    if not userdat.quietScrub:
        print "  +" + line
    
def scrub(filename, site):
    #filename = string
    #site = DICT structure containing full site info from sites.dat
 
    siteURL = FieldVal(site, 'url').upper()
    dtHrs = FieldVal(site, 'timeOffset')
    accType = FieldVal(site, 'CAPS')[1]
    f = open(filename,'r')
    ofx = f.read()  #as-found ofx message
    #print ofx
    
    #NOTE:  Discover Card and Bank use the same server @ discovercard.com
    if 'DISCOVERCARD' in siteURL: ofx= _scrubDiscover(ofx, accType)
    
    if 'TROWEPRICE.COM' in siteURL: ofx = _scrubTRowePrice(ofx)
    
    ofx= _scrubTime(ofx)     #fix 000000 and NULL datetime stamps 

    if dtHrs <> 0: ofx = _scrubShiftTime(ofx, dtHrs)   #note: always call *after* _scrubTime()
    
    ofx= _scrubDTSTART(ofx)  #fix missing <DTEND> fields
      
    #fix malformed investment buy/sell/reinvest signs (neg vs pos), if they exist
    if "<INVSTMTTRNRS>" in ofx.upper(): 
        ofx= _scrubINVsign(ofx)  
        ofx= _scrubREINVESTsign(ofx)  
	
    #remove $0.00 transactions
    if userdat.skipZeroTransactions: ofx = _scrubRemoveZeroTrans(ofx)
    
    #perform general ofx cleanup
    ofx = _scrubGeneral(ofx)
  
    #close the input file
    f.close()
    
    #write the new version to the same file
    f = open(filename, 'w')
    f.write(ofx)
    f.close    

#-----------------------------------------------------------------------------
# OFX.DISCOVERCARD.COM
#   1.  Discover OFX files will contain transaction identifiers w/ the following format:
#           FITIDYYYYMMDDamt#####, where
#                FITID  = string literal
#                YYYY   = year (numeric)
#                MM     = month (numeric)
#                DD     = day (numeric)
#                amt    = dollar amount of the transaction, including a hypen for negative entries (e.g., -24.95)
#                #####  = 5 digit serial number

#   2.  The 5-digit serial number can change each time you connect to the server, 
#          meaning that the same transaction can download with different FITID numbers.  
#       That's not good, since Money requires a unique FITID value for each valid transaction.  
#       Varying serial numbers result in duplicate transactions!

#   3.  We'll replace the 5-digit serial number with one of our own.  
#       The default will be 0 for every transaction,
#          and we'll increment by one for each subsequent transaction that that matches
#          a previous transaction in the file.

# 8/14/2016: Discover BANK now uses an FITID format of SDF######, where ###### is unique for the day.
#            The length seems to vary, but the largest observed is 6 digits  
#            Unfortunately, the digits can be assigned to multiple transactions on the same day, so it isn't 
#               guaranteed to be unique.  
#            Modified routine to uniquely handle BASTMT vs CCSTMT statements.

# NOTE:  There was brief period in late 2017 where Discover Bank changed their fitid format, but soon
#        reverted to the same as described above.

_sD_knownvals = []  #global to keep track of Discover FITID values between regex.sub() calls

def _scrubDiscover(ofx, accType):

    if accType=='CCSTMT': 
        scrubPrint("Scrubber: Processing Discover Card statement.")
    else:
        scrubPrint("Scrubber: Processing Discover Bank statement.")

    ofx_final = ''      #new ofx message
    _sD_knownvals = []  #reset our global set of known vals (just in case)

    # dev: insert a line break after each transaction for readability.
    # also helps block multi-transaction matching in below regexes via ^\s option
    p = re.compile(r'(<STMTTRN>)',re.IGNORECASE)
    ofx = p.sub(r'\n<STMTTRN>', ofx)
    
    #regex p captures everything from <FITID> up to the next <tag>, but excludes the next "<".
    #p produces 2 results:  r.group(1) = <FITID> field, r.group(2)=value
    #the ^<\s prevents matching on the next < or newline
    p = re.compile(r'(<FITID>)([^<\s]+)',re.IGNORECASE)

    #call substitution (inline lamda, takes regex result = r as tuple)
    ofx_final = p.sub(lambda r: _scrubDiscover_r1(r, accType), ofx)
   
    if accType=='BASTMT':
        #regex p captures everything from <TRNTYPE>DEBIT up to the next "<" aftert the <NAME>Check tag and field.
        # Discover Bank codes checks as
        # <STMTTRN><TRNTYPE>DEBIT<...><NAME>Check ###########</STMTTRN>
        # p produces 4 results:
        #   r.group(1) = <TRNTYPE>DEBIT,
        #   r.group(2) = stuff up to next "<NAME>Check "
        #   r.group(3) = "<NAME>Check ", including the trailing spaces (at least 1)
        #   r.group(4) is the check number (1 or more digits)
        # Rearranged, the result should produce a entry that will import the check number in Money
        # <STMTTRN><TRNTYPE>CHECK<...><CHECKNUM>############<NAME>Check</STMTTRN>
        ofx = ofx_final
        p = re.compile(r'(<TRNTYPE>DEBIT)([^\s]+)(<NAME>Check[ ]+)([0-9]+)',re.IGNORECASE)
        ofx_final = p.sub(lambda r: _scrubDiscover_r2(r, accType), ofx)

    return ofx_final

def _scrubDiscover_r1(r, accType):
    #regex subsitution function: change fitid value
    global _sD_knownvals

    fieldtag = r.group(1)
    fitid = r.group(2).strip(' ')
    fitid_b = fitid                     #base fitid before annotating
    
    #strip the serial value for credit card transactions
    if accType=='CCSTMT': 
        bx = len(fitid) - 5
        fitid_b = fitid[:bx]
    
    #find a unique serial#, from 0 to 9999
    seq = 0   #default
    while seq < 9999:
        fitid = fitid_b + str(seq)
        exists = (fitid in _sD_knownvals)
        if exists:  #already used it... try another
            seq=seq+1
        else:
            break   #unique value... write it out
        
    _sD_knownvals.append(fitid)         #remember the assigned value between calls
    return fieldtag + fitid             #return the new string for regex.sub()

def _scrubDiscover_r2(r, accType):
    #regex subsitution function: insert checknum field for BANK statements
    trntype = r.group(1)
    rest = r.group(2)
    name = r.group(3).strip(' ')
    checknum = r.group(4)
    return '<TRNTYPE>CHECK' + rest + '<CHECKNUM>' + checknum + name

#--------------------------------    
def _scrubTime(ofx):
    #Replace NULL time stamps with noontime (12:00)

    #regex p captures everything from <DT*> up to the next <tag>, but excludes the next "<".
    #p produces 2 results:  group(1) = <DT*> field, group(2)=dateval
    p = re.compile(r'(<DT.+?>)([^<\s]+)',re.IGNORECASE)
    #call date correct function (inline lamda, takes regex result = r tuple)
    
    global stat
    stat = False
    ofx_final = p.sub(lambda r: _scrubTime_r1(r), ofx)
    if stat: scrubPrint("Scrubber: Null time values updated.")
    
    return ofx_final

def _scrubTime_r1(r):
    # Replace zero and NULL time fields with a "NOON" timestamp (120000)
    # Force "date" to be the same as the date listed, regardless of time zone by setting time to NOON.
    # Applies when no time is given, and when time == MIDNIGHT (000000)
    global stat
    fieldtag = r.group(1)
    DT = r.group(2).strip(' ')      #date+time
    
    # Full date/time format example:  20100730000000.000[-4:EDT]
    if DT[8:] == '' or DT[8:14] == '000000':
        #null time given.  Adjust to 120000 value (noon).
        DT = DT[:8] + '120000'
        stat = True
        
    return fieldtag + DT

#--------------------------------    
def _scrubDTSTART(ofx):
    # <DTSTART> field for an account statement must have a matching <DTEND> field
    # If DTEND is missing, insert <DTEND>="now"
    # The assumption is made that only one statement exists in the OFX file (no multi-statement files!)
    
    ofx_final = ofx
    now = datetime.now()
    nowstr = now.strftime("%Y%m%d%H%M00")
    
    if ofx.find('<DTSTART>') >= 0 and ofx.find('<DTEND>') < 0:
        #we have a dtstart, but no dtend... fix it.
        scrubPrint("Scrubber: Fixing missing <DTEND> field")
        
        #regex p captures everything from <DTSTART> up to the next <tag> or white space into group(1)
        p = re.compile(r'(<DTSTART>[^<\s]+)',re.IGNORECASE)
        if Debug: print "DTSTART: findall()=", p.findall(ofx_final)
        #replace group1 with (group1 + <DTEND> + datetime)
        ofx_final = p.sub(r'\1<DTEND>'+nowstr, ofx_final)
    
    return ofx_final

def _scrubShiftTime(ofx, h):
    #Shift DTASOF time values by (float) h hours
    #Added: 15-Feb-2011, rlc
    
    #regex p captures everything from <DTASOF> up to the next <tag> or white-space.
    #p produces 2 results:  group(1) = <DTASOF> field, group(2)=dateval
    p = re.compile(r'(<DTASOF>)([^<\s]+)',re.IGNORECASE | re.DOTALL)
    
    #call date correct function (inline lamda, takes regex result = r tuple)
    if p.search(ofx): 
        scrubPrint("Scrubber: Shifting DTASOF time values " + str(h) + " hours.")
        ofx_final = p.sub(lambda r: _scrubShiftTime_r1(r,h), ofx)    

    return ofx_final

def _scrubShiftTime_r1(r,h):
    #Shift time value by (float) h hours for regex search result r.
    #Added: 15-Feb-2011, rlc
    
    fieldtag = r.group(1)       #date field tag (e.g., <DTASOF>)
    DT = r.group(2).strip(' ')  #date+time

    if Debug: print "fieldtag=", fieldtag, "| DT=" + DT
    
    # Full date/time format example:  20100730120000.000[-4:EDT]
    #separate into date/time + timezone
    tz = ""
    if '[' in DT:
        p = DT.index('[')
        tz = DT[p:]
        DT = DT[:p]
    
    #strip the decimal fraction, if we have it
    if '.' in DT:
        d  = DT.index('.')
        DT = DT[:d]
        
    if Debug: scrubPrint("New DT=" + DT + "| tz=" + tz)
    
    #shift the time
    tval = datetime.strptime(DT,"%Y%m%d%H%M%S")  #convert str to datetime
    deltaT = timedelta(hours=h)    
    tval += deltaT                                        #add hours
    DT = tval.strftime("%Y%m%d%H%M%S") + tz               #convert new datetime to str
        
    return fieldtag + DT

def _scrubINVsign(ofx):
    #Fix malformed parameters in Investment buy/sell sections, if they exist
    #Issue  first noticed with Fidelity netbenefits 401k accounts:  rlc*2013
    
    #BUY transactions:
    #   UNITS must be positive
    #   TOTAL must be negative
    
    #SELL transactions:
    #   UNITS must be negative
    #   TOTAL must be positive
    
    global stat
    stat = False
    p = re.compile(r'(<INVBUY>|<INVSELL>)(.+?<UNITS>)(.+?)(<.+?<TOTAL>)([^<\r\n]+)', re.IGNORECASE)
    ofx_final=p.sub(lambda r: _scrubINVsign_r1(r), ofx)
    if stat:
        scrubPrint("Scrubber: Invalid investment sign (pos/neg) found.  Corrected.")
    
    return ofx_final
    
def _scrubINVsign_r1(r):
    
    global stat
    type=""
    if "INVBUY"  in r.group(1): type = "INVBUY"
    if "INVSELL" in r.group(1): type = "INVSELL"
    qty = r.group(3)
    total=r.group(5)
	
    qty_v=float2(qty)
    total_v=float2(total)
    
    if (type=="INVBUY" and qty_v<0) or (type=="INVSELL" and qty_v>0):
        stat=True
        qty=str(-1*qty_v)

    if (type=="INVBUY" and total_v>0) or (type=="INVSELL" and total_v<0):
        stat=True
        total=str(-1*total_v)
    
    return r.group(1) + r.group(2) + qty + r.group(4) + total

def _scrubREINVESTsign(ofx):
    #Fix malformed parameters in REINVEST transactions, if they exist
    #Issue  first noticed with Fidelity netbenefits 401k accounts:  cgn*2016
    
    #REINVEST transactions:
    #   UNITS must be positive
    #   TOTAL must be negative

    global stat
    stat=False
    p = re.compile(r'(<REINVEST>)(.+?<TOTAL>)(.+?)(<.+?<UNITS>)([^<\r\n]+)', re.IGNORECASE)
    ofx_final=p.sub(lambda r: _scrubREINVESTsign_r1(r), ofx)
    if stat:
        scrubPrint("  +Scrubber: Invalid reinvestment sign (pos/neg) found.  Corrected.")
    
    return ofx_final
    
def _scrubREINVESTsign_r1(r):
    global stat
    qty = r.group(5)
    total=r.group(3)
	
    qty_v=float2(qty)
    total_v=float2(total)
    
    if (qty_v<0):
        stat=True
        qty=str(-1*qty_v)

    if (total_v>0):
        stat=True
        total=str(-1*total_v)
    
    return r.group(1) + r.group(2) + total + r.group(4) + qty
    
def _scrubGeneral(ofx):    
    # General scrub routine for general updates  
    
    #1. Remove tag/value pairs that Money doesn't support
    #define unsupported tags that we've had trouble with
    uTags = []    
    uTags.append('CORRECTACTION')
    uTags.append('CORRECTFITID')
    
    for tag in uTags:
        p = re.compile(r'<'+tag+'>[^<]+',re.IGNORECASE) 
        if p.search(ofx):
            ofx = p.sub('',ofx)
            scrubPrint("Scrubber: <"+tag+"> tags removed.  Not supported by Money.")
    
    #2. Replace ampersands '&' that aren't part of a valid escape code (i.e., is NOT like &amp;, &#012; etc)
    #   literally:  replace '&' chars with '&amp;' when the next chars are not
    #               a '#' or valid alphanumerics followed by a ;
    p = re.compile(r'&(?!#?\w+;)')
    if p.search(ofx):
        scrubPrint("Scrubber: Replace invalid '&' chars with '&amp;'")
        ofx = p.sub('&amp;',ofx)
    
    return ofx

def _scrubRemoveZeroTrans(ofx):
    #Remove transactions with a $0.00 value

    #regex p captures transaction records
    #p produces 3 results:  group(1) = trans header, group(2)=Amount, group(3)=trans suffix

    global stat
    stat=False
    p = re.compile(r'(<STMTTRN>.*?<TRNAMT>)(.+?)(<.*?</STMTTRN>)', 
                   flags=re.DOTALL | re.IGNORECASE)

    ofx = p.sub(lambda r: _scrubRemoveZeroTrans_r1(r), ofx)    
    if stat: scrubPrint('Zero amount ($0.00) transactions removed.')
    return ofx
    
def _scrubRemoveZeroTrans_r1(r):
    # return null transaction when amount=0
    global stat
    amount = float2(r.group(2))
    if amount==0: stat=True
    return None if amount==0 else r.group(1)+r.group(2)+r.group(3) 
 

#-----------------------------------------------------------------------------
# OFX fiorg T. Rowe Price
#   1.  T. Rowe Price OFX reports paid-out (non-reinvested) dividends and capital gains incorrectly
#           in such a way that MS Money ignores the transaction.
#       These transactions can be identified as reinvestments having 0.0 shares.
#       The transaction will be bounded by <REINVEST> and </REINVEST> and contain <UNITS>0.0.
#       Furthermore, the payment is incorrectly shown as a negative value, the memo field is misleading, and
#           a required field (<SUBACCTFUND>) is missing.
#
#       For dividends (<INCOMETYPE>DIV) make these changes to the transaction:
#           Change <REINVEST> to <INCOME>
#           Change <MEMO>DIVIDEND (REINVEST) to <MEMO>DIVIDEND PAID
#           Change <TOTAL>-#.## to <TOTAL>#.##
#           Following <SUBACCTSEC>CASH add <SUBACCTFUND>CASH
#           Delete <UNITS>0.0
#           Delete <UNITPRICE>#.##
#           Change </REINVEST> to </INCOME>

#       For short term capital gains (<INCOMETYPE>CGSHORT) make these changes to the transaction:
#           Change <REINVEST> to <INCOME>
#           Change <MEMO>SHORT TERM CAP GAIN REIN to <MEMO>SHORT TERM CAP GAIN PAID
#           Change <TOTAL>-#.## to <TOTAL>#.##
#           Following <SUBACCTSEC>CASH add <SUBACCTFUND>CASH
#           Delete <UNITS>0.0
#           Delete <UNITPRICE>#.##
#           Change </REINVEST> to </INCOME>

#       For long term capital gains (<INCOMETYPE>CGLONG) make these changes to the transaction:
#           Change <REINVEST> to <INCOME>
#           Change <MEMO>LONG TERM CAPITAL GAI to <MEMO>LONG TERM CAPITAL GAIN PAID
#           Change <TOTAL>-#.## to <TOTAL>#.##
#           Following <SUBACCTSEC>CASH add <SUBACCTFUND>CASH
#           Delete <UNITS>0.0
#           Delete <UNITPRICE>#.##
#           Change </REINVEST> to </INCOME>

def _scrubTRowePrice(ofx):
    global stat
    ofx_final = ''      #new ofx message
    stat = False
    if Debug: print('Function _scrubTRowePrice(OFX) called')

    #Process all <REINVEST>...</REINVEST> transactions
    #Use non-greedy quantifier .+? to avoid matching across transactions.
    p = re.compile(r'<REINVEST>.+?</REINVEST>',re.IGNORECASE)
    #re.sub() command operates on every non-overlapping occurence of pattern when passed a function for replacement
    ofx_final = p.sub(lambda r: _scrubTRowePrice_r1(r), ofx)

    if stat: scrubPrint("Scrubber: T Rowe Price dividends/capital gains paid out.")

    return ofx_final

def _scrubTRowePrice_r1(r):
    #regex subsitution function: if <UNITS>0.0 then convert transaction from <REINVEST> to <INCOME>

    global stat
    
    #Copy the reinvested transaction for manipulation
    ReinvTrans = r.group(0)
    #Create variable for the paid out transaction
    PaidTrans = ''

    # If units are 0.0 then scrub the ofx transaction
    if '<UNITS>0.0<' in ReinvTrans.upper():
        #Flag that at least one transaction is scrubbed
        stat = True

        #Use regex to parse the REINVEST transaction with following format
        #<REINVEST>...<MEMO>erroneous memo</INVTRAN>...<INCOMETYPE>DIV or CGSHORT or CGLONG<TOTAL>-#.##<SUBACCTSEC>CASH<UNITS>0.0<UNITPRICE>33.33</REINVEST>
        #into these 10 groups:
        #   m.group(1) = <REINVEST>
        #   m.group(2) = ...<MEMO>
        #   m.group(3) = erroneous memo
        #   m.group(4) = </INVTRAN>...<INCOMETYPE>
        #   m.group(5) = type of income (eg DIV, CGSHORT, CGLONG)
        #   m.group(6) = <TOTAL>-#.##
        #   m.group(7) = <SUBACCTSEC>CASH
        #   m.group(8) = <UNITS>#.###
        #   m.group(9) = <UNITPRICE>#.##
        #   m.group(10) = </REINVEST>
        p = re.compile(r'(<REINVEST>)(<.+?<MEMO>)(.+?[^<]*)(</INVTRAN>.+?<INCOMETYPE>)(.+?[^<]*)(<TOTAL>.+?[^<]*)(<SUBACCTSEC>.+?[^<]*)(<UNITS>.+?[^<]*)(<UNITPRICE>.+?[^<]*)(</REINVEST>)',re.IGNORECASE)
        m = p.match(ReinvTrans)

        gr01 = '<INCOME>'   #Change from <REINVEST>
        gr02 = m.group(2)
        gr04 = m.group(4)
        gr05 = m.group(5)
        if     gr05 == 'DIV'     : gr03 = 'DIVIDEND PAID'
        elif   gr05 == 'CGSHORT' : gr03 = 'SHORT TERM CAP GAIN PAID'
        elif   gr05 == 'CGLONG'  : gr03 = 'LONG TERM CAPITAL GAIN PAID'
        else : gr03 = m.group(3)    #Leave as reported
        gr06 = m.group(6).replace('-','')
        gr07 = m.group(7) + '<SUBACCTFUND>CASH'
        #No need to capture m.group(8) since it is deleted
        #No need to capture m.group(9) since it is deleted
        gr10 = '</INCOME>'
        PaidTrans = gr01+gr02+gr03+gr04+gr05+gr06+gr07+gr10

    if Debug: print('Reinv Trans: '+ReinvTrans)
    if Debug: print('Paid  Trans: '+PaidTrans)

    return PaidTrans             #return the new string for regex.sub()
    
# end t.rowe.price div reinvest scrubber
#-----------------------------------------------------------------------------
