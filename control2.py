# control2.py
# http://sites.google.com/site/pocketsense/
# contains some common configuration data and modules for the ofx pkg
# Initial version: rlc: Feb-2010
#
#04-Jan-2010*rlc
#   - Added DefaultAppID and DefaultAppVer

# 27-Jul-2013: rlc
#   - Added locale support 
# 03-Sep-2014: rlc
#   - xfrdir is now platform independent
# 06-Mar-2017: rlc
#   - Set DefaultAppVer = 2400 (Quicken 2015)
#   - Moved utility functions to rlib1 module
#------------------------------------------------------------------------------------

#---MODULES---
import os

Debug = False             #debug mode = true only when testing
#Debug = True

AboutTitle    = 'PocketSense OFX Download Python Scripts'
AboutVersion  = '24-Aug-2018'
AboutSource   = 'http://sites.google.com/site/pocketsense'
AboutName     = 'Robert'

#xfrdir = temp directory for statement downloads.  Platform independent
xfrdir   = os.path.join(os.path.curdir,"xfr") + os.sep
if Debug: print "XFRDIR = " + xfrdir
cfgFile  = 'ofx_config.cfg'    #user account settings (can be encrypted)

DefaultAppID  = 'QWIN'
DefaultAppVer = '2400'


    


