#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Convert emails from Kmail/Kontact format to Thunderbird format.
(From Maildir format to MBOX format)

Original version downloaded from:
http://dcwww.fys.dtu.dk/~schiotz/comp/kmail2thunder.py

Rajarshi Guha, April 2003
<rajarshi@presidency.com> <http://jijo.cjb.net>

Updated by David Fenyes - voids adding the 'From' and
'Date' portions of the parsed message if they're not strings
"""

import os,sys,string,email,email.Parser,StringIO,getopt
from stat import *

localxml = """<?xml version="1.0"?>
<folderinfo>
  <folder type="mbox" name="mbox" index="1"/>
</folderinfo>
"""
folderxml = """<?xml version="1.0"?>
<efolder>
  <type>mail</type>
  <description></description>
</efolder>
"""

global logfile
global logfilename
global noconvert

def process_maildir(maildir,destdir):
    
    # Make the output evo dir and evo mbox file
    try:
        os.mkdir(destdir)
    except OSError:
        print "Couldn't create directory"+destdir

    assert destdir[-4:] == ".sbd"
    print "Creating mbox file", destdir[:-4]
    mbox = open(destdir[:-4],'w')
        
    # process the KMail maildir
    maildir = os.path.join(maildir,'cur')
    for f in os.listdir(maildir):

        tmp = os.path.join(maildir,f)

        fp = open(tmp,'r')
        p = email.Parser.Parser()
        try:
            msg = p.parse(fp)
        except:
            print ' ** Error processing a message. Logged to mail.log'
            logfile.write(tmp+'\n')

        fp.close()
        firstline = 'From '
        if type(msg['From']) == type(''):
            firstline = firstline+msg['From']
        if type(msg['Date']) == type(''):
            firstline = firstline+msg['Date']

        tmp = StringIO.StringIO()
        tmp.write(msg)
        tmp = tmp.getvalue()
        tmp = firstline+tmp[string.find(tmp,'\n')+1:]
        mbox.write(tmp)

    mbox.close()    

def process_mbox(filename,destdir):

    try:
        os.mkdir(destdir)
    except OSError:
        print "Couldn't create new directory"+destdir
        
    assert destdir[-4:] == ".sbd"
    destfile = destdir[:-4]
    
    os.system('cp "%s" "%s"' % (filename,destfile))

def main(startdir,evodir):
    
    olddir = os.getcwd()
    os.chdir(startdir)

    filelist = []
    dirlist = []
    chdirlist = []

    f = os.listdir(os.getcwd())
    for i in f:
        if string.find(i,'.index') >= 0: continue
        if i in noconvert: continue
        

        mode = os.stat(i)[ST_MODE]
        if S_ISDIR(mode):
            if string.find(i,'.') >= 0:
                if string.split(i,'.')[1] in f: 
                    chdirlist.append(string.split(i,'.')[1])
            else:
                dirlist.append(i)
        else:
            filelist.append(i)
            
    # Process all the single files
    for i in filelist:
        print 'Processing mbox: %s' % (i)
        filename = os.path.join(startdir,i)
        destdir = os.path.join(evodir,i)
        process_mbox(filename,destdir+".sbd")

    for i in dirlist:
        print 'Processing folder: %s' % (i)
        filename = os.path.join(startdir,i)
        destdir = os.path.join(evodir,i)
        process_maildir(filename,destdir+".sbd")
    
    # Now we need to recurse into the tree folders
    for i in chdirlist:

        # Check that there is indeed something under the directory we
        # are about to recurse into
        tmp = os.listdir( os.path.join(startdir,'.%s.directory' % (i)) )
        if not tmp: continue
        
        print 'Processing folders under .%s.directory' % (i)
        tk = os.path.join(startdir,'.%s.directory' % (i))
        te = os.path.join(evodir,i)
        te = te+'.sbd'
        if not os.path.exists(te):
            os.mkdir(te)
        main(tk,te)

    os.chdir(olddir)

def usage():
    print """
    Usage: kmail2evo.py [OPTIONS]

    Converts a KMail mail directory (in Maildir format) to the Evolution
    mbox fomat, maintaining folder structure. You can specify folders to
    ignore if required. The possible options are:

        -h,--help   This message
        -k,--kmail  The path to the KMail directory (default is ~/Mail)
        -e,--evo    The path to the local folder directory of the
                    Evolution mail store (default is ~/evolution/local) 
        -i,--ignore A comma separated list of folders to ignore (place
                    the list in quotes)

    By default, KMails inbox, outbox, sent-mail and drafts folders are
    ignored. To make sure that everything gets converted, specify '' to
    the -i option

    The code effectively only parses messages in Maildir format and
    simply copies mbox style folders to the corresponding Evolution
    directory. When it faces a Maildir message that it cannot parse it
    will log the message filename to mail.log

    Finally, if you have converted a *large* mail store then Evolution
    will take some time to initially load and display the messages. This
    is because this script does not do any indexing. Hence Evolution
    must create indices the first time it loads the new folders.
    """
    

if __name__ == '__main__':

    logfile = None
    logfilename = 'mail.log'
    noconvert = ['inbox','trash','drafts','sent-mail','outbox']
    kmaildir = os.environ.get('HOME')+'/Mail'
    #evodir = os.environ.get('HOME')+'/.evolution/local'
    evodir = os.environ.get('HOME')+'/tmpmail'

    if len(sys.argv) == 1:
        usage()
        sys.exit(0)
        
    try:
        opt,args = getopt.getopt(sys.argv[1:], 'hk:e:i:',\
        ['kmail=','evo=','ignore=','help'])
    except getopt.GetoptError:
        usage()
        sys.exit(0)

    for o,a in opt:
        if o in ('-h','--help'):
            usage()
            sys.exit(0)
        if o in ('-k','--kmail'):
            kmaildir = a
        if o in ('-e','--evo'):
            evodir = a
        if o in ('-i','--ignore'):
            noconvert = string.split(a,',')

    # some basic sanity checks
    if not os.path.exists(kmaildir):
        print 'Seems like %s does\'nt exist' % (kmaildir)
        sys.exit(0)
    if not os.path.exists(evodir):
        print 'Seems like %s does\'nt exist' % (evodir)
        sys.exit(0)
        
    # open the logfile 
    logfile = open(logfilename,'w')

    # start the processing
    main(os.path.abspath(kmaildir),os.path.abspath(evodir))

    logfile.close()



    
