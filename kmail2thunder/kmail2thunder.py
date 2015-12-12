#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Convert emails from Kmail/Kontact format to Thunderbird format.
(From Maildir format to MBOX format)

Original version downloaded from:
http://dcwww.fys.dtu.dk/~schiotz/comp/kmail2thunder.py

Martin Zibricky, December 2015
"""

import email
import getopt
import mailbox
import os
import sys
import traceback
from stat import S_ISDIR, ST_MODE

global logfile
global logfilename
global noconvert


def process_maildir(maildir_srcdir, mbox_filename):
    """
    Process all emails in the maildir subdirectories 'cur' and 'new'.

    :param maildir_srcdir: Path to maildir directory containing 'cur' and 'new'
    :param mbox_filename: Filename
    """
    print('Creating mbox file:', mbox_filename)

    # Create directory where messages from subdirectories will be put.
    # Thunderbird file format has this structure:
    #   mbox_filename    # File with messages in MBOX format.
    #   mbox_filename.sbd    # Directory with MBOX files for subdirectories.
    mbox_subdir = mbox_filename + '.sbd'
    try:
        os.makedirs(mbox_subdir)
    except OSError:
        print("Couldn't create directory:", mbox_subdir)

    # Process one KMail maildir directory.
    with mailbox.mbox(mbox_filename) as mbox:
        # Messages are usually found in 'cur' and 'new' subdirectories.
        for subdir in ('cur', 'new'):
            d = os.path.join(maildir_srcdir, subdir)
            with mailbox.Maildir(d, email.message_from_binary_file) as mdir:
                # Iterate over messages.
                n = len(mdir)
                for index, item in enumerate(mdir.items()):
                    key, msg = item
                    if index % 10 == 9:
                        print('Progress: msg %d of %d' % (index + 1, n))
                    try:
                        mbox.add(msg)
                    except Exception:
                        print('Error while processing msg with key:', key)
                        traceback.print_exc()


def main(startdir, evodir):
    olddir = os.getcwd()
    os.chdir(startdir)

    filelist = []
    dirlist = []
    chdirlist = []

    f = os.listdir(os.getcwd())
    for i in f:
        # Skip .index files.
        if i.endswith('.index'):
            continue
        # Skip ignored files.
        if i in noconvert:
            continue

        mode = os.stat(i)[ST_MODE]
        if S_ISDIR(mode):
            if i.find('.') >= 0:
                if i.split('.')[1] in f:
                    chdirlist.append(i.split('.')[1])
            else:
                dirlist.append(i)
        else:
            filelist.append(i)

    for i in dirlist:
        print('Processing folder: %s' % i)
        filename = os.path.join(startdir, i)
        destdir = os.path.join(evodir, i)
        process_maildir(filename, destdir)

    # Now we need to recurse into the tree folders.
    for i in chdirlist:

        # Check that there is indeed something under the directory we
        # are about to recurse into
        tmp = os.listdir(os.path.join(startdir, '.%s.directory' % i))
        if not tmp:
            continue

        print('Processing folders under .%s.directory' % i)
        tk = os.path.join(startdir, '.%s.directory' % i)
        te = os.path.join(evodir, i)
        te += '.sbd'
        if not os.path.exists(te):
            os.mkdir(te)
        main(tk, te)

    os.chdir(olddir)


def usage():
    print("""
    Usage: kmail2evo.py [OPTIONS]

    Converts a KMail mail directory (in Maildir format) to the Evolution
    mbox fomat, maintaining folder structure. You can specify folders to
    ignore if required. The possible options are:

        -h,--help    This message
        -k,--kmail   The path to the KMail directory
        -t,--thunder The path to the local folder directory of the
                     Thunderbird mail store
        -i,--ignore  A comma separated list of folders to ignore (place
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
    """)


if __name__ == '__main__':

    logfile = None
    logfilename = 'mail.log'
    noconvert = ['inbox', 'trash', 'drafts', 'sent-mail', 'outbox']

    if len(sys.argv) == 1:
        usage()
        sys.exit(0)

    try:
        opt, args = getopt.getopt(sys.argv[1:], 'hk:t:i:',
                                  ['kmail=', 'evo=', 'ignore=', 'help'])
    except getopt.GetoptError:
        usage()
        sys.exit(0)

    kmaildir = thunderdir = None

    for o, a in opt:
        if o in ('-h', '--help'):
            usage()
            sys.exit(0)
        if o in ('-k', '--kmail'):
            kmaildir = a
        if o in ('-t', '--thunder'):
            thunderdir = a
        if o in ('-i', '--ignore'):
            noconvert = a.split(',')

    # some basic sanity checks
    if not os.path.exists(kmaildir):
        print('Seems like %s does\'nt exist' % kmaildir)
        sys.exit(1)
    if not os.path.exists(thunderdir):
        print('Seems like %s does\'nt exist' % thunderdir)
        sys.exit(1)

    # open the logfile 
    logfile = open(logfilename, 'w')

    # start the processing
    main(os.path.abspath(kmaildir), os.path.abspath(thunderdir))

    logfile.close()
