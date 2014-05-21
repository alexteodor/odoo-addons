#!/usr/bin/python
# -*- coding: utf-8 -*-
#
#  Usage:
#        openerp2sie [options] database
#
#  Export an OpenERP database as a SIE file.
#
#  Bugs:
#     - Does not handle SRU codes
#     - Only SIE level 1 implemented
#     - Must run on database host by user with db access.
#     - NO TESTS!
#
#  See: http://www.sie.se [mostly in Swedish, sorry ;) ]
#
#  Type 1, Year end balance. Contains opening and closing balance for all accounts in the chart of account
#  Type 2, Period balance. The same content as in type 1 and in addition balance change per month for each account.
#  Type 3, Profit centre. The same content as in type 2 in addition balance for different profit centers.
#  Type 4, Transactions. The same content as in type 3 in addition all vouchers in the accounting year.
#  Type 4i, Transactions. Only vouchers. The purpose of this format is to make it possible to produce vouchers for different support systems and then import them in the accounting system, e g wage software or invoicing software.

import time
import os
import pwd
import argparse
import sys
import psycopg2
import datetime

today = time.strftime( "%Y%m%d")
this_year = time.strftime( "%Y" )

t = datetime.date.today()
last_year ="%d" %  (t.year - 1)


CONFIG_FILE       = "/etc/openerp-server.conf"
BALANCE_JOURNAL   = "Balance"
PROGRAM           = 'OpenERP/openerp2sie.py'
FTYP              = 'AB'

FTYP_HELP = """
AB  = Aktiebolag.
E   = Enskild näringsidkare.
HB  = Handelsbolag.
KB  = Kommanditbolag.
EK  = Ekonomisk förening.
KHF = Kooperativ hyresrättsförening.
BRF = Bostadsrättsförening.
BF  = Bostadsförening.
SF  = Sambruksförening.
I   = Ideell förening som bedriver näring.
S   = Stiftelse som bedriver näring
"""

class Config:
    def __init__(self):
        self.username = 'openerp'
        self.password = ''


class Writer:

    def write( self,  s):
        uc = unicode(s, 'utf-8')
        self._f.write( uc.encode( self._encoding))

    def close( self):
        self._f.close()

    def __init__(self, filename, encoding):
        self._f = open( filename, "w")
        self._encoding = encoding


class AccountsWriter:

    def __init__( self, cursor, type_query):
        self._cursor = cursor
        self._query = type_query

    def print_account( self, id, code):
        pass

    def print_accounts(self):
"""
TODO: Move over to OpenERP ORM!
"""
        self._cursor.execute( self._query)
        type_ids = ','.join( map( lambda x: "%d" % x[0],
                                  cursor.fetchall()))
        cursor.execute( "select id from account_account "
                         + " where user_type in(%s)" % type_ids
                         + " and type <> 'view'")
        for id in map( lambda x: x[0], cursor.fetchall()):
            q = "select code from account_account where id=%s" % id
            cursor.execute( q)
            code = cursor.fetchall()[0][0]
            self.print_account(id, code)


class ResultAccountsWriter( AccountsWriter):
"""
TODO: Move over to OpenERP ORM!
"""

    def __init__( self, cursor):
        q = ( "select id from account_account_type "
                     + " where close_method = 'none'")
        AccountsWriter.__init__(self, cursor, q)

    def print_account( self, id, code):
        q = "select sum(balance) from account_entries_report" \
            +  " where year = '%s' and account_id = %s" % ( args.year, id)
        cursor.execute( q)
        amount = cursor.fetchall()[0][0] or '0'
        print "#RES 	0	%s	%s" % ( code, amount)


class BalanceAccountsWriter( AccountsWriter):
"""
TODO: Move over to OpenERP ORM!
"""

    def __init__( self, cursor, args):
        q = ( "select id from account_account_type "
                     + " where close_method = 'balance'")
        AccountsWriter.__init__(self, cursor, q)
        self._year = args.year
        self.ib_count = 0
        q = "select id from account_journal where name='%s'" % args.journal
        cursor.execute( q)
        try:
            self._journal_id = cursor.fetchall()[0][0]
        except:
            raise Exception( "Can't find balance journal: " + args.journal)

    def print_account( self, id, code):
        q = "select sum(balance) from account_entries_report " \
            +  "where year = '%s' and account_id = %s" % ( self._year,
                                                           id)
        cursor.execute( q)
        amount = cursor.fetchall()[0][0] or '0'
        print "#UB 	0	%s	%s" % ( code, amount)
        cursor.execute( "select balance from account_entries_report"
                            + " where journal_id=%s" % self._journal_id
                            + " and date='%s-01-01'" % self._year
                            + " and account_id=%s;" % id)
        try:
            balance = cursor.fetchall()[0][0] or None
            self.ib_count += 1
            print "#IB 	0	%s	%s" % ( code, balance)
        except:
            pass

def parse_config():
    f = open( CONFIG_FILE, "r")
    lines = f.readlines()
    conf = Config()
    password = ''
    for l in lines:
         if l.find('=') == -1:
            continue
         [key, value] = l.split( '=', 1)
         key = key.strip()
         if key == 'db_user':
            conf.username = value.strip()
         if key == 'db_password':
            conf.password = value.strip()
    return conf

def parse_cmdline():

    parser = argparse.ArgumentParser(
        description='Export openerp accounts as a SIE file')
    parser.add_argument('--ftyp',
                        dest='FTYP',
                        action='store',
                        default='AB',
                        help='Företagstyp (standard: AB, flera: --help-ftyp)')
    parser.add_argument('--filename',
                        action='store',
                        default='bokslut.sa',
                        help='Filnamn (standard: bokslut.sa)')
    parser.add_argument('--help-ftyp',
                        action='store_true',
                        help='Lista företagstyper')
    parser.add_argument('--year',
                        action='store',
                        default=last_year,
                        help='Verksamhetsår att exportera (standard: %s)'
                             % last_year)
    parser.add_argument('--encoding',
                        action='store',
                        default='cp437',
                        help='Teckenkodning (endast för avlusning)')
    parser.add_argument('--balance-journal',
                        action='store',
                        dest='journal',
                        default='Balance',
                        help='Journal med ingående balans (standard: %s)'
                             % BALANCE_JOURNAL)
    parser.add_argument('database',
                         nargs='?',
                         default='',
                         help='Namn på databas som skall exporteras')

    sys.argv.pop(0)
    args = parser.parse_args(sys.argv)
    if args.help_ftyp:
        print  FTYP_HELP
        sys.exit(0)
    if args.database == '':
        print 'Databas saknas (-h för hjälp)'
        sys.exit(1)
    return args

def print_header( cursor, args):
"""
TODO: Move over to OpenERP ORM!
"""

    me = pwd.getpwuid( os.getuid())[ 0]
    today = time.strftime( "%Y%m%d")
    cursor.execute( "select name from res_company where id = 1")
    print '#FLAGGA	0'
    print '#PROGRAM	"%s"' % PROGRAM
    print '#FORMAT	PC8'
    print '#GEN		"%s %s' % ( today, me)
    print '#SIETYP	1'
    print '#FNAMN	' + cursor.fetchall()[0][0]
    print '#RAR		' + args.year
    print '#FTYP	' + FTYP


def print_account_types( cursor):
    cursor.execute( "select code, name from account_account"
                    + " where type <> 'view'")
    accounts = cursor.fetchall()
    for a in accounts:
        print '#KONTO %s	"%s"' % a


args = parse_cmdline()
conf = parse_config()
dbs = "dbname='%s' user='%s' password='%s'" % ( args.database,
                                                conf.username,
                                                conf.password)
con = psycopg2.connect( dbs)
cursor = con.cursor()

sys.stdout = Writer( args.filename, args.encoding)

print_header(cursor, args)
print_account_types(cursor)
ResultAccountsWriter( cursor).print_accounts()

balance_accounts_writer = BalanceAccountsWriter( cursor, args)
balance_accounts_writer.print_accounts()
if balance_accounts_writer.ib_count == 0:
    print >> sys.stderr, "Warning: no incoming balance found"

sys.stdout.close()
sys.stdout = sys.__stdout__
