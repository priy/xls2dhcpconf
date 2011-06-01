#!/usr/bin/python
# coding: utf-8
'''
Generator for Python

Generates dhcpd.conf from xls as a defined formate.

Usage: dumpDhcpConf.py [options] [source]

Options:
  -f ..., --file=...        use for the computer-info.xls
  -h, --help                show this help

Examples:
  dumpDhcpConf.py -f computer-info.xls  generates the dhcpd.conf

'''
__author__ = "hebin (priy031@gmail.com)"
__version__ = "$Revision: 1.3 $"
__date__ = "$Date: 2011/05/31 19:59:00 $"
__copyright__ = "Copyright (c) 2001 hebin"
__license__ = "Python"


import xlrd
import sys
import getopt


ConfString_Default = None
ConfString_ShareNetwork = None
Hosts = []
Subnets = None

def getGlobalData(sheet):
    global ConfString_Default,ConfString_ShareNetwork,Subnets

    ShareNetWorkNameColumn = [0,1]
    DefaultLeaseTimeColumn = [1,1]
    MaxLeaseTimeColumn = [2,1]
    SubnetStartRow = 4
    SubnetNetworkColumn = 1
    SubnetMaskColumn = 2
    SubnetDNSColumn = 3
    SubnetRouterColumn = 4
    SubnetRangeColumn = 5
    SubnetMaxLeaseTimeColumn = 6
    SubnetDomainName = 7

    ConfString_Default = "default-lease-time %d;\nmax-lease-time %d;\n" % (sheet.row_values(DefaultLeaseTimeColumn[0])[DefaultLeaseTimeColumn[1]],sheet.row_values(MaxLeaseTimeColumn[0])[MaxLeaseTimeColumn[1]])

    ConfString_ShareNetwork = "shared-network %s" % sheet.row_values(ShareNetWorkNameColumn[0])[ShareNetWorkNameColumn[1]]
    Subnets = []

    for i in range(SubnetStartRow,sheet.nrows):
        subnet = {}
        subnet["network"] = sheet.row_values(i)[SubnetNetworkColumn]
        subnet["netmask"] = sheet.row_values(i)[SubnetMaskColumn]
        subnet["router"] = sheet.row_values(i)[SubnetRouterColumn]
        subnet["domain-name"] = sheet.row_values(i)[SubnetDomainName]
        subnet["dns"] = sheet.row_values(i)[SubnetDNSColumn]
        if sheet.row_values(i)[SubnetRangeColumn] == "":
            subnet["range"] = None
        else:
            subnet["range"] = sheet.row_values(i)[SubnetRangeColumn]

        Subnets.append(subnet)

def excute(fname):
    global Hosts
    NameColumn = 0
    ComputerNameColumn = 1
    PortColumn = 2
    IPColumn = 3
    MacColumn = 4

    bk = xlrd.open_workbook(fname)
    shxrange = range(bk.nsheets)

    for i in shxrange:
        try:
            sh = bk.sheet_by_index(i - 1)
        except:
            return None

        if cmp(sh.name,"Global") == 0:
            getGlobalData(sh)
        else:
#            if sh.nrows > 0:
#                confFile.write(("##%s##\n" % sh.row_values(0)[0]).encode("utf-8"))
            for i in range(2,sh.nrows):
               row_data = sh.row_values(i)
               if row_data[MacColumn] != "" and row_data[ComputerNameColumn] != "":
                    host = {
                        "name" : row_data[ComputerNameColumn],
                        "uname" : row_data[NameColumn],
                        "st"    : sh.name,
                        "mac"   : row_data[MacColumn].lower(),
                        "ip"    : row_data[IPColumn]
                    }
                    Hosts.append(host)
    outPrint()

def outPrint():
    global ConfString_Default,ConfString_ShareNetwork,Subnets,Hosts
    hostRule = "\t\thost %s{\n\
\t\t\t#%s ## %s ##\n\
\t\t\thardware ethernet %s;\n\
\t\t\tfixed-address %s;\n\
\t\t}\n"
    subnetRule = "\tsubnet %s netmask %s {\n\
\t\t%srange\t%s;\n\
\t\toption\tsubnet-mask\t%s ;\n\
\t\toption\trouters\t%s ;\n\
\t\toption\tdomain-name\t\"%s\";\n\
\t\toption\tdomain-name-servers\t%s;\n"
    confFile = open("dhcp.conf.new","w")

    confFile.write(ConfString_Default)
    confFile.write(ConfString_ShareNetwork)

    confFile.write(" {\n")
    for subnet in Subnets:
        if subnet["range"] == None:
            rangeFlage = "#"
        else:
            rangeFlage = ""
        confFile.write(subnetRule % ( subnet["network"],subnet["netmask"],rangeFlage,subnet["range"],subnet["netmask"],subnet["router"],subnet["domain-name"],subnet["dns"]))
        vlan = subnet["network"][0:-2]
        l = len(vlan)
        for host in Hosts:
            if cmp(host["ip"][0:l],vlan) == 0:
                confFile.write((hostRule % (host["name"],host["uname"],host["st"],host["mac"],host["ip"])).encode("utf-8"))

        confFile.write("\t}\n")

    confFile.write("}\n")

def usage():
    print __doc__

def main(argv):
    fileName = None
    try:
        opts,args = getopt.getopt(argv, "h:f:", ["help", "file="])
    except getopt.GetoptError:
        usage()
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            usage()
            sys.exit()
        elif opt in ("-f", "--file"):
            fileName = arg
    if fileName == None:
        usage()
        sys.exit()
    else:
        excute(fileName)

if __name__ == "__main__":
    main(sys.argv[1:])
