#!python

"""
    Simplex Package Creator combines all neccessary documents into different packages

    Author: Theo Fountain III

    Prerequisites: All PDFs saved in job folder on desktop or checked out in Vault

Submittal Package
 1 Submittal Form
 2 LBD Assembly
 3 Control Box - Pictorial
 4 Electrical Drawing

Customer Package
 1 Assembly
 2 Control Box - Pictorial
 3 Customer Connections
 4 Subpanel
 5 Electrical Drawings

Release Package
 1 PRF
 2 Ship Loose BOM
 3 Electrical Drawings
 4 LBD Assembly
 5 Mech BOM
 6 Subpanel Layout
 7 Control Box -  BOM

Steps to build package

Ctrl/Environment
collect PDFs
sort
build package

"""

##        assert()
##        Check if PDF exists
                                                                                         
## Control Options
##    Local
##    Remote


## Environment Options
##    3R
##    1
                               
##if remote:
##    if envType == '3R'
##
##    if envType == '1'
# select Control box based on environmental type

## TODO:
## create keyword arguments for package types
## Add ability to check whether pdfs are missing
## Delete Simplex Reference pages 

import sys
import os
import argparse
import shutil
import re
import pdb

from pdfrw import PdfReader, PdfWriter


## REGEX

workorderPat = r'(\d{5}-\d-\d)' ## manually enter, or get from title of Quote Scrape
workorderRegex = re.compile(workorderPat)

subpanelPat = r'(Subpanel Layout)'
panelRegex = re.compile(subpanelPat)

schemaPat = r'(\d{6} LBD \d{3}-\d{3})'  ## Schematics
schemaRegex = re.compile(schemaPat)

customerPat = r'(LBD Field Connection)' ## Customer Connections
customerRegex = re.compile(customerPat)

prfPat = r'(PRF)'   ## PRF - automatically filled out from Quote Scrape
prfRegex = re.compile(prfPat)

shipLoosePat = r'(Ship Loose BOM)'
shipLooseRegex = re.compile(shipLoosePat)

submitPat = r'Bolted LBD submittal sheet' ## Comes from Quote PDF Scrape -- FIX REGEX
submitRegex = re.compile(submitPat)

assemblyPat = r'LBD Assembly'
assemblyRegex = re.compile(assemblyPat)

mechPat = r'MECH BOM'
mechRegex = re.compile(mechPat)

ctrlPictPat = r'Remote HMI'
ctrlRegex = re.compile(ctrlPictPat)


class SimPack():
    """Combines PDFs into Submittal, Customer, and Release Packages using control and environmental information. """
    
    def __init__(self,location, workorder, pkgType, ctrl, envType):
        
        self.location = location
        self.workorder = workorder
        self.pkgType = pkgType.lower()
        self.ctrl = ctrl.lower()
        self.envType = envType

        self.prf = None
        self.shipLoose = None
        self.customerConn = None
        self.schematic = None
        self.assembly = None
        self.ctrlBox = None
        self.subpanel = None
        self.ctrlPict = None
        self.ctrlBOM = None
        self.submitSheet = None

        self.pdfDict = {}
            
        if self.envType is '3R' and self.ctrl.lower is 'remote': self.getCtrlBox()
        
        self.buildPkg()
        
    def getCtrlBox(self):
        "Copiy Control Box pdf into job folder. Only for Remote Jobs"
        
        ctrlPdf = ctrlDict[self.ctrl]
        shutil(r'C:\Users\tfountain\Desktop\Control Box\{}'.format(ctrlPdf),
               r'C:\Users\tfountain\Desktop\{}'.format(self.workorder))

    def getPDF(self):
        """Return pdfs for packages in job folder in vault or on Desktop """

        if self.location == 'desktop':
            os.chdir(r'C:\\Users\tfountain\Desktop\{}'.format(self.workorder))
            
            for file in os.listdir():
                self.checkPDF(file)

        elif self.location == 'vault':
            
            os.chdir(r'C:\Simplex_Vault\Load Banks\Work Orders')
            
            for XXX in os.listdir():
                        
                if XXX.startswith(self.workorder[0:3]) and XXX.endswith('XXX'): ## 1st step
                    os.chdir(XXX)
                    
                    for XX in os.listdir():
                        
                        if XX.startswith(self.workorder[0:4]) and XX.endswith('XX'):
                            os.chdir(XX)
                            
                            for X in os.listdir():
                    
                                if X.startswith(self.workorder):
                                    os.chdir(X)
                                    folder = os.getcwd()
                                    print('Made it to {}'.format(folder))
                                    self.homeDir = folder
                                    
                                    for foldername, subfolders, filenames in os.walk(folder,topdown=True):
                                        os.chdir(foldername)
                                        print('Checking files in {}...\n'.format(foldername))
                                        print(subfolders)
                                        print(filenames)
                                   
                                        for file in filenames:
                                            self.checkPDF(file)
                                            ## add to PDFReader object/dict

            print('Done Checking. Going back to {}'.format(self.homeDir))                                    
##            os.chdir(self.homeDir)
##            Not able to save in Vault??
            
            os.chdir(r'C:\\Users\tfountain\Desktop')
      

    def checkPDF(self, file):
        """Check PDF's for thosed used in packages. """

        if file.endswith('.pdf'):
            print('Checking File: {}'.format(file))
            
            if schemaRegex.search(file):
                print(os.getcwd())
                print('Found {}'.format(file))
                self.schematic = file
                self.pdfDict[file] = PdfReader(file)

            if panelRegex.search(file):
                print(os.getcwd())
                self.subpanel = file
                self.pdfDict[file] = PdfReader(file)
                print('Found {}'.format(file))
                
            if customerRegex.search(file):
                self.customerConn = file
                self.pdfDict[file] = PdfReader(file)
                print('Found {}'.format(file))

            if prfRegex.search(file):
                self.prf = file
                self.pdfDict[file] = PdfReader(file)
                print('Found {}'.format(file))

            if submitRegex.search(file):
                self.submitSheet = file
                self.pdfDict[file] = PdfReader(file)
                print('Found {}'.format(file))

            if assemblyRegex.search(file):
                self.assembly = file
                self.pdfDict[file] = PdfReader(file)
                print('Found {}'.format(file))

            if mechRegex.search(file):
                self.mech = file
                self.pdfDict[file] = PdfReader(file)
                print('Found {}'.format(file))

            if ctrlRegex.search(file):
                self.ctrlPict = file
                self.pdfDict[file] = PdfReader(file)
                print('Found {}'.format(file))
                
    def buildPkg(self):
        """ Merge pdfs into desired package type """

        self.getPDF()
        
        submitPkg = (self.submitSheet,
            self.assembly,
            self.ctrlPict,
            self.schematic)
            
        customerPkg = (self.assembly,
            self.ctrlPict,
            self.customerConn,
            self.subpanel,
            self.schematic)
            
        releasePkg = (self.prf,
            self.shipLoose,
            self.schematic,
            self.assembly,
            self.mech,
            self.ctrlBOM)

        self.pkgDict = {
            'submittal':submitPkg, 
            'customer':customerPkg,
            'release':releasePkg,
           }

## Desktop
        
        if self.location == 'desktop':
            
            if self.pkgType == 'all':
                
                for pack in self.pkgDict:
                    print('Building {} package'.format(pack.upper()))
                    pkg = PdfWriter()

                    for pdf in self.pkgDict[pack]:
                        print(pdf)
                        
                        if len(pdf)!=0:
                            pkg.addpages(PdfReader(pdf).pages)
                            
                    pkg.write('{} {} Package.pdf'.format(self.workorder,pack.upper()))
                    print("Built {} Package".format(pack.upper()))
                    
            else:
                pkg = PdfWriter()
                
                for pdf in self.pkgDict[self.pkgType]:
                    print(pdf)
                    
                    if len(pdf)!=0:
                        pkg.addpages(PdfReader(pdf).pages)
                        
                pkg.write('{} {} Package.pdf'.format(self.workorder,self.pkgType.upper()))
                print("Built {} Package".format(self.pkgType.upper()))
## Vault
                
        else:
            
            if self.pkgType == 'all':
                print("Building all packages \n")

                for pack in self.pkgDict:
                    print('Building {} package'.format(pack.upper()))
                    pkg = PdfWriter()

                    for filename in self.pkgDict[pack]:
                        
                        if filename is not None:
                            print('Filename:{}'.format(filename))
                            pkg.addpages(self.pdfDict[filename].pages)
                            
                    pkg.write('{} {} Package.pdf'.format(self.workorder,pack.upper()))
                    print("Built {} Package\n".format(pack.upper()))

            else:
                pkg = PdfWriter()
                
                for pdf in self.pkgDict[self.pkgType]:
                    
                    
                    if pdf is not None:
                        
## Remove Simplex Reference pages from Schematic. Changes if using Terminal Block Drawings
                        print(pdf)
##                        pdb.set_trace()
                        
                        if self.pkgType == 'submittal' and pdf == self.pkgDict['submittal'][3]:
                            pkg.addpages(self.pdfDict[pdf].pages[0:-2])
                            
                        else: pkg.addpages(self.pdfDict[pdf].pages)
                        
                pkg.write('{} {} Package.pdf'.format(self.workorder,self.pkgType.upper()))
                print("Built {} Package".format(self.pkgType.upper()))
                


def main():
    
    pack = SimPack('vault', '097662', 'submittal', 'remote', '1')


##    ## Parse CLI for job info
##
##    parser = argparse.ArgumentParser(description='Build Simplex Packages')
##    parser.add_argument('job',metavar='j', type=str,help='Enter Job Number.')
##    parser.add_argument('package',metavar='p', type=str,default='all',help='Enter Package Type.')
##    parser.add_argument('env',metavar='e',type=str,help='Enter Control Type.')          
##    parser.add_argument('ctrl',metavar='e',type=str,help='Enter Environment Type.')
##
##    args = parser.parse_args()
##    print('Job Number: {} \nPackageType: {} \nEnvironment Type: {}\nControl Type: {}'.format(args.job,
##                                                                                             args.package.lower(),
##                                                                                             args.env.lower(),
##                                                                                             args.ctrl))
##
##
##    pack = SimPack(args.job, args.package.lower(), args.ctrl, args.env) ## Create Package object



if __name__ == '__main__':
    main()
