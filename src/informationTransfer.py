from datetime import datetime
import pandas as pd
from pathlib import Path
import openpyxl
import os
import time


class vEfec:
    def __init__(self):         
        self.bill = efBill
        self.moneda = efBill


class efBill:
    def __init__(self):
        self.corte = None
        self.cantidad = None
        self.total = None


class vouCard:
    def __init__(self):
        self.fndCambios = None
        self.impDep = None

class vales:
            pass
class qr:
            pass

class aclDif:
            pass

class informationTransfer():

    def __init__(self):

        self.vEfec = vEfec #tabla
        self.cCaja = None
        self.cash  = None
        self.aCont = None
        self.dCont = None
        self.tCont = None
        self.pCred = None
        self.oCred = None
        self.tCred = None
        self.tVent = None
        self.efBill = efBill #tabla
        # self.fndCamb = None
        # self.impDep = None
        self.rcntTBs = None
        self.tEfeBs = None
        self.vouCard = vouCard #tabla
        self.vales = vales #tabla
        self.qr = qr #tabla
        self.aclDif = aclDif #tabla


    def xlsxFormating(self, xlsxName):
        
        wb = openpyxl.load_workbook(xlsxName)
        ws = wb.active


        
        
        