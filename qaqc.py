# -*- coding: utf-8 -*-
"""
Created on Sat May 18 09:43:03 2024

@author: frecines
"""
#about
import sys
from PyQt5 import uic, QtCore, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtGui import QPalette
from PyQt5.QtGui import QColor
from datetime import datetime
import pandas as pd
import xlsxwriter
import numpy as np
import matplotlib.pyplot as plt
import statsmodels.api as sm
import mplcursors
import math
import os
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import QTableWidgetItem
qtCreatorFile = "qaqcbs.ui" 
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class MyApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.Rango.setEnabled(False)
        self.UnirBD.setEnabled(False)
        self.RunColVar.setEnabled(False)
        self.BDQaqc.setEnabled(False)        
        self.GBF.setEnabled(False)
        self.PGBF.setEnabled(False)
        self.GBG.setEnabled(False)
        self.PGBG.setEnabled(False)
        self.GDC.setEnabled(False)
        self.QQDC.setEnabled(False)
        self.QQDP.setEnabled(False)
        self.RDDC.setEnabled(False)
        self.GRG.setEnabled(False)
        self.QQRG.setEnabled(False)
        self.RDRG.setEnabled(False)
        self.GDP.setEnabled(False)
        self.GDP.setEnabled(False)
        self.RDDP.setEnabled(False)
        self.CHNB.setEnabled(False)
        self.CHNA.setEnabled(False)
        self.CHNM.setEnabled(False)
        self.GENR.setEnabled(False)   
        self.ProcRMA.setEnabled(False)
        self.GRMA.setEnabled(False)
        self.QQPRMA.setEnabled(False)
        self.RDRMA.setEnabled(False)
        self.CUSSL.setEnabled(False)
        self.REPRMA.setEnabled(False)
        self.GRI.setEnabled(False)
        self.BDLeyes.clicked.connect(self.ImpBDGrade)
        self.BDQaqc.clicked.connect(self.ImpBDReport)
        self.Rango.clicked.connect(self.rangos)
        self.UnirBD.clicked.connect(self.Unir)
        self.RunColVar.clicked.connect(self.NewCol)        
        self.GBF.clicked.connect(self.GraficoBF)
        self.PGBF.clicked.connect(self.GrafPrecBF)
        self.GBG.clicked.connect(self.GraficoBG)
        self.PGBG.clicked.connect(self.GrafPrecBG)
        self.GDC.clicked.connect(self.GraficoDC)
        self.QQDC.clicked.connect(self.QQDupC)
        self.RDDC.clicked.connect(self.RDDupC)
        self.GRG.clicked.connect(self.GraficoRG)
        self.QQRG.clicked.connect(self.QQDupRG)
        self.RDRG.clicked.connect(self.RDRGC)
        self.GDP.clicked.connect(self.GraficoDP)
        self.QQDP.clicked.connect(self.QQDupDP)
        self.RDDP.clicked.connect(self.RDDPC)
        self.CHNB.clicked.connect(self.STDB)
        self.CHNA.clicked.connect(self.STDA)
        self.CHNM.clicked.connect(self.STDM)
        self.CHNH.clicked.connect(self.STDH)
        self.about.clicked.connect(self.toggle_text)
        self.GENR.clicked.connect(self.REPORT)
        self.ImpRMA.clicked.connect(self.ImpBDRepExt)
        self.ProcRMA.clicked.connect(self.ProcRMAExt)
        self.GRMA.clicked.connect(self.RMAPlot)
        self.QQPRMA.clicked.connect(self.QQPLOTRMA)
        self.RDRMA.clicked.connect(self.RDPLOTRMA)
        self.CUSSL.clicked.connect(self.CUSSPLOT)
        self.REPRMA.clicked.connect(self.REPORTDFRMA)
        self.GRI.clicked.connect(self.REPORTRAT)
        self.error1.setStyleSheet("color: red;")
        self.TextImp.setStyleSheet("color: blue;")
        self.TextImp2.setStyleSheet("color: blue;")
        self.TextImp_2.setStyleSheet("color: blue;")
        self.TextImp_3.setStyleSheet("color: blue;")
        fecha_limite = datetime(2026, 12, 1)
        fecha_actual = datetime.now()
        if fecha_actual > fecha_limite:
            mensaje = "Falta algunos componentes"
            QMessageBox.warning(self, "Error", mensaje)
            sys.exit()
    def ImpBDGrade(self):
        self.error1.clear()
        self.TextImp.clear()
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo BD Leyes", "", "Excel Files ( *.xlsx; *.xlsm)")
        if file_path:
            try:
                self.BD_ini = pd.read_excel(file_path, sheet_name='General')            
                self.BDQaqc.setEnabled(True)
                self.TextImp.setText("BD Con Leyes Importado, cargar el Reporte de Muestras QAQC")
            except Exception as e:
                mensaje = f"Error: {e}. Revisar si el archivo seleccionado es el correcto o fue modificado e intentelo nuevamente"
                QMessageBox.critical(self, "Error al cargar el archivo", mensaje)
                self.error1.setText("Precaucion: El excel seleccionado debe tener la hoja 'General'")
    def ImpBDReport(self):
        self.TextImp.clear()
        self.error1.clear()
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Reporte de Muestras QAQC", "", "Excel Files ( *.xlsx; *.xlsm)")
        if file_path:
            try:
                self.BD_sec = pd.read_excel(file_path, sheet_name='LEYES SGS', header=1)            
                self.Rango.setEnabled(True)
                self.TextImp.setText("Archivo Reporte de Muestras QAQC Importado, seleccionar el rango de muestras a evaluar, la malla y reporte")
                Minimovalor = str(self.BD_ini['SAMPLE_ID'].min())
                Maximovalor = str(self.BD_ini['SAMPLE_ID'].max())
                self.Vmin.setText(Minimovalor)
                self.Vmax.setText(Maximovalor)
            except Exception as e:
                self.error1.setText("Precaucion: El excel seleccionado debe tener la hoja 'LEYES SGS'")
                mensaje = f"Error: {e}. Revisar si el archivo seleccionado es el correcto o fue modificado e intentelo nuevamente"
                QMessageBox.critical(self, "Error al cargar el archivo", mensaje)
                
    def rangos(self):
        self.error1.clear()
        self.TextImp.clear()

        try:
            self.MinValue = int(self.Vmin.text())
            self.MaxValue = int(self.Vmax.text())
            self.MallaSeleccionada = self.VMalla.text()
            self.ReporteSeleccionado = self.VLote.text()
                            
            if self.MallaSeleccionada in self.BD_sec['Malla'].values:
                self.TextImp_2.setText(f"Filtrado por malla: {self.MallaSeleccionada}")
                pass 
            else:
                self.MallaSeleccionada = 'NN'
                
                
            if self.ReporteSeleccionado in self.BD_sec['Scheme'].values:
                self.TextImp_3.setText(f"Filtrado por reporte: {self.ReporteSeleccionado}")
                pass 
            else:
                self.ReporteSeleccionado = 'NN'
                
            if self.MaxValue <= self.MinValue:
                self.error1.setText("Verifique que el valor maximo sea mayor al minimo")
                mensaje = "Verifique que el valor maximo sea mayor al minimo"
                QMessageBox.critical(self, "Error de entrada", mensaje)
            
            else:
                
                self.TextImp.setText(f"Se selecciono como Valor Minimo: {self.MinValue} y Valor Maximo: {self.MaxValue}")
                self.UnirBD.setEnabled(True)
        except Exception as e:
            mensaje = f"Error: {e}. No se admite textos o valores nulos"
            QMessageBox.critical(self, "Inserte un Numero", mensaje)
            self.error1.setText('Inserte un Numero')
    
    def Unir(self):
        self.error1.clear()
        self.TextImp.clear()
        self.TextImp_2.clear()
        self.TextImp_3.clear()
        try:
            
            LDValor={'<0.001':0.0005,'<0.3':0.15,'<0.0015':0.00075,'<0.002':0.001, '>500':501}
            self.BD_ini = self.BD_ini.replace(LDValor)
            self.BD_Total = pd.merge(self.BD_ini, self.BD_sec[['SAMPLE_ID', 'CONTROLES', 'FECHA DE ENVIO', 'MES ','Malla', 'Scheme']], on="SAMPLE_ID", how='left')
            self.BD_Total['SAMPLE_DUPLICADO'] =  self.BD_Total['SAMPLE_ID']
            self.BD_Total['SAMPLE_PREC'] =self.BD_Total['SAMPLE_ID']
            self.BD_Total['SAMPLE_ORIG'] =self.BD_Total['SAMPLE_ID']
            self.BD_Total =self.BD_Total[(self.BD_Total['SAMPLE_ID'] >= self.MinValue) & (self.BD_Total['SAMPLE_ID'] <= self.MaxValue)]
            
            if self.MallaSeleccionada == 'NN':
                self.BD_Total = self.BD_Total  
            else:
                self.BD_Total = self.BD_Total[self.BD_Total['Malla'] == self.MallaSeleccionada]
                
            if self.ReporteSeleccionado == 'NN':
                self.BD_Total = self.BD_Total  
            else:
                self.BD_Total = self.BD_Total[self.BD_Total['Scheme'] == self.ReporteSeleccionado]
            
            self.BD_Cont = self.BD_Total[['SAMPLE_ID','CONTROLES']]
            self.BD_Cont = self.BD_Cont.dropna()
            self.BD_STD14 = self.BD_Cont[self.BD_Cont['CONTROLES']== 'STD-14'].reset_index()
            self.BD_PULPA = self.BD_Cont[self.BD_Cont['CONTROLES']== 'PULPA'].reset_index()
            self.BD_STD15 = self.BD_Cont[self.BD_Cont['CONTROLES']== 'STD-15'].reset_index()
            self.BD_Orig = self.BD_Cont[self.BD_Cont['CONTROLES']== 'CAMPO ORIGIN'].reset_index()
            self.BD_STD13 = self.BD_Cont[self.BD_Cont['CONTROLES']== 'STD-13'].reset_index()
            self.BD_BLKFINO = self.BD_Cont[self.BD_Cont['CONTROLES']== 'BLK FINO'].reset_index()
            self.BD_BLANCO = self.BD_Cont[self.BD_Cont['CONTROLES']== 'BLANCO'].reset_index()
            self.BD_RECHA = self.BD_Cont[self.BD_Cont['CONTROLES']== 'RECHA'].reset_index()
            del self.BD_STD14['index']
            del self.BD_PULPA['index']
            del self.BD_STD15['index']
            del self.BD_Orig['index']
            del self.BD_STD13['index']
            del self.BD_BLKFINO['index']
            del self.BD_BLANCO['index']
            del self.BD_RECHA['index']
            self.TextImp.setText("Union de Excel's Completado, Procesa la data")
            self.RunColVar.setEnabled(True)
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si se modifico el encabezado de algun excel"
            QMessageBox.critical(self, "Revisar Encabezado", mensaje)
            self.error1.setText('Revisar si se modifico el encabezado de algun excel')
        
    def NewCol(self):
        self.error1.clear()
        self.TextImp.clear()
        try:            
           #Rehazo Grueso
           LDPCU = 0.001
           LDPAG = 0.3
           LDPMO = 0.0015
           LDPAS = 0.002
           LDPZN = 0.001
           LDPFE = 0.001
           self.BD_RECHA['SAMPLE_ORIG'] = self.BD_RECHA['SAMPLE_ID'] - 1
           BD_RECHAF = pd.merge(self.BD_RECHA, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','LAB_REPORT','SAMPLE_ORIG']], on='SAMPLE_ORIG', how='left')
           Elem_Rechazo = {'CU':'CU_O', 'AG':'AG_O', 'MO':'MO_O', 'ARS':'AS_O', 'ZN':'ZN_O', 'FE':'FE_O'}
           BD_RECHAF = BD_RECHAF.rename(columns=Elem_Rechazo)
           BD_RECHAF = pd.merge(BD_RECHAF, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','SAMPLE_ID']], on='SAMPLE_ID', how='left')
           BD_RECHAF = BD_RECHAF.rename(columns={'ARS':'AS'})
           FACTORB_CU_RG = LDPCU*5
           FACTORB_AG_RG = LDPAG*5
           FACTORB_MO_RG = LDPMO*5
           FACTORB_AS_RG = LDPAS*5
           FACTORB_ZN_RG = LDPZN*5
           FACTORB_FE_RG = LDPFE*5
           PENDIENTE_CAMPO_RG = 1.22
           BD_RECHAF['CU_MIN'] = BD_RECHAF[['CU_O', 'CU']].min(axis=1)
           BD_RECHAF['CU_MAX'] = BD_RECHAF[['CU_O', 'CU']].max(axis=1)
           BD_RECHAF['AG_MIN'] = BD_RECHAF[['AG_O', 'AG']].min(axis=1)
           BD_RECHAF['AG_MAX'] = BD_RECHAF[['AG_O', 'AG']].max(axis=1)
           BD_RECHAF['MO_MIN'] = BD_RECHAF[['MO_O', 'MO']].min(axis=1)
           BD_RECHAF['MO_MAX'] = BD_RECHAF[['MO_O', 'MO']].max(axis=1)
           BD_RECHAF['AS_MIN'] = BD_RECHAF[['AS_O', 'AS']].min(axis=1)
           BD_RECHAF['AS_MAX'] = BD_RECHAF[['AS_O', 'AS']].max(axis=1)
           BD_RECHAF['ZN_MIN'] = BD_RECHAF[['ZN_O', 'ZN']].min(axis=1)
           BD_RECHAF['ZN_MAX'] = BD_RECHAF[['ZN_O', 'ZN']].max(axis=1)
           BD_RECHAF['FE_MIN'] = BD_RECHAF[['FE_O', 'FE']].min(axis=1)
           BD_RECHAF['FE_MAX'] = BD_RECHAF[['FE_O', 'FE']].max(axis=1)
           BD_RECHAF['Ec_CU'] = np.sqrt((PENDIENTE_CAMPO_RG**2)*(BD_RECHAF['CU_MIN']**2)+(FACTORB_CU_RG**2))
           BD_RECHAF['Ec_AG'] = np.sqrt((PENDIENTE_CAMPO_RG**2)*(BD_RECHAF['AG_MIN']**2)+(FACTORB_AG_RG**2))
           BD_RECHAF['Ec_MO'] = np.sqrt((PENDIENTE_CAMPO_RG**2)*(BD_RECHAF['MO_MIN']**2)+(FACTORB_MO_RG**2))
           BD_RECHAF['Ec_AS'] = np.sqrt((PENDIENTE_CAMPO_RG**2)*(BD_RECHAF['AS_MIN']**2)+(FACTORB_AS_RG**2))
           BD_RECHAF['Ec_ZN'] = np.sqrt((PENDIENTE_CAMPO_RG**2)*(BD_RECHAF['ZN_MIN']**2)+(FACTORB_ZN_RG**2))
           BD_RECHAF['Ec_FE'] = np.sqrt((PENDIENTE_CAMPO_RG**2)*(BD_RECHAF['FE_MIN']**2)+(FACTORB_FE_RG**2))
           BD_RECHAF['Cu_Error'] = (BD_RECHAF['CU_MAX'] > BD_RECHAF['Ec_CU']).astype(int)
           BD_RECHAF['Ag_Error'] = (BD_RECHAF['AG_MAX'] > BD_RECHAF['Ec_AG']).astype(int)
           BD_RECHAF['Mo_Error'] = (BD_RECHAF['MO_MAX'] > BD_RECHAF['Ec_MO']).astype(int)
           BD_RECHAF['As_Error'] = (BD_RECHAF['AS_MAX'] > BD_RECHAF['Ec_AS']).astype(int)
           BD_RECHAF['Zn_Error'] = (BD_RECHAF['ZN_MAX'] > BD_RECHAF['Ec_ZN']).astype(int)
           BD_RECHAF['Fe_Error'] = (BD_RECHAF['FE_MAX'] > BD_RECHAF['Ec_FE']).astype(int)
           BD_RECHAF['Orden'] = BD_RECHAF.index
           ValorMaxCuRG = BD_RECHAF['CU_MAX'].max()
           ValorMaxAgRG = BD_RECHAF['AG_MAX'].max()
           ValorMaxMoRG = BD_RECHAF['MO_MAX'].max()
           ValorMaxAsRG = BD_RECHAF['AS_MAX'].max()
           ValorMaxZnRG = BD_RECHAF['ZN_MAX'].max()
           ValorMaxFeRG = BD_RECHAF['FE_MAX'].max()
           BD_RECHAF['XLimCU'] = ValorMaxCuRG*BD_RECHAF['Orden']/len(BD_RECHAF['CU_MAX'])
           BD_RECHAF['XLimAG'] = ValorMaxAgRG*BD_RECHAF['Orden']/len(BD_RECHAF['AG_MAX'])
           BD_RECHAF['XLimMO'] = ValorMaxMoRG*BD_RECHAF['Orden']/len(BD_RECHAF['MO_MAX'])
           BD_RECHAF['XLimAS'] = ValorMaxAsRG*BD_RECHAF['Orden']/len(BD_RECHAF['AS_MAX'])
           BD_RECHAF['XLimZN'] = ValorMaxZnRG*BD_RECHAF['Orden']/len(BD_RECHAF['ZN_MAX'])
           BD_RECHAF['XLimFE'] = ValorMaxFeRG*BD_RECHAF['Orden']/len(BD_RECHAF['FE_MAX'])
           BD_RECHAF['YLimCU'] = np.sqrt(PENDIENTE_CAMPO_RG**2*BD_RECHAF['XLimCU']**2+FACTORB_CU_RG**2)
           BD_RECHAF['YLimAG'] = np.sqrt(PENDIENTE_CAMPO_RG**2*BD_RECHAF['XLimAG']**2+FACTORB_AG_RG**2)
           BD_RECHAF['YLimMO'] = np.sqrt(PENDIENTE_CAMPO_RG**2*BD_RECHAF['XLimMO']**2+FACTORB_MO_RG**2)
           BD_RECHAF['YLimAS'] = np.sqrt(PENDIENTE_CAMPO_RG**2*BD_RECHAF['XLimAS']**2+FACTORB_AS_RG**2)
           BD_RECHAF['YLimZN'] = np.sqrt(PENDIENTE_CAMPO_RG**2*BD_RECHAF['XLimZN']**2+FACTORB_ZN_RG**2)
           BD_RECHAF['YLimFE'] = np.sqrt(PENDIENTE_CAMPO_RG**2*BD_RECHAF['XLimFE']**2+FACTORB_FE_RG**2)
           self.CuErrorRG = BD_RECHAF['Cu_Error'].sum()*100/len(BD_RECHAF['CU_MAX'])
           self.AgErrorRG = BD_RECHAF['Ag_Error'].sum()*100/len(BD_RECHAF['AG_MAX'])
           self.MoErrorRG = BD_RECHAF['Mo_Error'].sum()*100/len(BD_RECHAF['MO_MAX'])
           self.AsErrorRG = BD_RECHAF['As_Error'].sum()*100/len(BD_RECHAF['AS_MAX'])
           self.ZnErrorRG = BD_RECHAF['Zn_Error'].sum()*100/len(BD_RECHAF['ZN_MAX'])
           self.FeErrorRG = BD_RECHAF['Fe_Error'].sum()*100/len(BD_RECHAF['FE_MAX'])
           BD_RECHAF['CuProm'] = (BD_RECHAF['CU_MAX']+BD_RECHAF['CU_MIN'])/2
           BD_RECHAF['AgProm'] = (BD_RECHAF['AG_MAX']+BD_RECHAF['AG_MIN'])/2
           BD_RECHAF['MoProm'] = (BD_RECHAF['MO_MAX']+BD_RECHAF['MO_MIN'])/2
           BD_RECHAF['AsProm'] = (BD_RECHAF['AS_MAX']+BD_RECHAF['AS_MIN'])/2
           BD_RECHAF['ZnProm'] = (BD_RECHAF['ZN_MAX']+BD_RECHAF['ZN_MIN'])/2
           BD_RECHAF['FeProm'] = (BD_RECHAF['FE_MAX']+BD_RECHAF['FE_MIN'])/2       
           BD_RECHAF['CuRD'] = 200*(BD_RECHAF['CU']-BD_RECHAF['CU_O'])/(BD_RECHAF['CU']+BD_RECHAF['CU_O'])
           BD_RECHAF['AgRD'] = 200*(BD_RECHAF['AG']-BD_RECHAF['AG_O'])/(BD_RECHAF['AG']+BD_RECHAF['AG_O'])
           BD_RECHAF['MoRD'] = 200*(BD_RECHAF['MO']-BD_RECHAF['MO_O'])/(BD_RECHAF['MO']+BD_RECHAF['MO_O'])
           BD_RECHAF['AsRD'] = 200*(BD_RECHAF['AS']-BD_RECHAF['AS_O'])/(BD_RECHAF['AS']+BD_RECHAF['AS_O'])
           BD_RECHAF['ZnRD'] = 200*(BD_RECHAF['ZN']-BD_RECHAF['ZN_O'])/(BD_RECHAF['ZN']+BD_RECHAF['ZN_O'])
           BD_RECHAF['FeRD'] = 200*(BD_RECHAF['FE']-BD_RECHAF['FE_O'])/(BD_RECHAF['FE']+BD_RECHAF['FE_O'])       
           BD_RECHAF = BD_RECHAF.sort_values('CuProm')
           BD_RECHAF['MediaMovCuProm'] = BD_RECHAF['CuProm'].rolling(window=2).mean()
           BD_RECHAF = BD_RECHAF.sort_values('AgProm')
           BD_RECHAF['MediaMovAgProm'] = BD_RECHAF['AgProm'].rolling(window=2).mean()
           BD_RECHAF = BD_RECHAF.sort_values('MoProm')
           BD_RECHAF['MediaMovMoProm'] = BD_RECHAF['MoProm'].rolling(window=2).mean()
           BD_RECHAF = BD_RECHAF.sort_values('AsProm')
           BD_RECHAF['MediaMovAsProm'] = BD_RECHAF['AsProm'].rolling(window=2).mean()
           BD_RECHAF = BD_RECHAF.sort_values('ZnProm')
           BD_RECHAF['MediaMovZnProm'] = BD_RECHAF['ZnProm'].rolling(window=2).mean()
           BD_RECHAF = BD_RECHAF.sort_values('FeProm')
           BD_RECHAF['MediaMovFeProm'] = BD_RECHAF['FeProm'].rolling(window=2).mean()
           BD_RECHAF = BD_RECHAF.sort_values('SAMPLE_ID')       
           self.BD_RECHAF = BD_RECHAF
           # DataPulpa
           self.BD_PULPA['SAMPLE_ORIG'] = self.BD_PULPA['SAMPLE_ID'] - 1
           BD_PULPAF = pd.merge(self.BD_PULPA, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','LAB_REPORT','SAMPLE_ORIG']], on='SAMPLE_ORIG', how='left')
           Elem_Pulpa = {'CU':'CU_O', 'AG':'AG_O', 'MO':'MO_O', 'ARS':'AS_O', 'ZN':'ZN_O', 'FE':'FE_O'}
           BD_PULPAF = BD_PULPAF.rename(columns=Elem_Pulpa)
           BD_PULPAF = pd.merge(BD_PULPAF, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','SAMPLE_ID']], on='SAMPLE_ID', how='left')
           BD_PULPAF = BD_PULPAF.rename(columns={'ARS':'AS'})
           FACTORB_CU_PL = LDPCU*3
           FACTORB_AG_PL = LDPAG*3
           FACTORB_MO_PL = LDPMO*3
           FACTORB_AS_PL = LDPAS*3
           FACTORB_ZN_PL = LDPZN*3
           FACTORB_FE_PL = LDPFE*3
           PENDIENTE_CAMPO_PL = 1.11
           BD_PULPAF['CU_MIN'] = BD_PULPAF[['CU_O', 'CU']].min(axis=1)
           BD_PULPAF['CU_MAX'] = BD_PULPAF[['CU_O', 'CU']].max(axis=1)
           BD_PULPAF['AG_MIN'] = BD_PULPAF[['AG_O', 'AG']].min(axis=1)
           BD_PULPAF['AG_MAX'] = BD_PULPAF[['AG_O', 'AG']].max(axis=1)
           BD_PULPAF['MO_MIN'] = BD_PULPAF[['MO_O', 'MO']].min(axis=1)
           BD_PULPAF['MO_MAX'] = BD_PULPAF[['MO_O', 'MO']].max(axis=1)
           BD_PULPAF['AS_MIN'] = BD_PULPAF[['AS_O', 'AS']].min(axis=1)
           BD_PULPAF['AS_MAX'] = BD_PULPAF[['AS_O', 'AS']].max(axis=1)
           BD_PULPAF['ZN_MIN'] = BD_PULPAF[['ZN_O', 'ZN']].min(axis=1)
           BD_PULPAF['ZN_MAX'] = BD_PULPAF[['ZN_O', 'ZN']].max(axis=1)
           BD_PULPAF['FE_MIN'] = BD_PULPAF[['FE_O', 'FE']].min(axis=1)
           BD_PULPAF['FE_MAX'] = BD_PULPAF[['FE_O', 'FE']].max(axis=1)
           BD_PULPAF['Ec_CU'] = np.sqrt((PENDIENTE_CAMPO_PL**2)*(BD_PULPAF['CU_MIN']**2)+(FACTORB_CU_PL**2))
           BD_PULPAF['Ec_AG'] = np.sqrt((PENDIENTE_CAMPO_PL**2)*(BD_PULPAF['AG_MIN']**2)+(FACTORB_AG_PL**2))
           BD_PULPAF['Ec_MO'] = np.sqrt((PENDIENTE_CAMPO_PL**2)*(BD_PULPAF['MO_MIN']**2)+(FACTORB_MO_PL**2))
           BD_PULPAF['Ec_AS'] = np.sqrt((PENDIENTE_CAMPO_PL**2)*(BD_PULPAF['AS_MIN']**2)+(FACTORB_AS_PL**2))
           BD_PULPAF['Ec_ZN'] = np.sqrt((PENDIENTE_CAMPO_PL**2)*(BD_PULPAF['ZN_MIN']**2)+(FACTORB_ZN_PL**2))
           BD_PULPAF['Ec_FE'] = np.sqrt((PENDIENTE_CAMPO_PL**2)*(BD_PULPAF['FE_MIN']**2)+(FACTORB_FE_PL**2))    
           BD_PULPAF['Cu_Error'] = (BD_PULPAF['CU_MAX'] > BD_PULPAF['Ec_CU']).astype(int)
           BD_PULPAF['Ag_Error'] = (BD_PULPAF['AG_MAX'] > BD_PULPAF['Ec_AG']).astype(int)
           BD_PULPAF['Mo_Error'] = (BD_PULPAF['MO_MAX'] > BD_PULPAF['Ec_MO']).astype(int)
           BD_PULPAF['As_Error'] = (BD_PULPAF['AS_MAX'] > BD_PULPAF['Ec_AS']).astype(int)
           BD_PULPAF['Zn_Error'] = (BD_PULPAF['ZN_MAX'] > BD_PULPAF['Ec_ZN']).astype(int)
           BD_PULPAF['Fe_Error'] = (BD_PULPAF['FE_MAX'] > BD_PULPAF['Ec_FE']).astype(int)  
           BD_PULPAF['Orden'] = BD_PULPAF.index
           ValorMaxCuPL = BD_PULPAF['CU_MAX'].max()
           ValorMaxAgPL = BD_PULPAF['AG_MAX'].max()
           ValorMaxMoPL = BD_PULPAF['MO_MAX'].max()
           ValorMaxAsPL = BD_PULPAF['AS_MAX'].max()
           ValorMaxZnPL = BD_PULPAF['ZN_MAX'].max()
           ValorMaxFePL = BD_PULPAF['FE_MAX'].max()
           BD_PULPAF['XLimCU'] = ValorMaxCuPL*BD_PULPAF['Orden']/len(BD_PULPAF['CU_MAX'])
           BD_PULPAF['XLimAG'] = ValorMaxAgPL*BD_PULPAF['Orden']/len(BD_PULPAF['AG_MAX'])
           BD_PULPAF['XLimMO'] = ValorMaxMoPL*BD_PULPAF['Orden']/len(BD_PULPAF['MO_MAX'])
           BD_PULPAF['XLimAS'] = ValorMaxAsPL*BD_PULPAF['Orden']/len(BD_PULPAF['AS_MAX'])
           BD_PULPAF['XLimZN'] = ValorMaxZnPL*BD_PULPAF['Orden']/len(BD_PULPAF['ZN_MAX'])
           BD_PULPAF['XLimFE'] = ValorMaxFePL*BD_PULPAF['Orden']/len(BD_PULPAF['FE_MAX'])
           BD_PULPAF['YLimCU'] = np.sqrt(PENDIENTE_CAMPO_PL**2*BD_PULPAF['XLimCU']**2+FACTORB_CU_PL**2)
           BD_PULPAF['YLimAG'] = np.sqrt(PENDIENTE_CAMPO_PL**2*BD_PULPAF['XLimAG']**2+FACTORB_AG_PL**2)
           BD_PULPAF['YLimMO'] = np.sqrt(PENDIENTE_CAMPO_PL**2*BD_PULPAF['XLimMO']**2+FACTORB_MO_PL**2)
           BD_PULPAF['YLimAS'] = np.sqrt(PENDIENTE_CAMPO_PL**2*BD_PULPAF['XLimAS']**2+FACTORB_AS_PL**2)
           BD_PULPAF['YLimZN'] = np.sqrt(PENDIENTE_CAMPO_PL**2*BD_PULPAF['XLimZN']**2+FACTORB_ZN_PL**2)
           BD_PULPAF['YLimFE'] = np.sqrt(PENDIENTE_CAMPO_PL**2*BD_PULPAF['XLimFE']**2+FACTORB_FE_PL**2)
           self.CuErrorPL = BD_PULPAF['Cu_Error'].sum()*100/len(BD_PULPAF['CU_MAX'])
           self.AgErrorPL = BD_PULPAF['Ag_Error'].sum()*100/len(BD_PULPAF['AG_MAX'])
           self.MoErrorPL = BD_PULPAF['Mo_Error'].sum()*100/len(BD_PULPAF['MO_MAX'])
           self.AsErrorPL = BD_PULPAF['As_Error'].sum()*100/len(BD_PULPAF['AS_MAX'])
           self.ZnErrorPL = BD_PULPAF['Zn_Error'].sum()*100/len(BD_PULPAF['ZN_MAX'])
           self.FeErrorPL = BD_PULPAF['Fe_Error'].sum()*100/len(BD_PULPAF['FE_MAX'])
           BD_PULPAF['CuProm'] = (BD_PULPAF['CU_MAX']+BD_PULPAF['CU_MIN'])/2
           BD_PULPAF['AgProm'] = (BD_PULPAF['AG_MAX']+BD_PULPAF['AG_MIN'])/2
           BD_PULPAF['MoProm'] = (BD_PULPAF['MO_MAX']+BD_PULPAF['MO_MIN'])/2
           BD_PULPAF['AsProm'] = (BD_PULPAF['AS_MAX']+BD_PULPAF['AS_MIN'])/2
           BD_PULPAF['ZnProm'] = (BD_PULPAF['ZN_MAX']+BD_PULPAF['ZN_MIN'])/2
           BD_PULPAF['FeProm'] = (BD_PULPAF['FE_MAX']+BD_PULPAF['FE_MIN'])/2
           BD_PULPAF['CuRD'] = 200*(BD_PULPAF['CU_O']-BD_PULPAF['CU'])/(BD_PULPAF['CU_O']+BD_PULPAF['CU'])
           BD_PULPAF['AgRD'] = 200*(BD_PULPAF['AG_O']-BD_PULPAF['AG'])/(BD_PULPAF['AG_O']+BD_PULPAF['AG'])
           BD_PULPAF['MoRD'] = 200*(BD_PULPAF['MO_O']-BD_PULPAF['MO'])/(BD_PULPAF['MO_O']+BD_PULPAF['MO'])
           BD_PULPAF['AsRD'] = 200*(BD_PULPAF['AS_O']-BD_PULPAF['AS'])/(BD_PULPAF['AS_O']+BD_PULPAF['AS'])
           BD_PULPAF['ZnRD'] = 200*(BD_PULPAF['ZN_O']-BD_PULPAF['ZN'])/(BD_PULPAF['ZN_O']+BD_PULPAF['ZN'])
           BD_PULPAF['FeRD'] = 200*(BD_PULPAF['FE_O']-BD_PULPAF['FE'])/(BD_PULPAF['FE_O']+BD_PULPAF['FE'])
           BD_PULPAF = BD_PULPAF.sort_values('CuProm')
           BD_PULPAF['MediaMovCuProm'] = BD_PULPAF['CuProm'].rolling(window=2).mean()
           BD_PULPAF = BD_PULPAF.sort_values('AgProm')
           BD_PULPAF['MediaMovAgProm'] = BD_PULPAF['AgProm'].rolling(window=2).mean()
           BD_PULPAF = BD_PULPAF.sort_values('MoProm')
           BD_PULPAF['MediaMovMoProm'] = BD_PULPAF['MoProm'].rolling(window=2).mean()
           BD_PULPAF = BD_PULPAF.sort_values('AsProm')
           BD_PULPAF['MediaMovAsProm'] = BD_PULPAF['AsProm'].rolling(window=2).mean()
           BD_PULPAF = BD_PULPAF.sort_values('ZnProm')
           BD_PULPAF['MediaMovZnProm'] = BD_PULPAF['ZnProm'].rolling(window=2).mean()
           BD_PULPAF = BD_PULPAF.sort_values('FeProm')
           BD_PULPAF['MediaMovFeProm'] = BD_PULPAF['FeProm'].rolling(window=2).mean()
           BD_PULPAF = BD_PULPAF.sort_values('SAMPLE_ID')       
           self.BD_PULPAF = BD_PULPAF 
           # Data STD13
           BD_STD13F = pd.merge(self.BD_STD13, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','LAB_REPORT','SAMPLE_ID', 'MES ']], on='SAMPLE_ID', how='left')
           BD_STD13F = BD_STD13F.rename(columns={'ARS':'AS'})
           BD_STD13F['AVGCU']= BD_STD13F['CU'].mean()
           BD_STD13F['AVGAG']= BD_STD13F['AG'].mean()
           BD_STD13F['AVGMO']= BD_STD13F['MO'].mean()
           BD_STD13F['AVGAS']= BD_STD13F['AS'].mean()
           BD_STD13F['AVGZN']= BD_STD13F['ZN'].mean()
           MHCU13 = 0.266184914841849
           MHAG13 = 2.6863779342723
           MHMO13 = 0.00419716312056736
           MHAS13 = 0.0072113022113022
           MHZN13 = 0.0253397590361446
           BVCU13 = 0.266031884057971
           BVMO13 = 0.00423220338983051
           BVZN13 = 0.0263231884057971       
           BVCU13 = 0.266031884057971
           BVMO13 = 0.00423220338983051
           BVZN13 = 0.0263231884057971
           BD_STD13F['MHCU'] = MHCU13
           BD_STD13F['MHAG'] = MHAG13
           BD_STD13F['MHMO'] = MHMO13
           BD_STD13F['MHAS'] = MHAS13
           BD_STD13F['MHZN'] = MHZN13
           STDCU = 0.0036884613991108
           STDAG = 0.184056641873746
           STDMO = 0.000121177044741277
           STDAS = 0.000701560557845971
           STDZN = 0.000902350628509791
           BD_STD13F['-2STDCU'] =  BD_STD13F['AVGCU'] - 2*STDCU
           BD_STD13F['+2STDCU'] =  BD_STD13F['AVGCU'] + 2*STDCU
           BD_STD13F['-2STDAG'] =  BD_STD13F['AVGAG'] - 2*STDAG
           BD_STD13F['+2STDAG'] =  BD_STD13F['AVGAG'] + 2*STDAG
           BD_STD13F['-2STDMO'] =  BD_STD13F['AVGMO'] - 2*STDMO
           BD_STD13F['+2STDMO'] =  BD_STD13F['AVGMO'] + 2*STDMO
           BD_STD13F['-2STDAS'] =  BD_STD13F['AVGAS'] - 2*STDAS
           BD_STD13F['+2STDAS'] =  BD_STD13F['AVGAS'] + 2*STDAS
           BD_STD13F['-2STDZN'] =  BD_STD13F['AVGZN'] - 2*STDZN
           BD_STD13F['+2STDZN'] =  BD_STD13F['AVGZN'] + 2*STDZN
           BD_STD13F['-2STDCU'] =  MHCU13 - 2*STDCU
           BD_STD13F['+2STDCU'] =  MHCU13 + 2*STDCU
           BD_STD13F['-2STDAG'] =  MHAG13 - 2*STDAG
           BD_STD13F['+2STDAG'] =  MHAG13 + 2*STDAG
           BD_STD13F['-2STDMO'] =  MHMO13 - 2*STDMO
           BD_STD13F['+2STDMO'] =  MHMO13 + 2*STDMO
           BD_STD13F['-2STDAS'] =  MHAS13 - 2*STDAS
           BD_STD13F['+2STDAS'] =  MHAS13 + 2*STDAS
           BD_STD13F['-2STDZN'] =  MHZN13 - 2*STDZN
           BD_STD13F['+2STDZN'] =  MHZN13 + 2*STDZN
           BD_STD13F['-3STDCU'] =  MHCU13 - 3*STDCU
           BD_STD13F['+3STDCU'] =  MHCU13 + 3*STDCU
           BD_STD13F['-3STDAG'] =  MHAG13 - 3*STDAG
           BD_STD13F['+3STDAG'] =  MHAG13 + 3*STDAG
           BD_STD13F['-3STDMO'] =  MHMO13 - 3*STDMO
           BD_STD13F['+3STDMO'] =  MHMO13 + 3*STDMO
           BD_STD13F['-3STDAS'] =  MHAS13 - 3*STDAS
           BD_STD13F['+3STDAS'] =  MHAS13 + 3*STDAS
           BD_STD13F['-3STDZN'] =  MHZN13 - 3*STDZN
           BD_STD13F['+3STDZN'] =  MHZN13 + 3*STDZN
           BD_STD13F['BVCU'] = BVCU13
           BD_STD13F['BVMO'] = BVMO13
           BD_STD13F['BVZN'] = BVZN13
           BD_STD13F['Cu_Out_Pos'] = (BD_STD13F['CU'] > BD_STD13F['+3STDCU']).astype(int)
           BD_STD13F['Ag_Out_Pos'] = (BD_STD13F['AG'] > BD_STD13F['+3STDAG']).astype(int)
           BD_STD13F['Mo_Out_Pos'] = (BD_STD13F['MO'] > BD_STD13F['+3STDMO']).astype(int)
           BD_STD13F['As_Out_Pos'] = (BD_STD13F['AS'] > BD_STD13F['+3STDAS']).astype(int)
           BD_STD13F['Zn_Out_Pos'] = (BD_STD13F['ZN'] > BD_STD13F['+3STDZN']).astype(int)
           BD_STD13F['Cu_Out_Neg'] = (BD_STD13F['CU'] < BD_STD13F['-3STDCU']).astype(int)
           BD_STD13F['Ag_Out_Neg'] = (BD_STD13F['AG'] < BD_STD13F['-3STDAG']).astype(int)
           BD_STD13F['Mo_Out_Neg'] = (BD_STD13F['MO'] < BD_STD13F['-3STDMO']).astype(int)
           BD_STD13F['As_Out_Neg'] = (BD_STD13F['AS'] < BD_STD13F['-3STDAS']).astype(int)
           BD_STD13F['Zn_Out_Neg'] = (BD_STD13F['ZN'] < BD_STD13F['-3STDZN']).astype(int)
           BD_STD13F['Cu_Out'] = BD_STD13F['Cu_Out_Pos'] + BD_STD13F['Cu_Out_Neg']
           BD_STD13F['Ag_Out'] = BD_STD13F['Ag_Out_Pos'] + BD_STD13F['Ag_Out_Neg']
           BD_STD13F['Mo_Out'] = BD_STD13F['Mo_Out_Pos'] + BD_STD13F['Mo_Out_Neg']
           BD_STD13F['As_Out'] = BD_STD13F['As_Out_Pos'] + BD_STD13F['As_Out_Neg']
           BD_STD13F['Zn_Out'] = BD_STD13F['Zn_Out_Pos'] + BD_STD13F['Zn_Out_Neg']
           self.FSTD13CU = BD_STD13F['Cu_Out'].sum()*100/len(BD_STD13F['CU'])
           self.FSTD13AG = BD_STD13F['Ag_Out'].sum()*100/len(BD_STD13F['AG'])
           self.FSTD13MO = BD_STD13F['Mo_Out'].sum()*100/len(BD_STD13F['MO'])
           self.FSTD13AS = BD_STD13F['As_Out'].sum()*100/len(BD_STD13F['AS'])
           self.FSTD13ZN = BD_STD13F['Zn_Out'].sum()*100/len(BD_STD13F['ZN'])
           self.STDB_ERROR_CU = ((BD_STD13F['CU'].mean()/BVCU13)-1)*100
           self.STDB_ERROR_MO = ((BD_STD13F['MO'].mean()/BVMO13)-1)*100
           self.STDB_ERROR_ZN = ((BD_STD13F['ZN'].mean()/BVZN13)-1)*100
           self.BD_STD13F = BD_STD13F
           #Data STD 14
           BD_STD14F = pd.merge(self.BD_STD14, self.BD_Total[['CU', 'AG', 'MO', 'ARS','LAB_REPORT', 'ZN', 'FE','SAMPLE_ID','MES ']], on='SAMPLE_ID', how='left')
           BD_STD14F = BD_STD14F.rename(columns={'ARS':'AS'})
           BD_STD14F['AVGCU']= BD_STD14F['CU'].mean()
           BD_STD14F['AVGAG']= BD_STD14F['AG'].mean()
           BD_STD14F['AVGMO']= BD_STD14F['MO'].mean()
           BD_STD14F['AVGAS']= BD_STD14F['AS'].mean()
           BD_STD14F['AVGZN']= BD_STD14F['ZN'].mean()
           MHCU14 = 0.427990543735225
           MHAG14 = 4.6302027972028
           MHMO14 = 0.00978690476190473
           MHAS14 = 0.016616091954023
           MHZN14 = 0.0456080760095013
           BVCU14 = 0.420332352941176
           BVMO14 = 0.00996842105263158
           BVAS14 = 0.0162754385964912
           BVZN14 = 0.0453863636363636
           BD_STD14F['MHCU'] = MHCU14
           BD_STD14F['MHAG'] = MHAG14
           BD_STD14F['MHMO'] = MHMO14
           BD_STD14F['MHAS'] = MHAS14
           BD_STD14F['MHZN'] = MHZN14
           STDCU14 = 0.00745666186049466
           STDAG14 = 0.177246972191448
           STDMO14 = 0.000271115331782672
           STDAS14 = 0.000827086603439742
           STDZN14 = 0.00114052258444863
           BD_STD14F['-2STDCU'] =  BD_STD14F['AVGCU'] - 2*STDCU14
           BD_STD14F['+2STDCU'] =  BD_STD14F['AVGCU'] + 2*STDCU14
           BD_STD14F['-2STDAG'] =  BD_STD14F['AVGAG'] - 2*STDAG14
           BD_STD14F['+2STDAG'] =  BD_STD14F['AVGAG'] + 2*STDAG14
           BD_STD14F['-2STDMO'] =  BD_STD14F['AVGMO'] - 2*STDMO14
           BD_STD14F['+2STDMO'] =  BD_STD14F['AVGMO'] + 2*STDMO14
           BD_STD14F['-2STDAS'] =  BD_STD14F['AVGAS'] - 2*STDAS14
           BD_STD14F['+2STDAS'] =  BD_STD14F['AVGAS'] + 2*STDAS14
           BD_STD14F['-2STDZN'] =  BD_STD14F['AVGZN'] - 2*STDZN14
           BD_STD14F['+2STDZN'] =  BD_STD14F['AVGZN'] + 2*STDZN14
           BD_STD14F['-2STDCU'] =  MHCU14 - 2*STDCU14
           BD_STD14F['+2STDCU'] =  MHCU14 + 2*STDCU14
           BD_STD14F['-2STDAG'] =  MHAG14 - 2*STDAG14
           BD_STD14F['+2STDAG'] =  MHAG14 + 2*STDAG14
           BD_STD14F['-2STDMO'] =  MHMO14 - 2*STDMO14
           BD_STD14F['+2STDMO'] =  MHMO14 + 2*STDMO14
           BD_STD14F['-2STDAS'] =  MHAS14 - 2*STDAS14
           BD_STD14F['+2STDAS'] =  MHAS14 + 2*STDAS14
           BD_STD14F['-2STDZN'] =  MHZN14 - 2*STDZN14
           BD_STD14F['+2STDZN'] =  MHZN14 + 2*STDZN14
           BD_STD14F['-3STDCU'] =  MHCU14 - 3*STDCU14
           BD_STD14F['+3STDCU'] =  MHCU14 + 3*STDCU14
           BD_STD14F['-3STDAG'] =  MHAG14 - 3*STDAG14
           BD_STD14F['+3STDAG'] =  MHAG14 + 3*STDAG14
           BD_STD14F['-3STDMO'] =  MHMO14 - 3*STDMO14
           BD_STD14F['+3STDMO'] =  MHMO14 + 3*STDMO14
           BD_STD14F['-3STDAS'] =  MHAS14 - 3*STDAS14
           BD_STD14F['+3STDAS'] =  MHAS14 + 3*STDAS14
           BD_STD14F['-3STDZN'] =  MHZN14 - 3*STDZN14
           BD_STD14F['+3STDZN'] =  MHZN14 + 3*STDZN14     
           BD_STD14F['BVCU'] = BVCU14
           BD_STD14F['BVMO'] = BVMO14
           BD_STD14F['BVAS'] = BVAS14
           BD_STD14F['BVZN'] = BVZN14 
           BD_STD14F['Cu_Out_Pos'] = (BD_STD14F['CU'] > BD_STD14F['+3STDCU']).astype(int)
           BD_STD14F['Ag_Out_Pos'] = (BD_STD14F['AG'] > BD_STD14F['+3STDAG']).astype(int)
           BD_STD14F['Mo_Out_Pos'] = (BD_STD14F['MO'] > BD_STD14F['+3STDMO']).astype(int)
           BD_STD14F['As_Out_Pos'] = (BD_STD14F['AS'] > BD_STD14F['+3STDAS']).astype(int)
           BD_STD14F['Zn_Out_Pos'] = (BD_STD14F['ZN'] > BD_STD14F['+3STDZN']).astype(int)
           BD_STD14F['Cu_Out_Neg'] = (BD_STD14F['CU'] < BD_STD14F['-3STDCU']).astype(int)
           BD_STD14F['Ag_Out_Neg'] = (BD_STD14F['AG'] < BD_STD14F['-3STDAG']).astype(int)
           BD_STD14F['Mo_Out_Neg'] = (BD_STD14F['MO'] < BD_STD14F['-3STDMO']).astype(int)
           BD_STD14F['As_Out_Neg'] = (BD_STD14F['AS'] < BD_STD14F['-3STDAS']).astype(int)
           BD_STD14F['Zn_Out_Neg'] = (BD_STD14F['ZN'] < BD_STD14F['-3STDZN']).astype(int)
           BD_STD14F['Cu_Out'] = BD_STD14F['Cu_Out_Pos'] + BD_STD14F['Cu_Out_Neg']
           BD_STD14F['Ag_Out'] = BD_STD14F['Ag_Out_Pos'] + BD_STD14F['Ag_Out_Neg']
           BD_STD14F['Mo_Out'] = BD_STD14F['Mo_Out_Pos'] + BD_STD14F['Mo_Out_Neg']
           BD_STD14F['As_Out'] = BD_STD14F['As_Out_Pos'] + BD_STD14F['As_Out_Neg']
           BD_STD14F['Zn_Out'] = BD_STD14F['Zn_Out_Pos'] + BD_STD14F['Zn_Out_Neg']
           self.STDM_ERROR_CU = ((BD_STD14F['CU'].mean()/BVCU14)-1)*100
           self.STDM_ERROR_MO = ((BD_STD14F['MO'].mean()/BVMO14)-1)*100
           self.STDM_ERROR_AS = ((BD_STD14F['AS'].mean()/BVAS14)-1)*100
           self.STDM_ERROR_ZN = ((BD_STD14F['ZN'].mean()/BVZN14)-1)*100
           self.FSTD14CU = BD_STD14F['Cu_Out'].sum()*100/len(BD_STD14F['CU'])
           self.FSTD14AG = BD_STD14F['Ag_Out'].sum()*100/len(BD_STD14F['AG'])
           self.FSTD14MO = BD_STD14F['Mo_Out'].sum()*100/len(BD_STD14F['MO'])
           self.FSTD14AS = BD_STD14F['As_Out'].sum()*100/len(BD_STD14F['AS'])
           self.FSTD14ZN = BD_STD14F['Zn_Out'].sum()*100/len(BD_STD14F['ZN'])
           self.BD_STD14F = BD_STD14F
           #Data STD 15
           BD_STD15F = pd.merge(self.BD_STD15, self.BD_Total[['CU', 'AG', 'MO', 'ARS','LAB_REPORT', 'ZN', 'FE','SAMPLE_ID', 'MES ']], on='SAMPLE_ID', how='left')
           BD_STD15F = BD_STD15F.rename(columns={'ARS':'AS'})
           BD_STD15F['AVGCU']= BD_STD15F['CU'].mean()
           BD_STD15F['AVGAG']= BD_STD15F['AG'].mean()
           BD_STD15F['AVGMO']= BD_STD15F['MO'].mean()
           BD_STD15F['AVGAS']= BD_STD15F['AS'].mean()
           BD_STD15F['AVGZN']= BD_STD15F['ZN'].mean()
           MHCU15 = 0.662866666666667
           MHAG15 = 10.3921145584725
           MHMO15 = 0.0104511961722488
           MHAS15 = 0.0685093896713616
           MHZN15 = 0.155243309002433
           BVCU15 = 0.676945762711864
           BVAG15 = 10.140350877193
           BVMO15 = 0.0105706896551724
           BVAS15 = 0.07035
           BVZN15 = 0.149810294117647
           BD_STD15F['MHCU'] = MHCU15
           BD_STD15F['MHAG'] = MHAG15
           BD_STD15F['MHMO'] = MHMO15
           BD_STD15F['MHAS'] = MHAS15
           BD_STD15F['MHZN'] = MHZN15
           STDCU15 = 0.0119090159918189
           STDAG15 = 0.181529114416518
           STDMO15 = 0.000309809262563219
           STDAS15 = 0.00221782216510573
           STDZN15 = 0.00304505560834989
           BD_STD15F['-2STDCU'] =  BD_STD15F['AVGCU'] - 2*STDCU15
           BD_STD15F['+2STDCU'] =  BD_STD15F['AVGCU'] + 2*STDCU15
           BD_STD15F['-2STDAG'] =  BD_STD15F['AVGAG'] - 2*STDAG15
           BD_STD15F['+2STDAG'] =  BD_STD15F['AVGAG'] + 2*STDAG15
           BD_STD15F['-2STDMO'] =  BD_STD15F['AVGMO'] - 2*STDMO15
           BD_STD15F['+2STDMO'] =  BD_STD15F['AVGMO'] + 2*STDMO15
           BD_STD15F['-2STDAS'] =  BD_STD15F['AVGAS'] - 2*STDAS15
           BD_STD15F['+2STDAS'] =  BD_STD15F['AVGAS'] + 2*STDAS15
           BD_STD15F['-2STDZN'] =  BD_STD15F['AVGZN'] - 2*STDZN15
           BD_STD15F['+2STDZN'] =  BD_STD15F['AVGZN'] + 2*STDZN15
           BD_STD15F['-2STDCU'] =  MHCU15 - 2*STDCU15
           BD_STD15F['+2STDCU'] =  MHCU15 + 2*STDCU15
           BD_STD15F['-2STDAG'] =  MHAG15 - 2*STDAG15
           BD_STD15F['+2STDAG'] =  MHAG15 + 2*STDAG15
           BD_STD15F['-2STDMO'] =  MHMO15 - 2*STDMO15
           BD_STD15F['+2STDMO'] =  MHMO15 + 2*STDMO15
           BD_STD15F['-2STDAS'] =  MHAS15 - 2*STDAS15
           BD_STD15F['+2STDAS'] =  MHAS15 + 2*STDAS15
           BD_STD15F['-2STDZN'] =  MHZN15 - 2*STDZN15
           BD_STD15F['+2STDZN'] =  MHZN15 + 2*STDZN15
           BD_STD15F['-3STDCU'] =  MHCU15 - 3*STDCU15
           BD_STD15F['+3STDCU'] =  MHCU15 + 3*STDCU15
           BD_STD15F['-3STDAG'] =  MHAG15 - 3*STDAG15
           BD_STD15F['+3STDAG'] =  MHAG15 + 3*STDAG15
           BD_STD15F['-3STDMO'] =  MHMO15 - 3*STDMO15
           BD_STD15F['+3STDMO'] =  MHMO15 + 3*STDMO15
           BD_STD15F['-3STDAS'] =  MHAS15 - 3*STDAS15
           BD_STD15F['+3STDAS'] =  MHAS15 + 3*STDAS15
           BD_STD15F['-3STDZN'] =  MHZN15 - 3*STDZN15
           BD_STD15F['+3STDZN'] =  MHZN15 + 3*STDZN15
           BD_STD15F['BVCU'] = BVCU15
           BD_STD15F['BVAG'] = BVAG15
           BD_STD15F['BVMO'] = BVMO15
           BD_STD15F['BVAS'] = BVAS15
           BD_STD15F['BVZN'] = BVZN15
           BD_STD15F['Cu_Out_Pos'] = (BD_STD15F['CU'] > BD_STD15F['+3STDCU']).astype(int)
           BD_STD15F['Ag_Out_Pos'] = (BD_STD15F['AG'] > BD_STD15F['+3STDAG']).astype(int)
           BD_STD15F['Mo_Out_Pos'] = (BD_STD15F['MO'] > BD_STD15F['+3STDMO']).astype(int)
           BD_STD15F['As_Out_Pos'] = (BD_STD15F['AS'] > BD_STD15F['+3STDAS']).astype(int)
           BD_STD15F['Zn_Out_Pos'] = (BD_STD15F['ZN'] > BD_STD15F['+3STDZN']).astype(int)
           BD_STD15F['Cu_Out_Neg'] = (BD_STD15F['CU'] < BD_STD15F['-3STDCU']).astype(int)
           BD_STD15F['Ag_Out_Neg'] = (BD_STD15F['AG'] < BD_STD15F['-3STDAG']).astype(int)
           BD_STD15F['Mo_Out_Neg'] = (BD_STD15F['MO'] < BD_STD15F['-3STDMO']).astype(int)
           BD_STD15F['As_Out_Neg'] = (BD_STD15F['AS'] < BD_STD15F['-3STDAS']).astype(int)
           BD_STD15F['Zn_Out_Neg'] = (BD_STD15F['ZN'] < BD_STD15F['-3STDZN']).astype(int)
           BD_STD15F['Cu_Out'] = BD_STD15F['Cu_Out_Pos'] + BD_STD15F['Cu_Out_Neg']
           BD_STD15F['Ag_Out'] = BD_STD15F['Ag_Out_Pos'] + BD_STD15F['Ag_Out_Neg']
           BD_STD15F['Mo_Out'] = BD_STD15F['Mo_Out_Pos'] + BD_STD15F['Mo_Out_Neg']
           BD_STD15F['As_Out'] = BD_STD15F['As_Out_Pos'] + BD_STD15F['As_Out_Neg']
           BD_STD15F['Zn_Out'] = BD_STD15F['Zn_Out_Pos'] + BD_STD15F['Zn_Out_Neg']
           self.STDA_ERROR_CU = ((BD_STD15F['CU'].mean()/BVCU15)-1)*100
           self.STDA_ERROR_AG = ((BD_STD15F['AG'].mean()/BVAG15)-1)*100
           self.STDA_ERROR_MO = ((BD_STD15F['MO'].mean()/BVMO15)-1)*100
           self.STDA_ERROR_AS = ((BD_STD15F['AS'].mean()/BVAS15)-1)*100
           self.STDA_ERROR_ZN = ((BD_STD15F['ZN'].mean()/BVZN15)-1)*100      
           self.FSTD15CU = BD_STD15F['Cu_Out'].sum()*100/len(BD_STD15F['CU'])
           self.FSTD15AG = BD_STD15F['Ag_Out'].sum()*100/len(BD_STD15F['AG'])
           self.FSTD15MO = BD_STD15F['Mo_Out'].sum()*100/len(BD_STD15F['MO'])
           self.FSTD15AS = BD_STD15F['As_Out'].sum()*100/len(BD_STD15F['AS'])
           self.FSTD15ZN = BD_STD15F['Zn_Out'].sum()*100/len(BD_STD15F['ZN'])       
           self.BD_STD15F = BD_STD15F
           # Data Blanco Grueso
           self.BD_BLANCO['SAMPLE_PREC'] = self.BD_BLANCO['SAMPLE_ID']-2
           BD_BLANCOF = pd.merge(self.BD_BLANCO,self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','LAB_REPORT','SAMPLE_PREC', 'MES ']], on='SAMPLE_PREC', how='left')
           Elem_Blanco = {'CU':'CU_P', 'AG':'AG_P', 'MO':'MO_P', 'ARS':'AS_P', 'ZN':'ZN_P', 'FE':'FE_P'}
           BD_BLANCOF = BD_BLANCOF.rename(columns=Elem_Blanco)
           BD_BLANCOF = pd.merge(BD_BLANCOF, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','SAMPLE_ID']], on = 'SAMPLE_ID', how='left')       
           BD_BLANCOF = BD_BLANCOF.rename(columns={'ARS':'AS'})    
           self.LDCu= 0.001
           self.LDAg= 0.3
           self.LDMo= 0.0015
           self.LDAs= 0.002
           self.LDZnf= 0.0016667
           self.LDZng= 0.001
           self.LDFe_fin= 0.495
           self.LDFe_gr= 0.127       
           BD_BLANCOF['Cu_MAV'] = 5*self.LDCu
           BD_BLANCOF['Ag_MAV'] = 5*self.LDAg
           BD_BLANCOF['Mo_MAV'] = 5*self.LDMo
           BD_BLANCOF['As_MAV'] = 5*self.LDAs
           BD_BLANCOF['Zn_MAV'] = 5*self.LDZng
           BD_BLANCOF['Fe_MAV'] = 5*self.LDFe_gr
           BD_BLANCOF['Cu_Error'] = (BD_BLANCOF['CU'] > BD_BLANCOF['Cu_MAV']).astype(int)
           BD_BLANCOF['Ag_Error'] = (BD_BLANCOF['AG'] > BD_BLANCOF['Ag_MAV']).astype(int)
           BD_BLANCOF['Mo_Error'] = (BD_BLANCOF['MO'] > BD_BLANCOF['Mo_MAV']).astype(int)
           BD_BLANCOF['As_Error'] = (BD_BLANCOF['AS'] > BD_BLANCOF['As_MAV']).astype(int)
           BD_BLANCOF['Zn_Error'] = (BD_BLANCOF['ZN'] > BD_BLANCOF['Zn_MAV']).astype(int)
           BD_BLANCOF['Fe_Error'] = (BD_BLANCOF['FE'] > BD_BLANCOF['Fe_MAV']).astype(int)
           self.CuErrorBLKG  = BD_BLANCOF['Cu_Error'].sum()*100/len(BD_BLANCOF['CU'])
           self.AgErrorBLKG  = BD_BLANCOF['Ag_Error'].sum()*100/len(BD_BLANCOF['AG'])
           self.MoErrorBLKG  = BD_BLANCOF['Mo_Error'].sum()*100/len(BD_BLANCOF['MO'])
           self.AsErrorBLKG  = BD_BLANCOF['As_Error'].sum()*100/len(BD_BLANCOF['AS'])
           self.ZnErrorBLKG  = BD_BLANCOF['Zn_Error'].sum()*100/len(BD_BLANCOF['ZN'])
           self.FeErrorBLKG  = BD_BLANCOF['Fe_Error'].sum()*100/len(BD_BLANCOF['FE'])
           self.BD_BLANCOF  = BD_BLANCOF
           #Data Blanco Fino
           self.BD_BLKFINO['SAMPLE_PREC'] = self.BD_BLKFINO['SAMPLE_ID']-1
           BD_BLKFINOF = pd.merge(self.BD_BLKFINO, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','LAB_REPORT','SAMPLE_PREC', 'MES ']], on= 'SAMPLE_PREC', how='left')
           BD_BLKFINOF = BD_BLKFINOF.rename(columns=Elem_Blanco)
           BD_BLKFINOF = pd.merge(BD_BLKFINOF, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','SAMPLE_ID']], on='SAMPLE_ID', how='left')
           BD_BLKFINOF = BD_BLKFINOF.rename(columns={'ARS':'AS'})
           BD_BLKFINOF['Cu_MAV'] = 3*self.LDCu
           BD_BLKFINOF['Ag_MAV'] = 3*self.LDAg
           BD_BLKFINOF['Mo_MAV'] = 3*self.LDMo
           BD_BLKFINOF['As_MAV'] = 3*self.LDAs
           BD_BLKFINOF['Zn_MAV'] = 3*self.LDZnf
           BD_BLKFINOF['FeFin_MAV'] = 3*self.LDFe_fin
           BD_BLKFINOF['Cu_Error'] = (BD_BLKFINOF['CU'] > BD_BLKFINOF['Cu_MAV']).astype(int)
           BD_BLKFINOF['Ag_Error'] = (BD_BLKFINOF['AG'] > BD_BLKFINOF['Ag_MAV']).astype(int)
           BD_BLKFINOF['Mo_Error'] = (BD_BLKFINOF['MO'] > BD_BLKFINOF['Mo_MAV']).astype(int)
           BD_BLKFINOF['As_Error'] = (BD_BLKFINOF['AS'] > BD_BLKFINOF['As_MAV']).astype(int)
           BD_BLKFINOF['Zn_Error'] = (BD_BLKFINOF['ZN'] > BD_BLKFINOF['Zn_MAV']).astype(int)
           BD_BLKFINOF['Fe_Error'] = (BD_BLKFINOF['FE'] > BD_BLKFINOF['FeFin_MAV']).astype(int)      
           self.CuErrorBLKF  = BD_BLKFINOF['Cu_Error'].sum()*100/len(BD_BLKFINOF['CU'])
           self.AgErrorBLKF  = BD_BLKFINOF['Ag_Error'].sum()*100/len(BD_BLKFINOF['AG'])
           self.MoErrorBLKF  = BD_BLKFINOF['Mo_Error'].sum()*100/len(BD_BLKFINOF['MO'])
           self.AsErrorBLKF  = BD_BLKFINOF['As_Error'].sum()*100/len(BD_BLKFINOF['AS'])
           self.ZnErrorBLKF  = BD_BLKFINOF['Zn_Error'].sum()*100/len(BD_BLKFINOF['ZN'])
           self.FeErrorBLKF  = BD_BLKFINOF['Fe_Error'].sum()*100/len(BD_BLKFINOF['FE'])     
           self.BD_BLKFINOF = BD_BLKFINOF
           # Duplicado de Campo
           self.BD_Orig['SAMPLE_DUPLICADO'] = self.BD_Orig['SAMPLE_ID'] + 1
           BD_OrigF = pd.merge(self.BD_Orig, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','LAB_REPORT','SAMPLE_DUPLICADO']], on='SAMPLE_DUPLICADO', how='left')
           Elem_Dup = {'CU':'CU_D', 'AG':'AG_D', 'MO':'MO_D', 'ARS':'AS_D', 'ZN':'ZN_D', 'FE':'FE_D'}
           BD_OrigF = BD_OrigF.rename(columns=Elem_Dup)
           BD_OrigF = pd.merge(BD_OrigF, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','SAMPLE_ID']], on='SAMPLE_ID', how='left')
           BD_OrigF = BD_OrigF.rename(columns={'ARS':'AS'})
           FACTORB_CU = LDPCU*10
           FACTORB_AG = LDPAG*10
           FACTORB_MO = LDPMO*10
           FACTORB_AS = LDPAS*10
           FACTORB_ZN = LDPZN*10
           FACTORB_FE = LDPFE*10
           PENDIENTE_CAMPO = 1.35
           BD_OrigF['CU_MIN'] = BD_OrigF[['CU_D', 'CU']].min(axis=1)
           BD_OrigF['CU_MAX'] = BD_OrigF[['CU_D', 'CU']].max(axis=1)
           BD_OrigF['AG_MIN'] = BD_OrigF[['AG_D', 'AG']].min(axis=1)
           BD_OrigF['AG_MAX'] = BD_OrigF[['AG_D', 'AG']].max(axis=1)
           BD_OrigF['MO_MIN'] = BD_OrigF[['MO_D', 'MO']].min(axis=1)
           BD_OrigF['MO_MAX'] = BD_OrigF[['MO_D', 'MO']].max(axis=1)
           BD_OrigF['AS_MIN'] = BD_OrigF[['AS_D', 'AS']].min(axis=1)
           BD_OrigF['AS_MAX'] = BD_OrigF[['AS_D', 'AS']].max(axis=1)
           BD_OrigF['ZN_MIN'] = BD_OrigF[['ZN_D', 'ZN']].min(axis=1)
           BD_OrigF['ZN_MAX'] = BD_OrigF[['ZN_D', 'ZN']].max(axis=1)
           BD_OrigF['FE_MIN'] = BD_OrigF[['FE_D', 'FE']].min(axis=1)
           BD_OrigF['FE_MAX'] = BD_OrigF[['FE_D', 'FE']].max(axis=1)       
           BD_OrigF['Ec_CU'] = np.sqrt((PENDIENTE_CAMPO**2)*(BD_OrigF['CU_MIN']**2)+(FACTORB_CU**2))
           BD_OrigF['Ec_AG'] = np.sqrt((PENDIENTE_CAMPO**2)*(BD_OrigF['AG_MIN']**2)+(FACTORB_AG**2))
           BD_OrigF['Ec_MO'] = np.sqrt((PENDIENTE_CAMPO**2)*(BD_OrigF['MO_MIN']**2)+(FACTORB_MO**2))
           BD_OrigF['Ec_AS'] = np.sqrt((PENDIENTE_CAMPO**2)*(BD_OrigF['AS_MIN']**2)+(FACTORB_AS**2))
           BD_OrigF['Ec_ZN'] = np.sqrt((PENDIENTE_CAMPO**2)*(BD_OrigF['ZN_MIN']**2)+(FACTORB_ZN**2))
           BD_OrigF['Ec_FE'] = np.sqrt((PENDIENTE_CAMPO**2)*(BD_OrigF['FE_MIN']**2)+(FACTORB_FE**2))
           BD_OrigF['Cu_Error'] = (BD_OrigF['CU_MAX'] > BD_OrigF['Ec_CU']).astype(int)
           BD_OrigF['Ag_Error'] = (BD_OrigF['AG_MAX'] > BD_OrigF['Ec_AG']).astype(int)
           BD_OrigF['Mo_Error'] = (BD_OrigF['MO_MAX'] > BD_OrigF['Ec_MO']).astype(int)
           BD_OrigF['As_Error'] = (BD_OrigF['AS_MAX'] > BD_OrigF['Ec_AS']).astype(int)
           BD_OrigF['Zn_Error'] = (BD_OrigF['ZN_MAX'] > BD_OrigF['Ec_ZN']).astype(int)
           BD_OrigF['Fe_Error'] = (BD_OrigF['FE_MAX'] > BD_OrigF['Ec_FE']).astype(int)
           BD_OrigF['Orden'] = BD_OrigF.index
           ValorMaxCu = BD_OrigF['CU_MAX'].max()
           ValorMaxAg = BD_OrigF['AG_MAX'].max()
           ValorMaxMo = BD_OrigF['MO_MAX'].max()
           ValorMaxAs = BD_OrigF['AS_MAX'].max()
           ValorMaxZn = BD_OrigF['ZN_MAX'].max()
           ValorMaxFe = BD_OrigF['FE_MAX'].max()
           BD_OrigF['XLimCU'] = ValorMaxCu*BD_OrigF['Orden']/len(BD_OrigF['CU_MAX'])
           BD_OrigF['XLimAG'] = ValorMaxAg*BD_OrigF['Orden']/len(BD_OrigF['AG_MAX'])
           BD_OrigF['XLimMO'] = ValorMaxMo*BD_OrigF['Orden']/len(BD_OrigF['MO_MAX'])
           BD_OrigF['XLimAS'] = ValorMaxAs*BD_OrigF['Orden']/len(BD_OrigF['AS_MAX'])
           BD_OrigF['XLimZN'] = ValorMaxZn*BD_OrigF['Orden']/len(BD_OrigF['ZN_MAX'])
           BD_OrigF['XLimFE'] = ValorMaxFe*BD_OrigF['Orden']/len(BD_OrigF['FE_MAX'])
           BD_OrigF['YLimCU'] = np.sqrt(PENDIENTE_CAMPO**2*BD_OrigF['XLimCU']**2+FACTORB_CU**2)
           BD_OrigF['YLimAG'] = np.sqrt(PENDIENTE_CAMPO**2*BD_OrigF['XLimAG']**2+FACTORB_AG**2)
           BD_OrigF['YLimMO'] = np.sqrt(PENDIENTE_CAMPO**2*BD_OrigF['XLimMO']**2+FACTORB_MO**2)
           BD_OrigF['YLimAS'] = np.sqrt(PENDIENTE_CAMPO**2*BD_OrigF['XLimAS']**2+FACTORB_AS**2)
           BD_OrigF['YLimZN'] = np.sqrt(PENDIENTE_CAMPO**2*BD_OrigF['XLimZN']**2+FACTORB_ZN**2)
           BD_OrigF['YLimFE'] = np.sqrt(PENDIENTE_CAMPO**2*BD_OrigF['XLimFE']**2+FACTORB_FE**2)
           self.CuErrorDUPC = BD_OrigF['Cu_Error'].sum()*100/len(BD_OrigF['CU_MAX'])
           self.AgErrorDUPC = BD_OrigF['Ag_Error'].sum()*100/len(BD_OrigF['AG_MAX'])
           self.MoErrorDUPC = BD_OrigF['Mo_Error'].sum()*100/len(BD_OrigF['MO_MAX'])
           self.AsErrorDUPC = BD_OrigF['As_Error'].sum()*100/len(BD_OrigF['AS_MAX'])
           self.ZnErrorDUPC = BD_OrigF['Zn_Error'].sum()*100/len(BD_OrigF['ZN_MAX'])
           self.FeErrorDUPC = BD_OrigF['Fe_Error'].sum()*100/len(BD_OrigF['FE_MAX'])
           BD_OrigF['CuProm'] = (BD_OrigF['CU_MAX']+BD_OrigF['CU_MIN'])/2
           BD_OrigF['AgProm'] = (BD_OrigF['AG_MAX']+BD_OrigF['AG_MIN'])/2
           BD_OrigF['MoProm'] = (BD_OrigF['MO_MAX']+BD_OrigF['MO_MIN'])/2
           BD_OrigF['AsProm'] = (BD_OrigF['AS_MAX']+BD_OrigF['AS_MIN'])/2
           BD_OrigF['ZnProm'] = (BD_OrigF['ZN_MAX']+BD_OrigF['ZN_MIN'])/2
           BD_OrigF['FeProm'] = (BD_OrigF['FE_MAX']+BD_OrigF['FE_MIN'])/2
           BD_OrigF['CuRD'] = 200*(BD_OrigF['CU']-BD_OrigF['CU_D'])/(BD_OrigF['CU']+BD_OrigF['CU_D'])
           BD_OrigF['AgRD'] = 200*(BD_OrigF['AG']-BD_OrigF['AG_D'])/(BD_OrigF['AG']+BD_OrigF['AG_D'])
           BD_OrigF['MoRD'] = 200*(BD_OrigF['MO']-BD_OrigF['MO_D'])/(BD_OrigF['MO']+BD_OrigF['MO_D'])
           BD_OrigF['AsRD'] = 200*(BD_OrigF['AS']-BD_OrigF['AS_D'])/(BD_OrigF['AS']+BD_OrigF['AS_D'])
           BD_OrigF['ZnRD'] = 200*(BD_OrigF['ZN']-BD_OrigF['ZN_D'])/(BD_OrigF['ZN']+BD_OrigF['ZN_D'])
           BD_OrigF['FeRD'] = 200*(BD_OrigF['FE']-BD_OrigF['FE_D'])/(BD_OrigF['FE']+BD_OrigF['FE_D'])
           BD_OrigF = BD_OrigF.sort_values('CuProm')
           BD_OrigF['MediaMovCuProm'] = BD_OrigF['CuProm'].rolling(window=2).mean()
           BD_OrigF = BD_OrigF.sort_values('AgProm')
           BD_OrigF['MediaMovAgProm'] = BD_OrigF['AgProm'].rolling(window=2).mean()
           BD_OrigF = BD_OrigF.sort_values('MoProm')
           BD_OrigF['MediaMovMoProm'] = BD_OrigF['MoProm'].rolling(window=2).mean()
           BD_OrigF = BD_OrigF.sort_values('AsProm')
           BD_OrigF['MediaMovAsProm'] = BD_OrigF['AsProm'].rolling(window=2).mean()
           BD_OrigF = BD_OrigF.sort_values('ZnProm')
           BD_OrigF['MediaMovZnProm'] = BD_OrigF['ZnProm'].rolling(window=2).mean()
           BD_OrigF = BD_OrigF.sort_values('FeProm')
           BD_OrigF['MediaMovFeProm'] = BD_OrigF['FeProm'].rolling(window=2).mean()
           BD_OrigF = BD_OrigF.sort_values('SAMPLE_ID')
           self.BD_OrigF = BD_OrigF
           self.GBF.setEnabled(True)
           self.PGBF.setEnabled(True)
           self.GBG.setEnabled(True)
           self.PGBG.setEnabled(True)
           self.GDC.setEnabled(True)
           self.QQDC.setEnabled(True)
           self.RDDC.setEnabled(True)
           self.GRG.setEnabled(True)
           self.QQRG.setEnabled(True)
           self.RDRG.setEnabled(True)
           self.GDP.setEnabled(True)
           self.GDP.setEnabled(True)
           self.RDDP.setEnabled(True)
           self.CHNB.setEnabled(True)
           self.CHNA.setEnabled(True)
           self.CHNM.setEnabled(True)
           self.GENR.setEnabled(True)
           self.QQDP.setEnabled(True)
           self.TextImp.setText("Creacion de Columnas y Variables Completado, genera el reporte e ingresa a las pestañas")
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si se modifico el encabezado de algun excel"
            QMessageBox.critical(self, "Revisar Encabezado", mensaje)
            self.error1.setText('Revisar si se modifico el encabezado de algun excel')

    def GraficoBF(self):
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['CU'], 'o-', color='gray', markerfacecolor='yellow', label='Cu')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['Cu_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino CU', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.CuErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['CU'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Fino CU.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['AG'], 'o-', color='gray', markerfacecolor='yellow', label='Ag')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['Ag_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino AG', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.AgErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['AG'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Fino AG.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Mo
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['MO'], 'o-', color='gray', markerfacecolor='yellow', label='Mo')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['Mo_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino MO', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.MoErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['MO'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Fino MO.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #as
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['AS'], 'o-', color='gray', markerfacecolor='yellow', label='As')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['As_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino AS', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.AsErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['AS'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Fino AS.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['ZN'], 'o-', color='gray', markerfacecolor='yellow', label='Zn')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['Zn_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino ZN', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.ZnErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['ZN'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Fino ZN.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #FE
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['FE'], 'o-', color='gray', markerfacecolor='yellow', label='Fe')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['FeFin_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino FE', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.FeErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['FE'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_addD(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nFE: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Fino FE.png', dpi=300, bbox_inches='tight')
        plt.show()

    def GrafPrecBF(self):
        
        try:
            #cu prec
            X = sm.add_constant(self.BD_BLKFINOF['CU_P'])
            model = sm.OLS(self.BD_BLKFINOF['CU'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLKFINOF['CU_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLKFINOF['CU_P'], self.BD_BLKFINOF['CU'], color='blue', label='CU')
            ax.plot(self.BD_BLKFINOF['CU_P'], self.BD_BLKFINOF['Cu_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLKFINOF['CU_P'], trend_line, color='green', linestyle='--', label='Linear Cu')
            ax.set_title('Muestra Precedente vs Blanco Fino Cu', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLKFINOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu Prec: {x:.3f}\nCu: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Fino Cu.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #ag prec
            X = sm.add_constant(self.BD_BLKFINOF['AG_P'])
            model = sm.OLS(self.BD_BLKFINOF['AG'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLKFINOF['AG_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLKFINOF['AG_P'], self.BD_BLKFINOF['AG'], color='blue', label='AG')
            ax.plot(self.BD_BLKFINOF['AG_P'], self.BD_BLKFINOF['Ag_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLKFINOF['AG_P'], trend_line, color='green', linestyle='--', label='Linear Ag')
            ax.set_title('Muestra Precedente vs Blanco Fino Ag', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLKFINOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg Prec: {x:.3f}\nAg: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Fino Ag.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #MO prec
            X = sm.add_constant(self.BD_BLKFINOF['MO_P'])
            model = sm.OLS(self.BD_BLKFINOF['MO'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLKFINOF['MO_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLKFINOF['MO_P'], self.BD_BLKFINOF['MO'], color='blue', label='MO')
            ax.plot(self.BD_BLKFINOF['MO_P'], self.BD_BLKFINOF['Mo_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLKFINOF['MO_P'], trend_line, color='green', linestyle='--', label='Linear Mo')
            ax.set_title('Muestra Precedente vs Blanco Fino Mo', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLKFINOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo Prec: {x:.3f}\nMo: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Fino Mo.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #AS prec
            X = sm.add_constant(self.BD_BLKFINOF['AS_P'])
            model = sm.OLS(self.BD_BLKFINOF['AS'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLKFINOF['AS_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLKFINOF['AS_P'], self.BD_BLKFINOF['AS'], color='blue', label='AS')
            ax.plot(self.BD_BLKFINOF['AS_P'], self.BD_BLKFINOF['As_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLKFINOF['AS_P'], trend_line, color='green', linestyle='--', label='Linear As')
            ax.set_title('Muestra Precedente vs Blanco Fino As', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLKFINOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs Prec: {x:.3f}\nAs: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Fino As.png', dpi=300, bbox_inches='tight')
            plt.show()       
    
            #ZN prec
            X = sm.add_constant(self.BD_BLKFINOF['ZN_P'])
            model = sm.OLS(self.BD_BLKFINOF['ZN'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLKFINOF['ZN_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLKFINOF['ZN_P'], self.BD_BLKFINOF['ZN'], color='blue', label='ZN')
            ax.plot(self.BD_BLKFINOF['ZN_P'], self.BD_BLKFINOF['Zn_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLKFINOF['ZN_P'], trend_line, color='green', linestyle='--', label='Linear Zn')
            ax.set_title('Muestra Precedente vs Blanco Fino Zn', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLKFINOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn Prec: {x:.3f}\nZn: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Fino Zn.png', dpi=300, bbox_inches='tight')
            plt.show()  
            
            #FE prec
            X = sm.add_constant(self.BD_BLKFINOF['FE_P'])
            model = sm.OLS(self.BD_BLKFINOF['FE'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLKFINOF['FE_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLKFINOF['FE_P'], self.BD_BLKFINOF['FE'], color='blue', label='FE')
            ax.plot(self.BD_BLKFINOF['FE_P'], self.BD_BLKFINOF['FeFin_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLKFINOF['FE_P'], trend_line, color='green', linestyle='--', label='Linear Fe')
            ax.set_title('Muestra Precedente vs Blanco Fino Fe', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLKFINOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nFe Prec: {x:.3f}\nFe: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Fino Fe.png', dpi=300, bbox_inches='tight')
            plt.show()   
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con muestra blanca, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')

    def GraficoBG(self):
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['CU'], 'o-', color='gray', markerfacecolor='yellow', label='Cu')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Cu_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso CU', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.CuErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['CU'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Grueso CU.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['AG'], 'o-', color='gray', markerfacecolor='yellow', label='Ag')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Ag_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso AG', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.AgErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['AG'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Grueso AG.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Mo
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['MO'], 'o-', color='gray', markerfacecolor='yellow', label='Mo')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Mo_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso MO', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.MoErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['MO'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Grueso MO.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #as
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['AS'], 'o-', color='gray', markerfacecolor='yellow', label='As')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['As_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso AS', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.AsErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['AS'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Grueso AS.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['ZN'], 'o-', color='gray', markerfacecolor='yellow', label='Zn')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Zn_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso ZN', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.ZnErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['ZN'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Grueso ZN.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #FE
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['FE'], 'o-', color='gray', markerfacecolor='yellow', label='Fe')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Fe_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso FE', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.FeErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['FE'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_addD(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nFE: {y}')
        fig.tight_layout()
        plt.savefig('Blanco Grueso FE.png', dpi=300, bbox_inches='tight')
        plt.show()

    def GrafPrecBG(self):
        try:
            #cu prec
            X = sm.add_constant(self.BD_BLANCOF['CU_P'])
            model = sm.OLS(self.BD_BLANCOF['CU'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLANCOF['CU_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLANCOF['CU_P'], self.BD_BLANCOF['CU'], color='blue', label='CU')
            ax.plot(self.BD_BLANCOF['CU_P'], self.BD_BLANCOF['Cu_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLANCOF['CU_P'], trend_line, color='green', linestyle='--', label='Linear Cu')
            ax.set_title('Muestra Precedente vs Blanco Grueso Cu', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLANCOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu Prec: {x:.3f}\nCu: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Grueso Cu.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #ag prec
            X = sm.add_constant(self.BD_BLANCOF['AG_P'])
            model = sm.OLS(self.BD_BLANCOF['AG'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLANCOF['AG_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLANCOF['AG_P'], self.BD_BLANCOF['AG'], color='blue', label='AG')
            ax.plot(self.BD_BLANCOF['AG_P'], self.BD_BLANCOF['Ag_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLANCOF['AG_P'], trend_line, color='green', linestyle='--', label='Linear Ag')
            ax.set_title('Muestra Precedente vs Blanco Grueso Ag', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLANCOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg Prec: {x:.3f}\nAg: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Grueso Ag.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #MO prec
            X = sm.add_constant(self.BD_BLANCOF['MO_P'])
            model = sm.OLS(self.BD_BLANCOF['MO'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLANCOF['MO_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLANCOF['MO_P'], self.BD_BLANCOF['MO'], color='blue', label='MO')
            ax.plot(self.BD_BLANCOF['MO_P'], self.BD_BLANCOF['Mo_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLANCOF['MO_P'], trend_line, color='green', linestyle='--', label='Linear Mo')
            ax.set_title('Muestra Precedente vs Blanco Grueso Mo', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLANCOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo Prec: {x:.3f}\nMo: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Grueso Mo.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #AS prec
            X = sm.add_constant(self.BD_BLANCOF['AS_P'])
            model = sm.OLS(self.BD_BLANCOF['AS'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLANCOF['AS_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLANCOF['AS_P'], self.BD_BLANCOF['AS'], color='blue', label='AS')
            ax.plot(self.BD_BLANCOF['AS_P'], self.BD_BLANCOF['As_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLANCOF['AS_P'], trend_line, color='green', linestyle='--', label='Linear As')
            ax.set_title('Muestra Precedente vs Blanco Grueso As', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLANCOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs Prec: {x:.3f}\nAs: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Grueso As.png', dpi=300, bbox_inches='tight')
            plt.show()       
    
            #ZN prec
            X = sm.add_constant(self.BD_BLANCOF['ZN_P'])
            model = sm.OLS(self.BD_BLANCOF['ZN'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLANCOF['ZN_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLANCOF['ZN_P'], self.BD_BLANCOF['ZN'], color='blue', label='ZN')
            ax.plot(self.BD_BLANCOF['ZN_P'], self.BD_BLANCOF['Zn_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLANCOF['ZN_P'], trend_line, color='green', linestyle='--', label='Linear Zn')
            ax.set_title('Muestra Precedente vs Blanco Grueso Zn', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLANCOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn Prec: {x:.3f}\nZn: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Grueso Zn.png', dpi=300, bbox_inches='tight')
            plt.show()  
            
            #FE prec
            X = sm.add_constant(self.BD_BLANCOF['FE_P'])
            model = sm.OLS(self.BD_BLANCOF['FE'], X).fit()
            slope, intercept = model.params[1], model.params[0]
            trend_line = intercept + slope * self.BD_BLANCOF['FE_P']
            fig, ax = plt.subplots(figsize=(10, 6))
            scatter = ax.scatter(self.BD_BLANCOF['FE_P'], self.BD_BLANCOF['FE'], color='blue', label='FE')
            ax.plot(self.BD_BLANCOF['FE_P'], self.BD_BLANCOF['Fe_MAV'], color='red', label='Max Accepted Value')
            ax.plot(self.BD_BLANCOF['FE_P'], trend_line, color='green', linestyle='--', label='Linear Fe')
            ax.set_title('Muestra Precedente vs Blanco Grueso Fe', fontsize=14, pad=20)
            ax.set_ylabel('Ley Muestra Blanca')
            ax.set_xlabel('Ley Muestra Precedente')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3)
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_BLANCOF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nFe Prec: {x:.3f}\nFe: {y:.3f}')
            fig.tight_layout()
            plt.savefig('Muestra Precedente vs Blanco Grueso Fe.png', dpi=300, bbox_inches='tight')
            plt.show()   
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con muestra blanca, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')

    def GraficoDC(self):
        try:
            #Cu
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'CU'), (1, 'red', 'x', 'CU Error')]:
                df_subset = self.BD_OrigF[self.BD_OrigF['Cu_Error'] == error_type]
                scatter = ax.scatter(df_subset['CU_MIN'], df_subset['CU_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu MAX: {y:.2f}\nCu MIN: {x:.2f}')
            ax.plot([0, self.BD_OrigF['CU_MAX'].max()], [0, self.BD_OrigF['CU_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_OrigF['XLimCU'], self.BD_OrigF['YLimCU'], color='red', label='Fun. Hiper Cu')
            ax.set_title('DUPLICADO DE CAMPO CU', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.CuErrorDUPC:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO DE CAMPO CU.png', dpi=300, bbox_inches='tight')
            plt.show()
    
            #AG
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'AG'), (1, 'red', 'x', 'AG Error')]:
                df_subset = self.BD_OrigF[self.BD_OrigF['Ag_Error'] == error_type]
                scatter = ax.scatter(df_subset['AG_MIN'], df_subset['AG_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg MAX: {y:.2f}\nAg MIN: {x:.2f}')
            ax.plot([0, self.BD_OrigF['AG_MAX'].max()], [0, self.BD_OrigF['AG_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_OrigF['XLimAG'], self.BD_OrigF['YLimAG'], color='red', label='Fun. Hiper Ag')
            ax.set_title('DUPLICADO DE CAMPO AG', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.AgErrorDUPC:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO DE CAMPO AG.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Mo
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'MO'), (1, 'red', 'x', 'MO Error')]:
                df_subset = self.BD_OrigF[self.BD_OrigF['Mo_Error'] == error_type]
                scatter = ax.scatter(df_subset['MO_MIN'], df_subset['MO_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo MAX: {y:.2f}\nMo MIN: {x:.2f}')
            ax.plot([0, self.BD_OrigF['MO_MAX'].max()], [0, self.BD_OrigF['MO_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_OrigF['XLimMO'], self.BD_OrigF['YLimMO'], color='red', label='Fun. Hiper Mo')
            ax.set_title('DUPLICADO DE CAMPO MO', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.MoErrorDUPC:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO DE CAMPO MO.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #As
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'AS'), (1, 'red', 'x', 'AS Error')]:
                df_subset = self.BD_OrigF[self.BD_OrigF['As_Error'] == error_type]
                scatter = ax.scatter(df_subset['AS_MIN'], df_subset['AS_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs MAX: {y:.2f}\nAs MIN: {x:.2f}')
            ax.plot([0, self.BD_OrigF['AS_MAX'].max()], [0, self.BD_OrigF['AS_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_OrigF['XLimAS'], self.BD_OrigF['YLimAS'], color='red', label='Fun. Hiper As')
            ax.set_title('DUPLICADO DE CAMPO AS', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.AsErrorDUPC:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO DE CAMPO AS.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #ZN
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'ZN'), (1, 'red', 'x', 'ZN Error')]:
                df_subset = self.BD_OrigF[self.BD_OrigF['Zn_Error'] == error_type]
                scatter = ax.scatter(df_subset['ZN_MIN'], df_subset['ZN_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn MAX: {y:.2f}\nZn MIN: {x:.2f}')
            ax.plot([0, self.BD_OrigF['ZN_MAX'].max()], [0, self.BD_OrigF['ZN_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_OrigF['XLimZN'], self.BD_OrigF['YLimZN'], color='red', label='Fun. Hiper Zn')
            ax.set_title('DUPLICADO DE CAMPO ZN', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.ZnErrorDUPC:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO DE CAMPO ZN.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            
            #Fe
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'FE'), (1, 'red', 'x', 'FE Error')]:
                df_subset = self.BD_OrigF[self.BD_OrigF['Fe_Error'] == error_type]
                scatter = ax.scatter(df_subset['FE_MIN'], df_subset['FE_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nFe MAX: {y:.2f}\nFe MIN: {x:.2f}')
            ax.plot([0, self.BD_OrigF['FE_MAX'].max()], [0, self.BD_OrigF['FE_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_OrigF['XLimFE'], self.BD_OrigF['YLimFE'], color='red', label='Fun. Hiper Fe')
            ax.set_title('DUPLICADO DE CAMPO FE', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.FeErrorDUPC:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO DE CAMPO FE.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de campo, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')

    def QQDupC(self):
        try:
            #Cu
            quantiles_CU_D = np.percentile(self.BD_OrigF['CU_D'], np.linspace(0, 100, 101))
            quantiles_CU = np.percentile(self.BD_OrigF['CU'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_CU, quantiles_CU, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_CU, quantiles_CU_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Cu Orig: {x:.2f}\nCu Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DC CU', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot CU DC.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Ag
            quantiles_AG_D = np.percentile(self.BD_OrigF['AG_D'], np.linspace(0, 100, 101))
            quantiles_AG = np.percentile(self.BD_OrigF['AG'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_AG, quantiles_AG, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_AG, quantiles_AG_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Ag Orig: {x:.2f}\nAg Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DC AG', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot AG DC.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #mo
            quantiles_MO_D = np.percentile(self.BD_OrigF['MO_D'], np.linspace(0, 100, 101))
            quantiles_MO = np.percentile(self.BD_OrigF['MO'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_MO, quantiles_MO, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_MO, quantiles_MO_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Mo Orig: {x:.2f}\nMo Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DC MO', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot MO DC.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #as
            quantiles_AS_D = np.percentile(self.BD_OrigF['AS_D'], np.linspace(0, 100, 101))
            quantiles_AS = np.percentile(self.BD_OrigF['AS'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_AS, quantiles_AS, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_AS, quantiles_AS_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'As Orig: {x:.2f}\nAs Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DC AS', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot AS DC.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #zn
            quantiles_ZN_D = np.percentile(self.BD_OrigF['ZN_D'], np.linspace(0, 100, 101))
            quantiles_ZN = np.percentile(self.BD_OrigF['ZN'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_ZN, quantiles_ZN, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_ZN, quantiles_ZN_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Zn Orig: {x:.2f}\nZn Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DC ZN', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot ZN DC.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #fe
            quantiles_FE_D = np.percentile(self.BD_OrigF['FE_D'], np.linspace(0, 100, 101))
            quantiles_FE = np.percentile(self.BD_OrigF['FE'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_FE, quantiles_FE, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_FE, quantiles_FE_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Fe Orig: {x:.2f}\nFe Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DC FE', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot FE DC.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de campo, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')

    def RDDupC(self):
        try:
            
            #Cu
            self.BD_OrigF = self.BD_OrigF.sort_values('CuProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_OrigF['MediaMovCuProm'], self.BD_OrigF['CuRD'], color='blue', label='RD Plot Cu')
            ax.plot(self.BD_OrigF['CuProm'], self.BD_OrigF['CuRD'], color='red', label='RD Plot Cu')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu_prom: {x:.2f}\nRD Cu: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DC Cu RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DC Cu RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_OrigF = self.BD_OrigF.sort_values('SAMPLE_ID')
            
            #Ag
            self.BD_OrigF = self.BD_OrigF.sort_values('AgProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_OrigF['MediaMovAgProm'], self.BD_OrigF['AgRD'], color='blue', label='RD Plot Ag')
            ax.plot(self.BD_OrigF['AgProm'], self.BD_OrigF['AgRD'], color='red', label='RD Plot Ag')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg_prom: {x:.2f}\nRD Ag: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DC Ag RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DC Ag RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_OrigF = self.BD_OrigF.sort_values('SAMPLE_ID')
            
            #mo
            self.BD_OrigF = self.BD_OrigF.sort_values('MoProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_OrigF['MediaMovMoProm'], self.BD_OrigF['MoRD'], color='blue', label='RD Plot Mo')
            ax.plot(self.BD_OrigF['MoProm'], self.BD_OrigF['MoRD'], color='red', label='RD Plot Mo')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo_prom: {x:.2f}\nRD Mo: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DC Mo RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DC Mo RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_OrigF = self.BD_OrigF.sort_values('SAMPLE_ID')
            
            #as
            self.BD_OrigF = self.BD_OrigF.sort_values('AsProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_OrigF['MediaMovAsProm'], self.BD_OrigF['AsRD'], color='blue', label='RD Plot As')
            ax.plot(self.BD_OrigF['AsProm'], self.BD_OrigF['AsRD'], color='red', label='RD Plot As')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs_prom: {x:.2f}\nRD As: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DC As RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DC As RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_OrigF = self.BD_OrigF.sort_values('SAMPLE_ID')
            
            #zn
            self.BD_OrigF = self.BD_OrigF.sort_values('ZnProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_OrigF['MediaMovZnProm'], self.BD_OrigF['ZnRD'], color='blue', label='RD Plot Zn')
            ax.plot(self.BD_OrigF['ZnProm'], self.BD_OrigF['ZnRD'], color='red', label='RD Plot Zn')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn_prom: {x:.2f}\nRD Zn: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DC Zn RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DC Zn RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_OrigF = self.BD_OrigF.sort_values('SAMPLE_ID')
            
            #Fe
            self.BD_OrigF = self.BD_OrigF.sort_values('FeProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_OrigF['MediaMovFeProm'], self.BD_OrigF['FeRD'], color='blue', label='RD Plot Fe')
            ax.plot(self.BD_OrigF['FeProm'], self.BD_OrigF['FeRD'], color='red', label='RD Plot Fe')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_OrigF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nFe_prom: {x:.2f}\nRD Fe: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DC Fe RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DC Fe RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_OrigF = self.BD_OrigF.sort_values('SAMPLE_ID')
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de campo, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')

    def GraficoRG(self):
        try:
            #Cu
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'CU'), (1, 'red', 'x', 'CU Error')]:
                df_subset = self.BD_RECHAF[self.BD_RECHAF['Cu_Error'] == error_type]
                scatter = ax.scatter(df_subset['CU_MIN'], df_subset['CU_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu MAX: {y:.2f}\nCu MIN: {x:.2f}')
            ax.plot([0, self.BD_RECHAF['CU_MAX'].max()], [0, self.BD_RECHAF['CU_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_RECHAF['XLimCU'], self.BD_RECHAF['YLimCU'], color='red', label='Fun. Hiper Cu')
            ax.set_title('RECHAZO GRUESO CU', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.CuErrorRG:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RECHAZO GRUESO CU.png', dpi=300, bbox_inches='tight')
            plt.show()
    
            #AG
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'AG'), (1, 'red', 'x', 'AG Error')]:
                df_subset = self.BD_RECHAF[self.BD_RECHAF['Ag_Error'] == error_type]
                scatter = ax.scatter(df_subset['AG_MIN'], df_subset['AG_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg MAX: {y:.2f}\nAg MIN: {x:.2f}')
            ax.plot([0, self.BD_RECHAF['AG_MAX'].max()], [0, self.BD_RECHAF['AG_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_RECHAF['XLimAG'], self.BD_RECHAF['YLimAG'], color='red', label='Fun. Hiper Ag')
            ax.set_title('RECHAZO GRUESO AG', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.AgErrorRG:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RECHAZO GRUESO AG.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Mo
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'MO'), (1, 'red', 'x', 'MO Error')]:
                df_subset = self.BD_RECHAF[self.BD_RECHAF['Mo_Error'] == error_type]
                scatter = ax.scatter(df_subset['MO_MIN'], df_subset['MO_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo MAX: {y:.2f}\nMo MIN: {x:.2f}')
            ax.plot([0, self.BD_RECHAF['MO_MAX'].max()], [0, self.BD_RECHAF['MO_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_RECHAF['XLimMO'], self.BD_RECHAF['YLimMO'], color='red', label='Fun. Hiper Mo')
            ax.set_title('RECHAZO GRUESO MO', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.MoErrorRG:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RECHAZO GRUESO MO.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #As
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'AS'), (1, 'red', 'x', 'AS Error')]:
                df_subset = self.BD_RECHAF[self.BD_RECHAF['As_Error'] == error_type]
                scatter = ax.scatter(df_subset['AS_MIN'], df_subset['AS_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs MAX: {y:.2f}\nAs MIN: {x:.2f}')
            ax.plot([0, self.BD_RECHAF['AS_MAX'].max()], [0, self.BD_RECHAF['AS_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_RECHAF['XLimAS'], self.BD_RECHAF['YLimAS'], color='red', label='Fun. Hiper As')
            ax.set_title('RECHAZO GRUESO AS', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.AsErrorRG:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RECHAZO GRUESO AS.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #ZN
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'ZN'), (1, 'red', 'x', 'ZN Error')]:
                df_subset = self.BD_RECHAF[self.BD_RECHAF['Zn_Error'] == error_type]
                scatter = ax.scatter(df_subset['ZN_MIN'], df_subset['ZN_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn MAX: {y:.2f}\nZn MIN: {x:.2f}')
            ax.plot([0, self.BD_RECHAF['ZN_MAX'].max()], [0, self.BD_RECHAF['ZN_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_RECHAF['XLimZN'], self.BD_RECHAF['YLimZN'], color='red', label='Fun. Hiper Zn')
            ax.set_title('RECHAZO GRUESO ZN', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.ZnErrorRG:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RECHAZO GRUESO ZN.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            
            #Fe
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'FE'), (1, 'red', 'x', 'FE Error')]:
                df_subset = self.BD_RECHAF[self.BD_RECHAF['Fe_Error'] == error_type]
                scatter = ax.scatter(df_subset['FE_MIN'], df_subset['FE_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nFe MAX: {y:.2f}\nFe MIN: {x:.2f}')
            ax.plot([0, self.BD_RECHAF['FE_MAX'].max()], [0, self.BD_RECHAF['FE_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_RECHAF['XLimFE'], self.BD_RECHAF['YLimFE'], color='red', label='Fun. Hiper Fe')
            ax.set_title('RECHAZO GRUESO FE', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.FeErrorRG:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RECHAZO GRUESO FE.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de rechazo grueso, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')
        
    def QQDupRG(self):
        try:
            #Cu
            quantiles_CU_D = np.percentile(self.BD_RECHAF['CU'], np.linspace(0, 100, 101))
            quantiles_CU = np.percentile(self.BD_RECHAF['CU_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_CU, quantiles_CU, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_CU, quantiles_CU_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Cu Orig: {x:.2f}\nCu Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot RG CU', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot CU RG.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Ag
            quantiles_AG_D = np.percentile(self.BD_RECHAF['AG'], np.linspace(0, 100, 101))
            quantiles_AG = np.percentile(self.BD_RECHAF['AG_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_AG, quantiles_AG, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_AG, quantiles_AG_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Ag Orig: {x:.2f}\nAg Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot RG AG', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot AG RG.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #mo
            quantiles_MO_D = np.percentile(self.BD_RECHAF['MO'], np.linspace(0, 100, 101))
            quantiles_MO = np.percentile(self.BD_RECHAF['MO_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_MO, quantiles_MO, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_MO, quantiles_MO_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Mo Orig: {x:.2f}\nMo Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot RG MO', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot MO RG.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #as
            quantiles_AS_D = np.percentile(self.BD_RECHAF['AS'], np.linspace(0, 100, 101))
            quantiles_AS = np.percentile(self.BD_RECHAF['AS_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_AS, quantiles_AS, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_AS, quantiles_AS_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'As Orig: {x:.2f}\nAs Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot RG AS', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot AS RG.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #zn
            quantiles_ZN_D = np.percentile(self.BD_RECHAF['ZN'], np.linspace(0, 100, 101))
            quantiles_ZN = np.percentile(self.BD_RECHAF['ZN_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_ZN, quantiles_ZN, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_ZN, quantiles_ZN_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Zn Orig: {x:.2f}\nZn Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot RG ZN', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot ZN RG.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #fe
            quantiles_FE_D = np.percentile(self.BD_RECHAF['FE'], np.linspace(0, 100, 101))
            quantiles_FE = np.percentile(self.BD_RECHAF['FE_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_FE, quantiles_FE, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_FE, quantiles_FE_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Fe Orig: {x:.2f}\nFe Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot RG FE', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot FE RG.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de rechazo grueso, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')
        

    def RDRGC(self):
        try:
            #Cu
            self.BD_RECHAF = self.BD_RECHAF.sort_values('CuProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RECHAF['MediaMovCuProm'], self.BD_RECHAF['CuRD'], color='blue', label='RD Plot Cu')
            ax.plot(self.BD_RECHAF['CuProm'], self.BD_RECHAF['CuRD'], color='red', label='RD Plot Cu')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu_prom: {x:.2f}\nRD Cu: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RG Cu RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RG Cu RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RECHAF = self.BD_RECHAF.sort_values('SAMPLE_ID')
            
            #Ag
            self.BD_RECHAF = self.BD_RECHAF.sort_values('AgProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RECHAF['MediaMovAgProm'], self.BD_RECHAF['AgRD'], color='blue', label='RD Plot Ag')
            ax.plot(self.BD_RECHAF['AgProm'], self.BD_RECHAF['AgRD'], color='red', label='RD Plot Ag')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg_prom: {x:.2f}\nRD Ag: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RG Ag RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RG Ag RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RECHAF = self.BD_RECHAF.sort_values('SAMPLE_ID')
            
            #mo
            self.BD_RECHAF = self.BD_RECHAF.sort_values('MoProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RECHAF['MediaMovMoProm'], self.BD_RECHAF['MoRD'], color='blue', label='RD Plot Mo')
            ax.plot(self.BD_RECHAF['MoProm'], self.BD_RECHAF['MoRD'], color='red', label='RD Plot Mo')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo_prom: {x:.2f}\nRD Mo: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RG Mo RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RG Mo RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RECHAF = self.BD_RECHAF.sort_values('SAMPLE_ID')
            
            #as
            self.BD_RECHAF = self.BD_RECHAF.sort_values('AsProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RECHAF['MediaMovAsProm'], self.BD_RECHAF['AsRD'], color='blue', label='RD Plot As')
            ax.plot(self.BD_RECHAF['AsProm'], self.BD_RECHAF['AsRD'], color='red', label='RD Plot As')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs_prom: {x:.2f}\nRD As: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RG As RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RG As RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RECHAF = self.BD_RECHAF.sort_values('SAMPLE_ID')
            
            #zn
            self.BD_RECHAF = self.BD_RECHAF.sort_values('ZnProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RECHAF['MediaMovZnProm'], self.BD_RECHAF['ZnRD'], color='blue', label='RD Plot Zn')
            ax.plot(self.BD_RECHAF['ZnProm'], self.BD_RECHAF['ZnRD'], color='red', label='RD Plot Zn')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn_prom: {x:.2f}\nRD Zn: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RG Zn RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RG Zn RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RECHAF = self.BD_RECHAF.sort_values('SAMPLE_ID')
            
            #Fe
            self.BD_RECHAF = self.BD_RECHAF.sort_values('FeProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RECHAF['MediaMovFeProm'], self.BD_RECHAF['FeRD'], color='blue', label='RD Plot Fe')
            ax.plot(self.BD_RECHAF['FeProm'], self.BD_RECHAF['FeRD'], color='red', label='RD Plot Fe')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RECHAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nFe_prom: {x:.2f}\nRD Fe: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RG Fe RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RG Fe RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RECHAF = self.BD_RECHAF.sort_values('SAMPLE_ID')
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de rechazo grueso, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')
        

    def GraficoDP(self):
        try:
            #Cu
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'CU'), (1, 'red', 'x', 'CU Error')]:
                df_subset = self.BD_PULPAF[self.BD_PULPAF['Cu_Error'] == error_type]
                scatter = ax.scatter(df_subset['CU_MIN'], df_subset['CU_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu MAX: {y:.2f}\nCu MIN: {x:.2f}')
            ax.plot([0, self.BD_PULPAF['CU_MAX'].max()], [0, self.BD_PULPAF['CU_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_PULPAF['XLimCU'], self.BD_PULPAF['YLimCU'], color='red', label='Fun. Hiper Cu')
            ax.set_title('DUPLICADO PULPA CU', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.CuErrorPL:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO PULPA CU.png', dpi=300, bbox_inches='tight')
            plt.show()
    
            #AG
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'AG'), (1, 'red', 'x', 'AG Error')]:
                df_subset = self.BD_PULPAF[self.BD_PULPAF['Ag_Error'] == error_type]
                scatter = ax.scatter(df_subset['AG_MIN'], df_subset['AG_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg MAX: {y:.2f}\nAg MIN: {x:.2f}')
            ax.plot([0, self.BD_PULPAF['AG_MAX'].max()], [0, self.BD_PULPAF['AG_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_PULPAF['XLimAG'], self.BD_PULPAF['YLimAG'], color='red', label='Fun. Hiper Ag')
            ax.set_title('DUPLICADO PULPA AG', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.AgErrorPL:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO PULPA AG.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Mo
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'MO'), (1, 'red', 'x', 'MO Error')]:
                df_subset = self.BD_PULPAF[self.BD_PULPAF['Mo_Error'] == error_type]
                scatter = ax.scatter(df_subset['MO_MIN'], df_subset['MO_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo MAX: {y:.2f}\nMo MIN: {x:.2f}')
            ax.plot([0, self.BD_PULPAF['MO_MAX'].max()], [0, self.BD_PULPAF['MO_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_PULPAF['XLimMO'], self.BD_PULPAF['YLimMO'], color='red', label='Fun. Hiper Mo')
            ax.set_title('DUPLICADO PULPA MO', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.MoErrorPL:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO PULPA MO.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #As
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'AS'), (1, 'red', 'x', 'AS Error')]:
                df_subset = self.BD_PULPAF[self.BD_PULPAF['As_Error'] == error_type]
                scatter = ax.scatter(df_subset['AS_MIN'], df_subset['AS_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs MAX: {y:.2f}\nAs MIN: {x:.2f}')
            ax.plot([0, self.BD_PULPAF['AS_MAX'].max()], [0, self.BD_PULPAF['AS_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_PULPAF['XLimAS'], self.BD_PULPAF['YLimAS'], color='red', label='Fun. Hiper As')
            ax.set_title('DUPLICADO PULPA AS', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.AsErrorPL:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO PULPA AS.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #ZN
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'ZN'), (1, 'red', 'x', 'ZN Error')]:
                df_subset = self.BD_PULPAF[self.BD_PULPAF['Zn_Error'] == error_type]
                scatter = ax.scatter(df_subset['ZN_MIN'], df_subset['ZN_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn MAX: {y:.2f}\nZn MIN: {x:.2f}')
            ax.plot([0, self.BD_PULPAF['ZN_MAX'].max()], [0, self.BD_PULPAF['ZN_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_PULPAF['XLimZN'], self.BD_PULPAF['YLimZN'], color='red', label='Fun. Hiper Zn')
            ax.set_title('DUPLICADO PULPA ZN', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.ZnErrorPL:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO PULPA ZN.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            
            #Fe
            fig, ax = plt.subplots(figsize=(8, 8))
            for error_type, color, marker, legend_name in [(0, 'orange', 'o', 'FE'), (1, 'red', 'x', 'FE Error')]:
                df_subset = self.BD_PULPAF[self.BD_PULPAF['Fe_Error'] == error_type]
                scatter = ax.scatter(df_subset['FE_MIN'], df_subset['FE_MAX'], color=color, marker=marker, s=20, label=legend_name)
                cursor = mplcursors.cursor(scatter, hover=True)
                @cursor.connect("add")
                def on_add(sel):
                    x, y = sel.target
                    index = sel.index
                    sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                    sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nFe MAX: {y:.2f}\nFe MIN: {x:.2f}')
            ax.plot([0, self.BD_PULPAF['FE_MAX'].max()], [0, self.BD_PULPAF['FE_MAX'].max()], color='black', label='Y=X')
            ax.plot(self.BD_PULPAF['XLimFE'], self.BD_PULPAF['YLimFE'], color='red', label='Fun. Hiper Fe')
            ax.set_title('DUPLICADO PULPA FE', fontsize=14, pad=20)
            ax.set_xlabel('Valor Minimo')
            ax.set_ylabel('Valor Maximo')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.annotate(
                f"Error: {self.FeErrorPL:.2f}%",
                xy=(0.7, 0.1),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DUPLICADO PULPA FE.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de pulpa, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')
        
        
    def QQDupDP(self):
        try:
            #Cu
            quantiles_CU_D = np.percentile(self.BD_PULPAF['CU'], np.linspace(0, 100, 101))
            quantiles_CU = np.percentile(self.BD_PULPAF['CU_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_CU, quantiles_CU, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_CU, quantiles_CU_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Cu Orig: {x:.2f}\nCu Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DP CU', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot CU DP.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Ag
            quantiles_AG_D = np.percentile(self.BD_PULPAF['AG'], np.linspace(0, 100, 101))
            quantiles_AG = np.percentile(self.BD_PULPAF['AG_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_AG, quantiles_AG, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_AG, quantiles_AG_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Ag Orig: {x:.2f}\nAg Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DP AG', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot AG DP.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #mo
            quantiles_MO_D = np.percentile(self.BD_PULPAF['MO'], np.linspace(0, 100, 101))
            quantiles_MO = np.percentile(self.BD_PULPAF['MO_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_MO, quantiles_MO, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_MO, quantiles_MO_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Mo Orig: {x:.2f}\nMo Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DP MO', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot MO DP.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #as
            quantiles_AS_D = np.percentile(self.BD_PULPAF['AS'], np.linspace(0, 100, 101))
            quantiles_AS = np.percentile(self.BD_PULPAF['AS_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_AS, quantiles_AS, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_AS, quantiles_AS_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'As Orig: {x:.2f}\nAs Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DP AS', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot AS DP.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #zn
            quantiles_ZN_D = np.percentile(self.BD_PULPAF['ZN'], np.linspace(0, 100, 101))
            quantiles_ZN = np.percentile(self.BD_PULPAF['ZN_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_ZN, quantiles_ZN, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_ZN, quantiles_ZN_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Zn Orig: {x:.2f}\nZn Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DP ZN', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot ZN DP.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #fe
            quantiles_FE_D = np.percentile(self.BD_PULPAF['FE'], np.linspace(0, 100, 101))
            quantiles_FE = np.percentile(self.BD_PULPAF['FE_O'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quantiles_FE, quantiles_FE, color='red', label='Y=X')
            scatter = ax.scatter(quantiles_FE, quantiles_FE_D, color='blue', marker='x', label='Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Fe Orig: {x:.2f}\nFe Dup: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('Q-Q Plot DP FE', fontsize=14, pad=20)
            ax.set_xlabel('Original')
            ax.set_ylabel('Duplicado')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot FE DP.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de pulpa, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')
        

    def RDDPC(self):
        try:
            #Cu
            self.BD_PULPAF = self.BD_PULPAF.sort_values('CuProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_PULPAF['MediaMovCuProm'], self.BD_PULPAF['CuRD'], color='blue', label='RD Plot Cu')
            ax.plot(self.BD_PULPAF['CuProm'], self.BD_PULPAF['CuRD'], color='red', label='RD Plot Cu')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu_prom: {x:.2f}\nRD Cu: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DP Cu RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DP Cu RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_PULPAF = self.BD_PULPAF.sort_values('SAMPLE_ID')
            
            #Ag
            self.BD_PULPAF = self.BD_PULPAF.sort_values('AgProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_PULPAF['MediaMovAgProm'], self.BD_PULPAF['AgRD'], color='blue', label='RD Plot Ag')
            ax.plot(self.BD_PULPAF['AgProm'], self.BD_PULPAF['AgRD'], color='red', label='RD Plot Ag')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg_prom: {x:.2f}\nRD Ag: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DP Ag RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DP Ag RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_PULPAF = self.BD_PULPAF.sort_values('SAMPLE_ID')
            
            #mo
            self.BD_PULPAF = self.BD_PULPAF.sort_values('MoProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_PULPAF['MediaMovMoProm'], self.BD_PULPAF['MoRD'], color='blue', label='RD Plot Mo')
            ax.plot(self.BD_PULPAF['MoProm'], self.BD_PULPAF['MoRD'], color='red', label='RD Plot Mo')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo_prom: {x:.2f}\nRD Mo: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DP Mo RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DP Mo RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_PULPAF = self.BD_PULPAF.sort_values('SAMPLE_ID')
            
            #as
            self.BD_PULPAF = self.BD_PULPAF.sort_values('AsProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_PULPAF['MediaMovAsProm'], self.BD_PULPAF['AsRD'], color='blue', label='RD Plot As')
            ax.plot(self.BD_PULPAF['AsProm'], self.BD_PULPAF['AsRD'], color='red', label='RD Plot As')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs_prom: {x:.2f}\nRD As: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DP As RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DP As RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_PULPAF = self.BD_PULPAF.sort_values('SAMPLE_ID')
            
            #zn
            self.BD_PULPAF = self.BD_PULPAF.sort_values('ZnProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_PULPAF['MediaMovZnProm'], self.BD_PULPAF['ZnRD'], color='blue', label='RD Plot Zn')
            ax.plot(self.BD_PULPAF['ZnProm'], self.BD_PULPAF['ZnRD'], color='red', label='RD Plot Zn')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn_prom: {x:.2f}\nRD Zn: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DP Zn RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DP Zn RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_PULPAF = self.BD_PULPAF.sort_values('SAMPLE_ID')
            
            #Fe
            self.BD_PULPAF = self.BD_PULPAF.sort_values('FeProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_PULPAF['MediaMovFeProm'], self.BD_PULPAF['FeRD'], color='blue', label='RD Plot Fe')
            ax.plot(self.BD_PULPAF['FeProm'], self.BD_PULPAF['FeRD'], color='red', label='RD Plot Fe')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_PULPAF.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nFe_prom: {x:.2f}\nRD Fe: {y:.2f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('DP Fe RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('DP Fe RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_PULPAF = self.BD_PULPAF.sort_values('SAMPLE_ID')
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si en el rango insertado no se inicie con duplicado de pulpa, ya que no se encontraria muestra precedente"
            QMessageBox.critical(self, "Revisar Rango", mensaje)
            self.error1.setText('Error en lectura de muestra Inicial')
        

    def STDB(self):
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['CU'], 'o-', color='blue', label='Cu')
        lines = ['MHCU', 'BVCU', '+2STDCU', '-2STDCU', '-3STDCU', '+3STDCU']
        line_colors = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names = ['Media Historica Cu', 'BV Cu', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines, line_colors, line_names):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['CU'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Baja Ley CHNB-13 Cu', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)

        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDB_ERROR_CU:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 Cu.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['AG'], 'o-', color='blue', label='Ag')
        lines1 = ['MHAG',  '+2STDAG', '-2STDAG', '-3STDAG', '+3STDAG']
        line_colors1 = ['orange', 'brown', 'brown', 'red', 'red']
        line_names1 = ['Media Historica Ag',  'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines1, line_colors1, line_names1):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['AG'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)             
        ax.set_title('Estandar Baja Ley CHNB-13 Ag', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 Ag.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #AS
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['AS'], 'o-', color='blue', label='As')
        lines2 = ['MHAS',  '+2STDAS', '-2STDAS', '-3STDAS', '+3STDAS']
        line_colors2 = ['orange',  'brown', 'brown', 'red', 'red']
        line_names2 = ['Media Historica As', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines2, line_colors2, line_names2):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['AS'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)             
        ax.set_title('Estandar Baja Ley CHNB-13 As', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 As.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #MO
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['MO'], 'o-', color='blue', label='Mo')
        lines3 = ['MHMO', 'BVMO', '+2STDMO', '-2STDMO', '-3STDMO', '+3STDMO']
        line_colors3 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names3 = ['Media Historica Mo', 'BV Mo', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines3, line_colors3, line_names3):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['MO'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)             
        ax.set_title('Estandar Baja Ley CHNB-13 Mo', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDB_ERROR_MO:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 Mo.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['ZN'], 'o-', color='blue', label='Zn')
        lines4 = ['MHZN', 'BVZN', '+2STDZN', '-2STDZN', '-3STDZN', '+3STDZN']
        line_colors4 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names4 = ['Media Historica Zn', 'BV Zn', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines4, line_colors4, line_names4):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['ZN'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)             
        ax.set_title('Estandar Baja Ley CHNB-13 Zn', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDB_ERROR_ZN:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 Zn.png', dpi=300, bbox_inches='tight')
        plt.show()
           


    def STDM(self):  
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['CU'], 'o-', color='blue', label='Cu')
        lines = ['MHCU', 'BVCU', '+2STDCU', '-2STDCU', '-3STDCU', '+3STDCU']
        line_colors = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names = ['Media Historica Cu', 'BV Cu', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines, line_colors, line_names):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['CU'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)
        ax.set_title('Estandar Media Ley CHNM-14 Cu', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDM_ERROR_CU:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 Cu.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['AG'], 'o-', color='blue', label='Ag')
        lines1 = ['MHAG', '+2STDAG', '-2STDAG', '-3STDAG', '+3STDAG']
        line_colors1 = ['orange',  'brown', 'brown', 'red', 'red']
        line_names1 = ['Media Historica Ag',  'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines1, line_colors1, line_names1):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['AG'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Media Ley CHNM-14 Ag', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 Ag.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #AS
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['AS'], 'o-', color='blue', label='As')
        lines2 = ['MHAS', 'BVAS', '+2STDAS', '-2STDAS', '-3STDAS', '+3STDAS']
        line_colors2 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names2 = ['Media Historica As', 'BV As', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines2, line_colors2, line_names2):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['AS'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Media Ley CHNM-14 As', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDM_ERROR_AS:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 As.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #MO
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['MO'], 'o-', color='blue', label='Mo')
        lines3 = ['MHMO', 'BVMO', '+2STDMO', '-2STDMO', '-3STDMO', '+3STDMO']
        line_colors3 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names3 = ['Media Historica Mo', 'BV Mo', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines3, line_colors3, line_names3):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['MO'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Media Ley CHNM-14 Mo', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDM_ERROR_MO:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 Mo.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['ZN'], 'o-', color='blue', label='Zn')
        lines4 = ['MHZN', 'BVZN', '+2STDZN', '-2STDZN', '-3STDZN', '+3STDZN']
        line_colors4 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names4 = ['Media Historica Zn', 'BV Zn', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines4, line_colors4, line_names4):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['ZN'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Media Ley CHNM-14 Zn', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDM_ERROR_ZN:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 Zn.png', dpi=300, bbox_inches='tight')
        plt.show()
    
    def STDA(self):
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['CU'], 'o-', color='blue', label='Cu')
        lines = ['MHCU', 'BVCU', '+2STDCU', '-2STDCU', '-3STDCU', '+3STDCU']
        line_colors = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names = ['Media Historica Cu', 'BV Cu', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines, line_colors, line_names):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['CU'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Alta Ley CHNA-15 Cu', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDA_ERROR_CU:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 Cu.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['AG'], 'o-', color='blue', label='Ag')
        lines1 = ['MHAG', 'BVAG', '+2STDAG', '-2STDAG', '-3STDAG', '+3STDAG']
        line_colors1 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names1 = ['Media Historica Ag', 'BV Ag', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines1, line_colors1, line_names1):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['AG'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Alta Ley CHNA-15 Ag', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDA_ERROR_AG:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 Ag.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #AS
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['AS'], 'o-', color='blue', label='As')
        lines2 = ['MHAS', 'BVAS', '+2STDAS', '-2STDAS', '-3STDAS', '+3STDAS']
        line_colors2 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names2 = ['Media Historica As', 'BV As', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines2, line_colors2, line_names2):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['AS'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Alta Ley CHNA-15 As', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDA_ERROR_AS:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 As.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #MO
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['MO'], 'o-', color='blue', label='Mo')
        lines3 = ['MHMO', 'BVMO', '+2STDMO', '-2STDMO', '-3STDMO', '+3STDMO']
        line_colors3 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names3 = ['Media Historica Mo', 'BV Mo', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines3, line_colors3, line_names3):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['MO'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Alta Ley CHNA-15 Mo', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDA_ERROR_MO:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 Mo.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['ZN'], 'o-', color='blue', label='Zn')
        lines4 = ['MHZN', 'BVZN', '+2STDZN', '-2STDZN', '-3STDZN', '+3STDZN']
        line_colors4 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names4 = ['Media Historica Zn', 'BV Zn', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines4, line_colors4, line_names4):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['ZN'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Alta Ley CHNA-15 Zn', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        
        ax.annotate(f"Sesgo: {self.STDA_ERROR_ZN:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 Zn.png', dpi=300, bbox_inches='tight')
        plt.show()
    
    def STDH(self):
        self.BD_BLKFINOF['Mes'] = self.BD_BLKFINOF['MES ']  # Mes en español
        unique_months_BLF = self.BD_BLKFINOF.groupby('Mes')['SAMPLE_ID'].mean() # Posición promedio de cada mes

        self.BD_BLANCOF['Mes'] = self.BD_BLANCOF['MES '] # Mes en español
        unique_months_BG = self.BD_BLANCOF.groupby('Mes')['SAMPLE_ID'].mean()

        self.BD_STD13F['Mes'] = self.BD_STD13F['MES '] # Mes en español
        unique_months_STDB = self.BD_STD13F.groupby('Mes')['SAMPLE_ID'].mean()

        self.BD_STD14F['Mes'] = self.BD_STD14F['MES ']  # Mes en español
        unique_months_STDM = self.BD_STD14F.groupby('Mes')['SAMPLE_ID'].mean()

        self.BD_STD15F['Mes'] = self.BD_STD15F['MES ']  # Mes en español
        unique_months_STDA = self.BD_STD15F.groupby('Mes')['SAMPLE_ID'].mean()
        #Cu

        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['CU'], 'o-', color='gray', markerfacecolor='yellow', label='Cu')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['Cu_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino CU', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.CuErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['CU'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BLF.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLKFINOF['Mes'].ne(self.BD_BLKFINOF['Mes'].shift())
        month_boundaries = self.BD_BLKFINOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])    
        fig.tight_layout()
        plt.savefig('Blanco Fino CU Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['AG'], 'o-', color='gray', markerfacecolor='yellow', label='Ag')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['Ag_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino AG', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.AgErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['AG'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BLF.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLKFINOF['Mes'].ne(self.BD_BLKFINOF['Mes'].shift())
        month_boundaries = self.BD_BLKFINOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([]) 
        fig.tight_layout()
        plt.savefig('Blanco Fino AG Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Mo
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['MO'], 'o-', color='gray', markerfacecolor='yellow', label='Mo')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['Mo_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino MO', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.MoErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['MO'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BLF.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLKFINOF['Mes'].ne(self.BD_BLKFINOF['Mes'].shift())
        month_boundaries = self.BD_BLKFINOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])  
        fig.tight_layout()
        plt.savefig('Blanco Fino MO Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #as
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['AS'], 'o-', color='gray', markerfacecolor='yellow', label='As')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['As_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino AS', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.AsErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['AS'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BLF.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLKFINOF['Mes'].ne(self.BD_BLKFINOF['Mes'].shift())
        month_boundaries = self.BD_BLKFINOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])  
        fig.tight_layout()
        plt.savefig('Blanco Fino AS Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['ZN'], 'o-', color='gray', markerfacecolor='yellow', label='Zn')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['Zn_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino ZN', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.ZnErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['ZN'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BLF.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLKFINOF['Mes'].ne(self.BD_BLKFINOF['Mes'].shift())
        month_boundaries = self.BD_BLKFINOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])  
        fig.tight_layout()
        plt.savefig('Blanco Fino ZN Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #FE
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['FE'], 'o-', color='gray', markerfacecolor='yellow', label='Fe')
        line2, = ax.plot(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['FeFin_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Fino FE', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.FeErrorBLKF:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLKFINOF['SAMPLE_ID'], self.BD_BLKFINOF['FE'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_addD(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nFE: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BLF.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLKFINOF['Mes'].ne(self.BD_BLKFINOF['Mes'].shift())
        month_boundaries = self.BD_BLKFINOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])  
        fig.tight_layout()
        plt.savefig('Blanco Fino FE Hist.png', dpi=300, bbox_inches='tight')
        plt.show()

        ## Blanco Grueso
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['CU'], 'o-', color='gray', markerfacecolor='yellow', label='Cu')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Cu_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso CU', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.CuErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['CU'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BG.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLANCOF['Mes'].ne(self.BD_BLANCOF['Mes'].shift())
        month_boundaries = self.BD_BLANCOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])    
        fig.tight_layout()
        plt.savefig('Blanco Grueso CU Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['AG'], 'o-', color='gray', markerfacecolor='yellow', label='Ag')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Ag_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso AG', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.AgErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['AG'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BG.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLANCOF['Mes'].ne(self.BD_BLANCOF['Mes'].shift())
        month_boundaries = self.BD_BLANCOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([]) 
        fig.tight_layout()
        plt.savefig('Blanco Grueso AG Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Mo
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['MO'], 'o-', color='gray', markerfacecolor='yellow', label='Mo')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Mo_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso MO', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.MoErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['MO'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y}')
            
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BG.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLANCOF['Mes'].ne(self.BD_BLANCOF['Mes'].shift())
        month_boundaries = self.BD_BLANCOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([]) 
        fig.tight_layout()
        plt.savefig('Blanco Grueso MO Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #as
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['AS'], 'o-', color='gray', markerfacecolor='yellow', label='As')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['As_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso AS', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.AsErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['AS'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BG.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLANCOF['Mes'].ne(self.BD_BLANCOF['Mes'].shift())
        month_boundaries = self.BD_BLANCOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([]) 
        fig.tight_layout()
        plt.savefig('Blanco Grueso AS Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['ZN'], 'o-', color='gray', markerfacecolor='yellow', label='Zn')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Zn_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso ZN', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.ZnErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['ZN'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BG.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLANCOF['Mes'].ne(self.BD_BLANCOF['Mes'].shift())
        month_boundaries = self.BD_BLANCOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([]) 
        fig.tight_layout()
        plt.savefig('Blanco Grueso ZN Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #FE
        fig, ax = plt.subplots(figsize=(10, 6))
        line1, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['FE'], 'o-', color='gray', markerfacecolor='yellow', label='Fe')
        line2, = ax.plot(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['Fe_MAV'], '-', color='red', label='Max Accepted Value')
        ax.set_title('Blanco Grueso FE', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
        ax.annotate(f"Error: {self.FeErrorBLKG:.2f}%", xy=(0.75, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        scatter_points = ax.scatter(self.BD_BLANCOF['SAMPLE_ID'], self.BD_BLANCOF['FE'], color='yellow')
        cursor = mplcursors.cursor(scatter_points, hover=True)
        @cursor.connect("add")
        def on_addD(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nFE: {y}')
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_BG.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_BLANCOF['Mes'].ne(self.BD_BLANCOF['Mes'].shift())
        month_boundaries = self.BD_BLANCOF[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([]) 
        fig.tight_layout()
        plt.savefig('Blanco Grueso FE Hist.png', dpi=300, bbox_inches='tight')
        plt.show()


        #STD Baja
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['CU'], 'o-', color='blue', label='Cu')
        lines = ['MHCU', 'BVCU', '+2STDCU', '-2STDCU', '-3STDCU', '+3STDCU']
        line_colors = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names = ['Media Historica Cu', 'BV Cu', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines, line_colors, line_names):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['CU'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Baja Ley CHNB-13 Cu', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)

        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDB.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD13F['Mes'].ne(self.BD_STD13F['Mes'].shift())
        month_boundaries = self.BD_STD13F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])         
        ax.annotate(f"Sesgo: {self.STDB_ERROR_CU:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 Cu Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['AG'], 'o-', color='blue', label='Ag')
        lines1 = ['MHAG',  '+2STDAG', '-2STDAG', '-3STDAG', '+3STDAG']
        line_colors1 = ['orange', 'brown', 'brown', 'red', 'red']
        line_names1 = ['Media Historica Ag',  'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines1, line_colors1, line_names1):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['AG'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)             
        ax.set_title('Estandar Baja Ley CHNB-13 Ag', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDB.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD13F['Mes'].ne(self.BD_STD13F['Mes'].shift())
        month_boundaries = self.BD_STD13F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])        
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 Ag Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #AS
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['AS'], 'o-', color='blue', label='As')
        lines2 = ['MHAS',  '+2STDAS', '-2STDAS', '-3STDAS', '+3STDAS']
        line_colors2 = ['orange',  'brown', 'brown', 'red', 'red']
        line_names2 = ['Media Historica As', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines2, line_colors2, line_names2):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['AS'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)             
        ax.set_title('Estandar Baja Ley CHNB-13 As', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDB.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD13F['Mes'].ne(self.BD_STD13F['Mes'].shift())
        month_boundaries = self.BD_STD13F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])        
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 As Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #MO
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['MO'], 'o-', color='blue', label='Mo')
        lines3 = ['MHMO', 'BVMO', '+2STDMO', '-2STDMO', '-3STDMO', '+3STDMO']
        line_colors3 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names3 = ['Media Historica Mo', 'BV Mo', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines3, line_colors3, line_names3):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['MO'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)             
        ax.set_title('Estandar Baja Ley CHNB-13 Mo', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDB.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD13F['Mes'].ne(self.BD_STD13F['Mes'].shift())
        month_boundaries = self.BD_STD13F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])        
        ax.annotate(f"Sesgo: {self.STDB_ERROR_MO:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 Mo Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['ZN'], 'o-', color='blue', label='Zn')
        lines4 = ['MHZN', 'BVZN', '+2STDZN', '-2STDZN', '-3STDZN', '+3STDZN']
        line_colors4 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names4 = ['Media Historica Zn', 'BV Zn', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines4, line_colors4, line_names4):
            ax.plot(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD13F['SAMPLE_ID'], self.BD_STD13F['ZN'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)             
        ax.set_title('Estandar Baja Ley CHNB-13 Zn', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDB.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD13F['Mes'].ne(self.BD_STD13F['Mes'].shift())
        month_boundaries = self.BD_STD13F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])        
        ax.annotate(f"Sesgo: {self.STDB_ERROR_ZN:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Baja Ley CHNB-13 Zn Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
           
        #STD Media
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['CU'], 'o-', color='blue', label='Cu')
        lines = ['MHCU', 'BVCU', '+2STDCU', '-2STDCU', '-3STDCU', '+3STDCU']
        line_colors = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names = ['Media Historica Cu', 'BV Cu', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines, line_colors, line_names):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['CU'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)
        ax.set_title('Estandar Media Ley CHNM-14 Cu', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDM.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD14F['Mes'].ne(self.BD_STD14F['Mes'].shift())
        month_boundaries = self.BD_STD14F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])        
    
        ax.annotate(f"Sesgo: {self.STDM_ERROR_CU:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 Cu Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['AG'], 'o-', color='blue', label='Ag')
        lines1 = ['MHAG', '+2STDAG', '-2STDAG', '-3STDAG', '+3STDAG']
        line_colors1 = ['orange',  'brown', 'brown', 'red', 'red']
        line_names1 = ['Media Historica Ag',  'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines1, line_colors1, line_names1):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['AG'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Media Ley CHNM-14 Ag', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDM.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD14F['Mes'].ne(self.BD_STD14F['Mes'].shift())
        month_boundaries = self.BD_STD14F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])           
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 Ag Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #AS
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['AS'], 'o-', color='blue', label='As')
        lines2 = ['MHAS', 'BVAS', '+2STDAS', '-2STDAS', '-3STDAS', '+3STDAS']
        line_colors2 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names2 = ['Media Historica As', 'BV As', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines2, line_colors2, line_names2):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['AS'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Media Ley CHNM-14 As', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDM.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD14F['Mes'].ne(self.BD_STD14F['Mes'].shift())
        month_boundaries = self.BD_STD14F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])           
        ax.annotate(f"Sesgo: {self.STDM_ERROR_AS:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 As Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #MO
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['MO'], 'o-', color='blue', label='Mo')
        lines3 = ['MHMO', 'BVMO', '+2STDMO', '-2STDMO', '-3STDMO', '+3STDMO']
        line_colors3 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names3 = ['Media Historica Mo', 'BV Mo', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines3, line_colors3, line_names3):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['MO'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Media Ley CHNM-14 Mo', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDM.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD14F['Mes'].ne(self.BD_STD14F['Mes'].shift())
        month_boundaries = self.BD_STD14F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])           
        ax.annotate(f"Sesgo: {self.STDM_ERROR_MO:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 Mo Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['ZN'], 'o-', color='blue', label='Zn')
        lines4 = ['MHZN', 'BVZN', '+2STDZN', '-2STDZN', '-3STDZN', '+3STDZN']
        line_colors4 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names4 = ['Media Historica Zn', 'BV Zn', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines4, line_colors4, line_names4):
            ax.plot(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD14F['SAMPLE_ID'], self.BD_STD14F['ZN'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Media Ley CHNM-14 Zn', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDM.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD14F['Mes'].ne(self.BD_STD14F['Mes'].shift())
        month_boundaries = self.BD_STD14F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])           
        ax.annotate(f"Sesgo: {self.STDM_ERROR_ZN:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Media Ley CHNM-14 Zn Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        # STD A
        #Cu
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['CU'], 'o-', color='blue', label='Cu')
        lines = ['MHCU', 'BVCU', '+2STDCU', '-2STDCU', '-3STDCU', '+3STDCU']
        line_colors = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names = ['Media Historica Cu', 'BV Cu', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines, line_colors, line_names):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['CU'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)        
        ax.set_title('Estandar Alta Ley CHNA-15 Cu', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCU: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDA.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD15F['Mes'].ne(self.BD_STD15F['Mes'].shift())
        month_boundaries = self.BD_STD15F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])          
        ax.annotate(f"Sesgo: {self.STDA_ERROR_CU:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 Cu Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #Ag
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['AG'], 'o-', color='blue', label='Ag')
        lines1 = ['MHAG', 'BVAG', '+2STDAG', '-2STDAG', '-3STDAG', '+3STDAG']
        line_colors1 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names1 = ['Media Historica Ag', 'BV Ag', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines1, line_colors1, line_names1):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['AG'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Alta Ley CHNA-15 Ag', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAG: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDA.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD15F['Mes'].ne(self.BD_STD15F['Mes'].shift())
        month_boundaries = self.BD_STD15F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])         
        ax.annotate(f"Sesgo: {self.STDA_ERROR_AG:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 Ag Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #AS
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['AS'], 'o-', color='blue', label='As')
        lines2 = ['MHAS', 'BVAS', '+2STDAS', '-2STDAS', '-3STDAS', '+3STDAS']
        line_colors2 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names2 = ['Media Historica As', 'BV As', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines2, line_colors2, line_names2):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['AS'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Alta Ley CHNA-15 As', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nAS: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDA.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD15F['Mes'].ne(self.BD_STD15F['Mes'].shift())
        month_boundaries = self.BD_STD15F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])         
        ax.annotate(f"Sesgo: {self.STDA_ERROR_AS:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 As Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #MO
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['MO'], 'o-', color='blue', label='Mo')
        lines3 = ['MHMO', 'BVMO', '+2STDMO', '-2STDMO', '-3STDMO', '+3STDMO']
        line_colors3 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names3 = ['Media Historica Mo', 'BV Mo', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines3, line_colors3, line_names3):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['MO'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Alta Ley CHNA-15 Mo', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nMO: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDA.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD15F['Mes'].ne(self.BD_STD15F['Mes'].shift())
        month_boundaries = self.BD_STD15F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])         
        ax.annotate(f"Sesgo: {self.STDA_ERROR_MO:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 Mo Hist.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        #ZN
        fig, ax = plt.subplots(figsize=(10, 6))
        line, = ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['ZN'], 'o-', color='blue', label='Zn')
        lines4 = ['MHZN', 'BVZN', '+2STDZN', '-2STDZN', '-3STDZN', '+3STDZN']
        line_colors4 = ['orange', 'black', 'brown', 'brown', 'red', 'red']
        line_names4 = ['Media Historica Zn', 'BV Zn', 'MH+2DE', 'MH-2DE', 'MH-3DE', 'MH+3DE']
        
        for line, color, name in zip(lines4, line_colors4, line_names4):
            ax.plot(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F[line], color=color, label=name)
        scatter = ax.scatter(self.BD_STD15F['SAMPLE_ID'], self.BD_STD15F['ZN'], color='blue', alpha=0)
        cursor = mplcursors.cursor(scatter, hover=True)         
        ax.set_title('Estandar Alta Ley CHNA-15 Zn', fontsize=14, pad=20)
        ax.set_ylabel('Values')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
        ax.grid(color='gray', linestyle='-', linewidth=0.5)
        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nZN: {y:.2f}')
            sel.annotation.get_bbox_patch().set_alpha(0.7)
        # Anotaciones estáticas para cada mes en el centro del rango correspondiente
        for month, pos in unique_months_STDA.items():
            ax.annotate(month, xy=(pos, ax.get_ylim()[0]), xytext=(0, -10),
                        textcoords='offset points', ha='center', va='top', fontsize=10, bbox=dict(facecolor='none', edgecolor='none'))
        
        # Marcar solo el inicio de cada mes
        month_changes = self.BD_STD15F['Mes'].ne(self.BD_STD15F['Mes'].shift())
        month_boundaries = self.BD_STD15F[month_changes].groupby('Mes')['SAMPLE_ID'].min()
        
        for pos in month_boundaries:
            ax.axvline(x=pos, color='gray', linestyle='--')
        ax.set_xticklabels([])         
        ax.annotate(f"Sesgo: {self.STDA_ERROR_ZN:.2f}%", xy=(0.8, -0.25), xycoords='axes fraction',
                    fontsize=15, color='red', bbox=dict(boxstyle="round,pad=0.3", edgecolor="black", facecolor="lightgray"))
        ax.set_xticklabels([])
        fig.tight_layout()  
        plt.savefig('Estandar Alta Ley CHNA-15 Zn Hist.png', dpi=300, bbox_inches='tight')
        plt.show()


    def REPORT(self):
       self.error1.clear()
       self.TextImp.clear()
       # blanco Grueso 
       MHCU13 = 0.266184914841849
       MHAG13 = 2.6863779342723
       MHMO13 = 0.00419716312056736
       MHAS13 = 0.0072113022113022
       MHZN13 = 0.0253397590361446
       BVCU13 = 0.266031884057971
       BVMO13 = 0.00423220338983051
       BVZN13 = 0.0263231884057971       
       BVCU13 = 0.266031884057971
       BVMO13 = 0.00423220338983051
       BVZN13 = 0.0263231884057971       
       MHCU15 = 0.662866666666667
       MHAG15 = 10.3921145584725
       MHMO15 = 0.0104511961722488
       MHAS15 = 0.0685093896713616
       MHZN15 = 0.155243309002433       
       BVCU15 = 0.676945762711864
       BVAG15 = 10.140350877193
       BVMO15 = 0.0105706896551724
       BVAS15 = 0.07035
       BVZN15 = 0.149810294117647       
       MHCU14 = 0.427990543735225
       MHAG14 = 4.6302027972028
       MHMO14 = 0.00978690476190473
       MHAS14 = 0.016616091954023
       MHZN14 = 0.0456080760095013
       BVCU14 = 0.420332352941176
       BVMO14 = 0.00996842105263158
       BVAS14 = 0.0162754385964912
       BVZN14 = 0.0453863636363636
       Report_BLKG = pd.DataFrame({
            'Analito' : ['Cu %','Ag gr/tn', 'Mo %', 'As %', 'Zn %', 'Fe %'],
            'N° Muestras' : [len(self.BD_BLANCOF['CU']), len(self.BD_BLANCOF['AG']), len(self.BD_BLANCOF['MO']),len(self.BD_BLANCOF['AS']),len(self.BD_BLANCOF['ZN']),len(self.BD_BLANCOF['FE'])],
            'Fallos' : [self.BD_BLANCOF['Cu_Error'].sum(),self.BD_BLANCOF['Ag_Error'].sum(),self.BD_BLANCOF['Mo_Error'].sum(),self.BD_BLANCOF['As_Error'].sum(),self.BD_BLANCOF['Zn_Error'].sum(),self.BD_BLANCOF['Fe_Error'].sum()],
            '%Fallos' : [self.CuErrorBLKG ,self.AgErrorBLKG,self.MoErrorBLKG,self.AsErrorBLKG,self.ZnErrorBLKG,self.FeErrorBLKG]})
       Report_BLKG['Clase'] = np.where(Report_BLKG['%Fallos'] > 5, 'INACEPTABLE', 'ACEPTABLE')
       
       # blanco Fino
       Report_BLKF = pd.DataFrame({
            'Analito' : ['Cu %','Ag gr/tn', 'Mo %', 'As %', 'Zn %', 'Fe %'],
            'N° Muestras' : [len(self.BD_BLKFINOF['CU']), len(self.BD_BLKFINOF['AG']), len(self.BD_BLKFINOF['MO']),len(self.BD_BLKFINOF['AS']),len(self.BD_BLKFINOF['ZN']),len(self.BD_BLKFINOF['FE'])],
            'Fallos' : [self.BD_BLKFINOF['Cu_Error'].sum(),self.BD_BLKFINOF['Ag_Error'].sum(),self.BD_BLKFINOF['Mo_Error'].sum(),self.BD_BLKFINOF['As_Error'].sum(),self.BD_BLKFINOF['Zn_Error'].sum(),self.BD_BLKFINOF['Fe_Error'].sum()],
            '%Fallos' : [self.CuErrorBLKF ,self.AgErrorBLKF,self.MoErrorBLKF,self.AsErrorBLKF,self.ZnErrorBLKF,self.FeErrorBLKF]})
       Report_BLKF['Clase'] = np.where(Report_BLKF['%Fallos'] > 5, 'INACEPTABLE', 'ACEPTABLE')
       
       #Dup Campo
       Report_Duplicado = pd.DataFrame({
            'Analito' : ['Cu %','Ag gr/tn', 'Mo %', 'As %', 'Zn %', 'Fe %'],
            'N° Muestras DC' : [len(self.BD_OrigF['CU']), len(self.BD_OrigF['AG']), len(self.BD_OrigF['MO']),len(self.BD_OrigF['AS']),len(self.BD_OrigF['ZN']),len(self.BD_OrigF['FE'])],
            '%Fallos DC' : [self.CuErrorDUPC ,self.AgErrorDUPC,self.MoErrorDUPC,self.AsErrorDUPC,self.ZnErrorDUPC,self.FeErrorDUPC],
            'Clase DC' : ['','', '', '', '', ''],
            'N° Muestras RG' : [len(self.BD_RECHAF['CU']), len(self.BD_RECHAF['AG']), len(self.BD_RECHAF['MO']),len(self.BD_RECHAF['AS']),len(self.BD_RECHAF['ZN']),len(self.BD_RECHAF['FE'])],
            '%Fallos RG' : [self.CuErrorRG ,self.AgErrorRG,self.MoErrorRG,self.AsErrorRG,self.ZnErrorRG,self.FeErrorRG],
            'Clase RG' : ['','', '', '', '', ''],
            'N° Muestras DP' : [len(self.BD_PULPAF['CU']), len(self.BD_PULPAF['AG']), len(self.BD_PULPAF['MO']),len(self.BD_PULPAF['AS']),len(self.BD_PULPAF['ZN']),len(self.BD_PULPAF['FE'])],
            '%Fallos DP' : [self.CuErrorPL ,self.AgErrorPL,self.MoErrorPL,self.AsErrorPL,self.ZnErrorPL,self.FeErrorPL],
            'Clase DP' : ['','', '', '', '', '']})
       Report_Duplicado['Clase DC'] = np.where(Report_Duplicado['%Fallos DC'] > 10, 'INACEPTABLE', 'ACEPTABLE')
       Report_Duplicado['Clase RG'] = np.where(Report_Duplicado['%Fallos RG'] > 10, 'INACEPTABLE', 'ACEPTABLE')
       Report_Duplicado['Clase DP'] = np.where(Report_Duplicado['%Fallos DP'] > 10, 'INACEPTABLE', 'ACEPTABLE')
       
       #EstNDARES
       Report_Estandares = pd.DataFrame({
            'Nombre\nSTD' : ['CHNB13','CHNB13', 'CHNB13', 'CHNB13', 'CHNB13','CHNM14','CHNM14','CHNM14','CHNM14','CHNM14','CHNA15','CHNA15','CHNA15','CHNA15','CHNA15'],
            'Analito' : ['Cu %','Ag gr/tn', 'Mo %', 'As %', 'Zn %','Cu %','Ag gr/tn', 'Mo %', 'As %', 'Zn %','Cu %','Ag gr/tn', 'Mo %', 'As %', 'Zn %'],
            'N° Muestras' : [len(self.BD_STD13F['CU']), len(self.BD_STD13F['AG']), len(self.BD_STD13F['MO']),len(self.BD_STD13F['AS']),len(self.BD_STD13F['ZN']), len(self.BD_STD14F['CU']), len(self.BD_STD14F['AG']), len(self.BD_STD14F['MO']),len(self.BD_STD14F['AS']),len(self.BD_STD14F['ZN']), len(self.BD_STD15F['CU']), len(self.BD_STD15F['AG']), len(self.BD_STD15F['MO']),len(self.BD_STD15F['AS']),len(self.BD_STD15F['ZN'])],
            'Media\nActual\n(MA)' : [self.BD_STD13F['CU'].mean(),self.BD_STD13F['AG'].mean(),self.BD_STD13F['MO'].mean(),self.BD_STD13F['AS'].mean(),self.BD_STD13F['ZN'].mean(),self.BD_STD14F['CU'].mean(),self.BD_STD14F['AG'].mean(),self.BD_STD14F['MO'].mean(),self.BD_STD14F['AS'].mean(),self.BD_STD14F['ZN'].mean(),self.BD_STD15F['CU'].mean(),self.BD_STD15F['AG'].mean(),self.BD_STD15F['MO'].mean(),self.BD_STD15F['AS'].mean(),self.BD_STD15F['ZN'].mean()],
            'Media\nHistorica\n(MH)' : [MHCU13,MHAG13,MHMO13,MHAS13,MHZN13,MHCU14, MHAG14,MHMO14,MHAS14,MHZN14,MHCU15,MHAG15,MHMO15,MHAS15,MHZN15],
            'Media\nCertificada\n(BV)' : [BVCU13,99,BVMO13, 99,BVZN13,BVCU14,99,BVMO14,BVAS14,BVZN14, BVCU15,BVAG15,BVMO15,BVAS15,BVZN15 ],
            'Lim. de\nControl\nMH-2STD' : [self.BD_STD13F['-2STDCU'].mean(),self.BD_STD13F['-2STDAG'].mean(),self.BD_STD13F['-2STDMO'].mean(),self.BD_STD13F['-2STDAS'].mean(),self.BD_STD13F['-2STDZN'].mean(), self.BD_STD14F['-2STDCU'].mean(),self.BD_STD14F['-2STDAG'].mean(),self.BD_STD14F['-2STDMO'].mean(),self.BD_STD14F['-2STDAS'].mean(),self.BD_STD14F['-2STDZN'].mean(), self.BD_STD15F['-2STDCU'].mean(),self.BD_STD15F['-2STDAG'].mean(),self.BD_STD15F['-2STDMO'].mean(),self.BD_STD15F['-2STDAS'].mean(),self.BD_STD15F['-2STDZN'].mean()],
            'Lim. de\nControl\nMH+2STD' : [self.BD_STD13F['+2STDCU'].mean(),self.BD_STD13F['+2STDAG'].mean(),self.BD_STD13F['+2STDMO'].mean(),self.BD_STD13F['+2STDAS'].mean(),self.BD_STD13F['+2STDZN'].mean(), self.BD_STD14F['+2STDCU'].mean(),self.BD_STD14F['+2STDAG'].mean(),self.BD_STD14F['+2STDMO'].mean(),self.BD_STD14F['+2STDAS'].mean(),self.BD_STD14F['+2STDZN'].mean(), self.BD_STD15F['+2STDCU'].mean(),self.BD_STD15F['+2STDAG'].mean(),self.BD_STD15F['+2STDMO'].mean(),self.BD_STD15F['+2STDAS'].mean(),self.BD_STD15F['+2STDZN'].mean()],
            'Sesgo(%)' : [self.STDB_ERROR_CU,99,self.STDB_ERROR_MO,99,self.STDB_ERROR_ZN,self.STDM_ERROR_CU,99,self.STDM_ERROR_MO,self.STDM_ERROR_AS,self.STDM_ERROR_ZN,self.STDA_ERROR_CU, self.STDA_ERROR_AG,self.STDA_ERROR_MO,self.STDA_ERROR_AS,self.STDA_ERROR_ZN],
            'Clase' : ['','No Certificado', '', 'No Certificado', '', '', 'No Certificado','','','','','','','',''],
            '# Fallidos': [self.BD_STD13F['Cu_Out'].sum(), self.BD_STD13F['Ag_Out'].sum(), self.BD_STD13F['Mo_Out'].sum(), self.BD_STD13F['As_Out'].sum(), self.BD_STD13F['Zn_Out'].sum(), self.BD_STD14F['Cu_Out'].sum(), self.BD_STD14F['Ag_Out'].sum(), self.BD_STD14F['Mo_Out'].sum(), self.BD_STD14F['As_Out'].sum(), self.BD_STD14F['Zn_Out'].sum(), self.BD_STD15F['Cu_Out'].sum(), self.BD_STD15F['Ag_Out'].sum(), self.BD_STD15F['Mo_Out'].sum(), self.BD_STD15F['As_Out'].sum(), self.BD_STD15F['Zn_Out'].sum()],
            '% Fallidos': [self.FSTD13CU,self.FSTD13AG,self.FSTD13MO,self.FSTD13AS,self.FSTD13ZN,self.FSTD14CU,self.FSTD14AG,self.FSTD14MO,self.FSTD14AS,self.FSTD14ZN,self.FSTD15CU,self.FSTD15AG,self.FSTD15MO,self.FSTD15AS,self.FSTD15ZN]})
       
       conditions = [
        (Report_Estandares['Sesgo(%)'] < 5) & (Report_Estandares['Sesgo(%)'] > -5),
        (Report_Estandares['Sesgo(%)'] >= -10) & (Report_Estandares['Sesgo(%)'] < -5),
        (Report_Estandares['Sesgo(%)'] > 5) & (Report_Estandares['Sesgo(%)'] <= 10),
        (Report_Estandares['Sesgo(%)'] == 99)]

       values = ['ACEPTABLE', 'CUESTIONABLE', 'CUESTIONABLE', 'No Certificado']
       Report_Estandares['Clase'] = np.select(conditions, values, default='INACEPTABLE')
       numcolstd = Report_Estandares.select_dtypes(include=['float', 'int']).columns
       Report_Estandares[numcolstd] = Report_Estandares[numcolstd].round(3)
       Report_Estandares['Sesgo(%)'] = Report_Estandares['Sesgo(%)'].replace(99, '')
       Report_Estandares['Media\nCertificada\n(BV)'] = Report_Estandares['Media\nCertificada\n(BV)'].replace(99, '')

       numcoldup = Report_Duplicado.select_dtypes(include=['float', 'int']).columns
       Report_Duplicado[numcoldup] = Report_Duplicado[numcoldup].round(3)    
       
       numcolBF = Report_BLKF.select_dtypes(include=['float', 'int']).columns
       Report_BLKF[numcolBF] = Report_BLKF[numcolBF].round(3)  
       
       numcolBG = Report_BLKG.select_dtypes(include=['float', 'int']).columns
       Report_BLKG[numcolBG] = Report_BLKG[numcolBG].round(3)  
       
       #EstNDARESm 2
       Report_Estandares2 = pd.DataFrame({
            'Analito' : ['Cu %','Cu %', 'Cu %', 'Ag gr/tn', 'Ag gr/tn','Ag gr/tn','Mo %','Mo %','Mo %','As %','As %','As %','Zn %','Zn %','Zn %'],
            'Nombre\nSTD' : ['CHNB13','CHNM14', 'CHNA15', 'CHNB13','CHNM14', 'CHNA15','CHNB13','CHNM14', 'CHNA15','CHNB13','CHNM14', 'CHNA15','CHNB13','CHNM14', 'CHNA15',],
            'N° Muestras' : [len(self.BD_STD13F['CU']),  len(self.BD_STD14F['CU']), len(self.BD_STD15F['CU']),  len(self.BD_STD13F['AG']), len(self.BD_STD14F['AG']), len(self.BD_STD15F['AG']) , len(self.BD_STD13F['MO']), len(self.BD_STD14F['MO']) , len(self.BD_STD15F['MO']),len(self.BD_STD13F['AS']),len(self.BD_STD14F['AS']),len(self.BD_STD15F['AS']) ,len(self.BD_STD13F['ZN']),len(self.BD_STD14F['ZN']),len(self.BD_STD15F['ZN'])],
            'Media\nActual\n(MA)' : [self.BD_STD13F['CU'].mean(),self.BD_STD14F['CU'].mean(),self.BD_STD15F['CU'].mean() ,self.BD_STD13F['AG'].mean(),self.BD_STD14F['AG'].mean(),self.BD_STD15F['AG'].mean(),self.BD_STD13F['MO'].mean(),self.BD_STD14F['MO'].mean(),self.BD_STD15F['MO'].mean(),self.BD_STD13F['AS'].mean(),self.BD_STD14F['AS'].mean(),self.BD_STD15F['AS'].mean(),self.BD_STD13F['ZN'].mean(),self.BD_STD14F['ZN'].mean(),self.BD_STD15F['ZN'].mean()],
            'Media\nHistorica\n(MH)' : [MHCU13,MHCU14,MHCU15,MHAG13, MHAG14,MHAG15 , MHMO13,MHMO14,MHMO15 ,MHAS13,MHAS14,MHAS15,MHZN13,MHZN14,MHZN15],
            'Media\nCertificada\n(BV)' : [BVCU13,BVCU14, BVCU15,99, 99,BVAG15,BVMO13,BVMO14,BVMO15,99,BVAS14,BVAS15,BVZN13,BVZN14,BVZN15 ],
            'Lim. de\nControl\nMH-2STD' : [self.BD_STD13F['-2STDCU'].mean(), self.BD_STD14F['-2STDCU'].mean(), self.BD_STD15F['-2STDCU'].mean(),self.BD_STD13F['-2STDAG'].mean(),self.BD_STD14F['-2STDAG'].mean(),self.BD_STD15F['-2STDAG'].mean(),self.BD_STD13F['-2STDMO'].mean(),self.BD_STD14F['-2STDMO'].mean(),self.BD_STD15F['-2STDMO'].mean(),self.BD_STD13F['-2STDAS'].mean(),self.BD_STD14F['-2STDAS'].mean(),self.BD_STD15F['-2STDAS'].mean(),self.BD_STD13F['-2STDZN'].mean(),self.BD_STD14F['-2STDZN'].mean(),self.BD_STD15F['-2STDZN'].mean()],
            'Lim. de\nControl\nMH+2STD' : [self.BD_STD13F['+2STDCU'].mean(), self.BD_STD14F['+2STDCU'].mean(), self.BD_STD15F['+2STDCU'].mean(),self.BD_STD13F['+2STDAG'].mean(),self.BD_STD14F['+2STDAG'].mean(),self.BD_STD15F['+2STDAG'].mean(),self.BD_STD13F['+2STDMO'].mean(),self.BD_STD14F['+2STDMO'].mean(),self.BD_STD15F['+2STDMO'].mean(),self.BD_STD13F['+2STDAS'].mean(),self.BD_STD14F['+2STDAS'].mean(),self.BD_STD15F['+2STDAS'].mean(),self.BD_STD13F['+2STDZN'].mean(),self.BD_STD14F['+2STDZN'].mean(),self.BD_STD15F['+2STDZN'].mean()],
            'Sesgo(%)' : [self.STDB_ERROR_CU,self.STDM_ERROR_CU,self.STDA_ERROR_CU,99,99, self.STDA_ERROR_AG,self.STDB_ERROR_MO,self.STDM_ERROR_MO,self.STDA_ERROR_MO,99,self.STDM_ERROR_AS,self.STDA_ERROR_AS,self.STDB_ERROR_ZN,self.STDM_ERROR_ZN,self.STDA_ERROR_ZN],
            'Clase' : ['', '', '','No Certificado', 'No Certificado', '','','','', 'No Certificado','','','','',''],
            '# Fallidos': [self.BD_STD13F['Cu_Out'].sum(), self.BD_STD14F['Cu_Out'].sum(), self.BD_STD15F['Cu_Out'].sum(), self.BD_STD13F['Ag_Out'].sum(), self.BD_STD14F['Ag_Out'].sum(), self.BD_STD15F['Ag_Out'].sum(), self.BD_STD13F['Mo_Out'].sum(), self.BD_STD14F['Mo_Out'].sum(), self.BD_STD15F['Mo_Out'].sum(), self.BD_STD13F['As_Out'].sum(), self.BD_STD14F['As_Out'].sum(), self.BD_STD15F['As_Out'].sum(), self.BD_STD13F['Zn_Out'].sum(), self.BD_STD14F['Zn_Out'].sum(), self.BD_STD15F['Zn_Out'].sum()],
            '% Fallidos': [self.FSTD13CU,self.FSTD14CU,self.FSTD15CU,self.FSTD13AG,self.FSTD14AG,self.FSTD15AG,self.FSTD13MO,self.FSTD14MO,self.FSTD15MO,self.FSTD13AS,self.FSTD14AS,self.FSTD15AS,self.FSTD13ZN,self.FSTD14ZN,self.FSTD15ZN]})      
       conditionsB = [
        (Report_Estandares2['Sesgo(%)'] < 5) & (Report_Estandares2['Sesgo(%)'] > -5),
        (Report_Estandares2['Sesgo(%)'] >= -10) & (Report_Estandares2['Sesgo(%)'] < -5),
        (Report_Estandares2['Sesgo(%)'] > 5) & (Report_Estandares2['Sesgo(%)'] <= 10),
        (Report_Estandares2['Sesgo(%)'] == 99)]

       valuesB = ['ACEPTABLE', 'CUESTIONABLE', 'CUESTIONABLE', 'No Certificado']
       Report_Estandares2['Clase'] = np.select(conditionsB, valuesB, default='INACEPTABLE')
       numcolstdB = Report_Estandares2.select_dtypes(include=['float', 'int']).columns
       Report_Estandares2[numcolstdB] = Report_Estandares2[numcolstdB].round(3)
       Report_Estandares2['Sesgo(%)'] = Report_Estandares2['Sesgo(%)'].replace(99, '')
       Report_Estandares2['Media\nCertificada\n(BV)'] = Report_Estandares2['Media\nCertificada\n(BV)'].replace(99, '')


       
       ErrorRechazoGrueso = pd.DataFrame(columns=['SAMPLE_ORIG', 'LAB_REPORT', 'OBS'])
       error_columns = ['Cu_Error', 'Ag_Error', 'Mo_Error', 'As_Error', 'Zn_Error', 'Fe_Error']
       obs_values = ['Cu', 'Ag', 'Mo', 'As', 'Zn', 'Fe']
       for error_column, obs_value in zip(error_columns, obs_values):
           error_rows = self.BD_RECHAF[self.BD_RECHAF[error_column] == 1][['SAMPLE_ORIG', 'LAB_REPORT']]
           error_rows = error_rows.copy()  
           error_rows['OBS'] = obs_value
           ErrorRechazoGrueso = pd.concat([ErrorRechazoGrueso, error_rows], ignore_index=True)
           
       ErrorRechazoGrueso = ErrorRechazoGrueso.sort_values('SAMPLE_ORIG')
           
       ErrorDuplicadoPulpa = pd.DataFrame(columns=['SAMPLE_ORIG', 'LAB_REPORT', 'OBS'])
       error_columns = ['Cu_Error', 'Ag_Error', 'Mo_Error', 'As_Error', 'Zn_Error', 'Fe_Error']
       obs_values = ['Cu', 'Ag', 'Mo', 'As', 'Zn', 'Fe']
       for error_column, obs_value in zip(error_columns, obs_values):
           error_rows = self.BD_PULPAF[self.BD_PULPAF[error_column] == 1][['SAMPLE_ORIG', 'LAB_REPORT']]
           error_rows = error_rows.copy()  
           error_rows['OBS'] = obs_value
           ErrorDuplicadoPulpa = pd.concat([ErrorDuplicadoPulpa, error_rows], ignore_index=True)
       ErrorDuplicadoPulpa = ErrorDuplicadoPulpa.sort_values('SAMPLE_ORIG')    
           
       ErrorDuplicadoCampo = pd.DataFrame(columns=['SAMPLE_ID', 'LAB_REPORT', 'OBS'])
       error_columns = ['Cu_Error', 'Ag_Error', 'Mo_Error', 'As_Error', 'Zn_Error', 'Fe_Error']
       obs_values = ['Cu', 'Ag', 'Mo', 'As', 'Zn', 'Fe']
       for error_column, obs_value in zip(error_columns, obs_values):
           error_rows = self.BD_PULPAF[self.BD_PULPAF[error_column] == 1][['SAMPLE_ID', 'LAB_REPORT']]
           error_rows = error_rows.copy()  
           error_rows['OBS'] = obs_value
           ErrorDuplicadoCampo = pd.concat([ErrorDuplicadoCampo, error_rows], ignore_index=True)
       ErrorDuplicadoCampo = ErrorDuplicadoCampo.sort_values('SAMPLE_ID')
       
       ErrorBlancoGrueso = pd.DataFrame(columns=['SAMPLE_ID', 'LAB_REPORT', 'OBS'])
       error_columns = ['Cu_Error', 'Ag_Error', 'Mo_Error', 'As_Error', 'Zn_Error', 'Fe_Error']
       obs_values = ['Cu', 'Ag', 'Mo', 'As', 'Zn', 'Fe']
       for error_column, obs_value in zip(error_columns, obs_values):
           error_rows = self.BD_BLANCOF[self.BD_BLANCOF[error_column] == 1][['SAMPLE_ID', 'LAB_REPORT']]
           error_rows = error_rows.copy()  
           error_rows['OBS'] = obs_value
           ErrorBlancoGrueso = pd.concat([ErrorBlancoGrueso, error_rows], ignore_index=True)
       ErrorBlancoGrueso = ErrorBlancoGrueso.sort_values('SAMPLE_ID')    
           
       ErrorBlancoFino = pd.DataFrame(columns=['SAMPLE_ID', 'LAB_REPORT', 'OBS'])
       error_columns = ['Cu_Error', 'Ag_Error', 'Mo_Error', 'As_Error', 'Zn_Error', 'Fe_Error']
       obs_values = ['Cu', 'Ag', 'Mo', 'As', 'Zn', 'Fe']
       for error_column, obs_value in zip(error_columns, obs_values):
           error_rows = self.BD_BLKFINOF[self.BD_BLKFINOF[error_column] == 1][['SAMPLE_ID', 'LAB_REPORT']]
           error_rows = error_rows.copy()  
           error_rows['OBS'] = obs_value
           ErrorBlancoFino = pd.concat([ErrorBlancoFino, error_rows], ignore_index=True)
       ErrorBlancoFino = ErrorBlancoFino.sort_values('SAMPLE_ID')
       
       ErrorCHNB = pd.DataFrame(columns=['SAMPLE_ID', 'LAB_REPORT', 'OBS'])
       error_columns = ['Cu_Out', 'Ag_Out', 'Mo_Out', 'As_Out', 'Zn_Out']
       obs_values = ['Cu', 'Ag', 'Mo', 'As', 'Zn']
       for error_column, obs_value in zip(error_columns, obs_values):
           error_rows = self.BD_STD13F[self.BD_STD13F[error_column] == 1][['SAMPLE_ID', 'LAB_REPORT']]
           error_rows = error_rows.copy()  
           error_rows['OBS'] = obs_value
           ErrorCHNB = pd.concat([ErrorCHNB, error_rows], ignore_index=True)
       ErrorCHNB = ErrorCHNB.sort_values('SAMPLE_ID')    
           
       ErrorCHNM = pd.DataFrame(columns=['SAMPLE_ID', 'LAB_REPORT', 'OBS'])
       error_columns = ['Cu_Out', 'Ag_Out', 'Mo_Out', 'As_Out', 'Zn_Out']
       obs_values = ['Cu', 'Ag', 'Mo', 'As', 'Zn']
       for error_column, obs_value in zip(error_columns, obs_values):
           error_rows = self.BD_STD14F[self.BD_STD14F[error_column] == 1][['SAMPLE_ID', 'LAB_REPORT']]
           error_rows = error_rows.copy()  
           error_rows['OBS'] = obs_value
           ErrorCHNM = pd.concat([ErrorCHNM, error_rows], ignore_index=True)
       ErrorCHNM = ErrorCHNM.sort_values('SAMPLE_ID')    
           
       ErrorCHNA = pd.DataFrame(columns=['SAMPLE_ID', 'LAB_REPORT', 'OBS'])
       error_columns = ['Cu_Out', 'Ag_Out', 'Mo_Out', 'As_Out', 'Zn_Out']
       obs_values = ['Cu', 'Ag', 'Mo', 'As', 'Zn']
       for error_column, obs_value in zip(error_columns, obs_values):
           error_rows = self.BD_STD15F[self.BD_STD15F[error_column] == 1][['SAMPLE_ID', 'LAB_REPORT']]
           error_rows = error_rows.copy()  
           error_rows['OBS'] = obs_value
           ErrorCHNA = pd.concat([ErrorCHNA, error_rows], ignore_index=True)
       ErrorCHNA = ErrorCHNA.sort_values('SAMPLE_ID')     

       self.fillDC(Report_Duplicado)
       self.fillSTD(Report_Estandares)
       self.fillSTD2(Report_Estandares2)
       self.fillBKF(Report_BLKF)
       self.fillBKG(Report_BLKG)
       self.ERG(ErrorRechazoGrueso)
       self.EDP(ErrorDuplicadoPulpa)
       self.EDC(ErrorDuplicadoCampo)
       self.EBG(ErrorBlancoGrueso)
       self.EBF(ErrorBlancoFino)
       self.ESTB(ErrorCHNB)
       self.ESTM(ErrorCHNM)
       self.ESTA(ErrorCHNA)
       
               
       dataframes = {
           "ErrorRechazoGrueso": ErrorRechazoGrueso,
           "ErrorDuplicadoPulpa": ErrorDuplicadoPulpa,
           "ErrorDuplicadoCampo": ErrorDuplicadoCampo,
           "ErrorBlancoGrueso": ErrorBlancoGrueso,
           "ErrorBlancoFino": ErrorBlancoFino,
           "ErrorCHNB": ErrorCHNB,
           "ErrorCHNM": ErrorCHNM,
           "ErrorCHNA": ErrorCHNA}        
       for name, df in dataframes.items():
           if 'SAMPLE_ORIG' in df.columns:
               df.rename(columns={'SAMPLE_ORIG': 'SAMPLE_ID'}, inplace=True)
           df['Ubicacion Error '] = name
       
       ErrorTotal = pd.concat(dataframes.values(), ignore_index=True)
       try:
           
           dataframesB = {
               "Duplicado": Report_Duplicado,
               "Estandares": Report_Estandares,
               "Blanco Fino": Report_BLKF,
               "Blanco Grueso": Report_BLKG,
               "ErrorRechazoGrueso": ErrorRechazoGrueso,
               "ErrorDuplicadoPulpa": ErrorDuplicadoPulpa,
               "ErrorDuplicadoCampo": ErrorDuplicadoCampo,
               "ErrorBlancoGrueso": ErrorBlancoGrueso,
               "ErrorBlancoFino": ErrorBlancoFino,
               "ErrorCHNB": ErrorCHNB,
               "ErrorCHNM": ErrorCHNM,
               "ErrorCHNA": ErrorCHNA,
               "TotalError": ErrorTotal}    
           filename = "Reporte_QAQC.xlsx"
           with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
               for sheet_name, df in dataframesB.items():
                   df.to_excel(writer, sheet_name=sheet_name, index=False)
           
           self.TextImp.setText('Reporte Generado y Exportado, ingresar en las pestañas y/o revisar carpeta')
       except Exception as e:
           mensaje = f"Error: {e}. Revisar que el archivo Reporte_QAQC.xlsx no este abierto"
           QMessageBox.critical(self, "No se creo el reporte", mensaje)
           self.error1.setText('Error al exportar reportel')
       

        

    def fillDC(self, df):
        self.DFDUP.setRowCount(df.shape[0])
        self.DFDUP.setColumnCount(df.shape[1])
        self.DFDUP.setHorizontalHeaderLabels(df.columns)
        
        columns_to_color = ['Clase RG', 'Clase DC', 'Clase DP']
        
        for j in range(df.shape[1]):
            header_item = self.DFDUP.horizontalHeaderItem(j)
            if header_item is not None:
                header_item.setBackground(QColor('lightgray'))
        
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                if df.columns[j] in columns_to_color:
                    if df.iat[i, j] == 'ACEPTABLE':
                        item.setForeground(QColor('green'))
                    elif df.iat[i, j] == 'INACEPTABLE':
                        item.setForeground(QColor('red'))
                
                if j == 0:
                    item.setBackground(QColor('lightgray'))
                
                self.DFDUP.setItem(i, j, item)
        
        self.DFDUP.resizeColumnsToContents()
        self.DFDUP.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: lightgray; }")

        
        
    def ESTA(self, df):
        self.E_STDA.setRowCount(df.shape[0])
        self.E_STDA.setColumnCount(df.shape[1])
        self.E_STDA.setHorizontalHeaderLabels(df.columns)
    
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                self.E_STDA.setItem(i, j, item)
    
        self.E_STDA.resizeColumnsToContents()
        
        
    def ESTM(self, df):
        self.E_STDM.setRowCount(df.shape[0])
        self.E_STDM.setColumnCount(df.shape[1])
        self.E_STDM.setHorizontalHeaderLabels(df.columns)
    
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                self.E_STDM.setItem(i, j, item)
    
        self.E_STDM.resizeColumnsToContents()
        
    def ESTB(self, df):
        self.E_STDB.setRowCount(df.shape[0])
        self.E_STDB.setColumnCount(df.shape[1])
        self.E_STDB.setHorizontalHeaderLabels(df.columns)
    
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                self.E_STDB.setItem(i, j, item)
    
        self.E_STDB.resizeColumnsToContents()  
        
    def EBF(self, df):
        self.ErrorBlancoF.setRowCount(df.shape[0])
        self.ErrorBlancoF.setColumnCount(df.shape[1])
        self.ErrorBlancoF.setHorizontalHeaderLabels(df.columns)
    
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                self.ErrorBlancoF.setItem(i, j, item)
    
        self.ErrorBlancoF.resizeColumnsToContents()
    def EBG(self, df):
        self.ErrorBlancoG.setRowCount(df.shape[0])
        self.ErrorBlancoG.setColumnCount(df.shape[1])
        self.ErrorBlancoG.setHorizontalHeaderLabels(df.columns)
    
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                self.ErrorBlancoG.setItem(i, j, item)
    
        self.ErrorBlancoG.resizeColumnsToContents()
    def ERG(self, df):
        self.E_RG.setRowCount(df.shape[0])
        self.E_RG.setColumnCount(df.shape[1])
        self.E_RG.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                self.E_RG.setItem(i, j, item)

        self.E_RG.resizeColumnsToContents()

    def EDP(self, df):
        self.E_PULP.setRowCount(df.shape[0])
        self.E_PULP.setColumnCount(df.shape[1])
        self.E_PULP.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                self.E_PULP.setItem(i, j, item)

        self.E_PULP.resizeColumnsToContents()
        
    def EDC(self, df):
        self.E_DUPC.setRowCount(df.shape[0])
        self.E_DUPC.setColumnCount(df.shape[1])
        self.E_DUPC.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                self.E_DUPC.setItem(i, j, item)

        self.E_DUPC.resizeColumnsToContents()

    def fillSTD(self, df):
        self.DFSTD.setRowCount(df.shape[0])
        self.DFSTD.setColumnCount(df.shape[1])
        self.DFSTD.setHorizontalHeaderLabels(df.columns)
    
        columns_to_color = ['Clase']
    
        # Aplicar color de fondo a los encabezados
        for j in range(df.shape[1]):
            header_item = self.DFSTD.horizontalHeaderItem(j)
            if header_item is not None:
                header_item.setBackground(QColor('lightgray'))
    
        # Aplicar colores a las celdas según el valor
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[i, j]))
    
                # Colorear celdas de la primera columna según el valor
                if j == 0:
                    if df.iat[i, j] == 'CHNB13':
                        item.setBackground(QColor('yellow'))
                    elif df.iat[i, j] == 'CHNM14':
                        item.setBackground(QColor('orange'))
                    elif df.iat[i, j] == 'CHNA15':
                        item.setBackground(QColor('green'))
                    else:
                        item.setBackground(QColor('lightgray'))  # Color por defecto
    
                # Colorear celdas según la columna 'Clase'
                if df.columns[j] in columns_to_color:
                    if df.iat[i, j] == 'ACEPTABLE':
                        item.setForeground(QColor('green'))
                    elif df.iat[i, j] == 'CUESTIONABLE':
                        item.setForeground(QColor('orange'))
                    elif df.iat[i, j] == 'INACEPTABLE':
                        item.setForeground(QColor('red'))
                    elif df.iat[i, j] == 'No Certificado':
                        item.setForeground(QColor('black'))
    
                self.DFSTD.setItem(i, j, item)
    
        self.DFSTD.resizeColumnsToContents()
        self.DFSTD.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: lightgray; }")
        
    def fillSTD2(self, df):
        self.DFSTDB.setRowCount(df.shape[0])
        self.DFSTDB.setColumnCount(df.shape[1])
        self.DFSTDB.setHorizontalHeaderLabels(df.columns)
        
        columns_to_color = ['Clase']
        
        # Aplicar color de fondo a los encabezados
        for j in range(df.shape[1]):
            header_item = self.DFSTDB.horizontalHeaderItem(j)
            if header_item is not None:
                header_item.setBackground(QColor('lightgray'))
        
        # Aplicar colores a las celdas según el valor
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[i, j]))
                
                # Colorear celdas de la primera columna con fondo gris
                if j == 0:
                    item.setBackground(QColor('lightgray'))
        
                # Colorear celdas de la segunda columna según el valor
                if j == 1:
                    if df.iat[i, j] == 'CHNB13':
                        item.setBackground(QColor('yellow'))
                    elif df.iat[i, j] == 'CHNM14':
                        item.setBackground(QColor('orange'))
                    elif df.iat[i, j] == 'CHNA15':
                        item.setBackground(QColor('green'))
                    else:
                        item.setBackground(QColor('lightgray'))  # Color por defecto
        
                # Colorear celdas según la columna 'Clase'
                if df.columns[j] in columns_to_color:
                    if df.iat[i, j] == 'ACEPTABLE':
                        item.setForeground(QColor('green'))
                    elif df.iat[i, j] == 'CUESTIONABLE':
                        item.setForeground(QColor('orange'))
                    elif df.iat[i, j] == 'INACEPTABLE':
                        item.setForeground(QColor('red'))
                    elif df.iat[i, j] == 'No Certificado':
                        item.setForeground(QColor('black'))
        
                self.DFSTDB.setItem(i, j, item)
        
        self.DFSTDB.resizeColumnsToContents()
        self.DFSTDB.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: lightgray; }")




    def fillBKF(self, df):
        self.DFBLF.setRowCount(df.shape[0])
        self.DFBLF.setColumnCount(df.shape[1])
        self.DFBLF.setHorizontalHeaderLabels(df.columns)
        
        columns_to_color = ['Clase']
        
        for j in range(df.shape[1]):
            header_item = self.DFBLF.horizontalHeaderItem(j)
            if header_item is not None:
                header_item.setBackground(QColor('lightgray'))
        
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                if df.columns[j] in columns_to_color:
                    if df.iat[i, j] == 'ACEPTABLE':
                        item.setForeground(QColor('green'))
                    elif df.iat[i, j] == 'CUESTIONABLE':
                        item.setForeground(QColor('orange'))
                    elif df.iat[i, j] == 'INACEPTABLE':
                        item.setForeground(QColor('red'))
                    elif df.iat[i, j] == 'No Certificado':
                        item.setForeground(QColor('black'))
                
                if j == 0:
                    item.setBackground(QColor('lightgray'))
                
                self.DFBLF.setItem(i, j, item)
        
        self.DFBLF.resizeColumnsToContents()
        self.DFBLF.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: lightgray; }")
    
    def fillBKG(self, df):
        self.DFBLG.setRowCount(df.shape[0])
        self.DFBLG.setColumnCount(df.shape[1])
        self.DFBLG.setHorizontalHeaderLabels(df.columns)
        
        columns_to_color = ['Clase']
        
        for j in range(df.shape[1]):
            header_item = self.DFBLG.horizontalHeaderItem(j)
            if header_item is not None:
                header_item.setBackground(QColor('lightgray'))
        
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                if df.columns[j] in columns_to_color:
                    if df.iat[i, j] == 'ACEPTABLE':
                        item.setForeground(QColor('green'))
                    elif df.iat[i, j] == 'CUESTIONABLE':
                        item.setForeground(QColor('orange'))
                    elif df.iat[i, j] == 'INACEPTABLE':
                        item.setForeground(QColor('red'))
                    elif df.iat[i, j] == 'No Certificado':
                        item.setForeground(QColor('black'))
                
                if j == 0:
                    item.setBackground(QColor('lightgray'))
                
                self.DFBLG.setItem(i, j, item)
        
        self.DFBLG.resizeColumnsToContents()
        self.DFBLG.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: lightgray; }")


    def toggle_text(self):
        if self.aboutlabel.isHidden():
            self.aboutlabel.setText("Desarrollado por: Felix Recines Trinidad")
            self.aboutlabel.show()
        else:
            self.aboutlabel.clear()
            self.aboutlabel.hide()
            
    def ImpBDRepExt(self):
        self.TextImp2.clear()
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Reporte de Muestras QAQC", "", "Excel Files (*.xlsx; *.xlsm)")
        if file_path:
            try:
                self.BD_cert = pd.read_excel(file_path, header=9)            
                
                if 'G0095MCR1' in self.BD_cert.columns:
                    self.BD_cert['G0172MCR1'] = self.BD_cert['G0095MCR1'].combine_first(self.BD_cert['G0172MCR1'])
                    self.BD_cert.drop(columns=['G0095MCR1'], inplace=True)
                if 'G0040MCR1' in self.BD_cert.columns:
                    self.BD_cert['G0125MCR1'] = self.BD_cert['G0040MCR1'].combine_first(self.BD_cert['G0125MCR1'])
                    self.BD_cert.drop(columns=['G0095MCR1'], inplace=True)
                columns_to_drop = ['G0465MCR1', 'G0781MCR1', 'G0732MCR1']
                existing_columns = [col for col in columns_to_drop if col in self.BD_cert.columns]
                if existing_columns:
                    self.BD_cert.drop(columns=existing_columns, inplace=True)
                self.BD_cert.columns = self.BD_cert.iloc[0]
                self.BD_cert = self.BD_cert[4:]  
                
                self.TextImp2.setText("Archivo Reporte de Muestras Externas Importado")
                self.ProcRMA.setEnabled(True)
                
            except Exception as e:
                mensaje = f"Error: {e}. Revisar si el archivo seleccionado es el correcto o fue modificado e intentelo nuevamente"
                QMessageBox.critical(self, "Error al cargar el archivo", mensaje)
                
    def ProcRMAExt(self):
        try:
            self.TextImp2.clear()
            Elem_RMA = {'Cu':'Cu_ct', 'Ag':'Ag_ct', 'Mo':'Mo_ct', 'As':'As_ct', 'Zn':'Zn_ct', 'CuSolH+':'CuSS_ct', 'CuScitrico':'CuCi_ct', 'SAMPLE':'SAMPLE_ID'}
            self.BD_cert = self.BD_cert.rename(columns=Elem_RMA)
            self.BD_cert.reset_index(drop=True, inplace=True)
            self.BD_cert['SAMPLE_ID'] = pd.to_numeric(self.BD_cert['SAMPLE_ID'], errors='coerce')
            self.BD_cert = self.BD_cert[self.BD_cert['SAMPLE_ID'].notna()]
            self.BD_cert['SAMPLE_ID'] = self.BD_cert['SAMPLE_ID'].astype(int)
            self.BD_cert =self.BD_cert[(self.BD_cert['SAMPLE_ID'] >= self.MinValue) & (self.BD_cert['SAMPLE_ID'] <= self.MaxValue)]
            ValorRMA={
            '<0.001':0.0005,
            '<0.5':0.25,
            '<0.0005':0.00025,
            '<0.003':0.0015,
            '<0.002':0.001,
            '>2.5':2.6,
            '>300':301,
            '>1':1.1,
            '>2.5':2.6,
            '>10':10.1,
            '>2.500':2.6}
            self.BD_cert = self.BD_cert.replace(ValorRMA)
            colsConvert = ['Cu_ct','Ag_ct','Mo_ct', 'As_ct', 'Zn_ct','CuSS_ct', 'CuCi_ct']
            self.BD_cert[colsConvert] = self.BD_cert[colsConvert].astype(float) 
            print(self.BD_cert.dtypes) 
            self.BD_RMA = pd.merge(self.BD_cert, self.BD_Total[['CU', 'AG', 'MO', 'ARS', 'ZN', 'FE','CUOX','SAMPLE_ID']], on='SAMPLE_ID', how='left')
            self.BD_RMA['CU_ER'] = 2*(self.BD_RMA['Cu_ct'] - self.BD_RMA['CU']) / (self.BD_RMA['Cu_ct'] + self.BD_RMA['CU']) 
            self.BD_RMA['AG_ER'] = 2*(self.BD_RMA['Ag_ct'] - self.BD_RMA['AG']) / (self.BD_RMA['Ag_ct'] + self.BD_RMA['AG']) 
            self.BD_RMA['MO_ER'] = 2*(self.BD_RMA['Mo_ct'] - self.BD_RMA['MO']) / (self.BD_RMA['Mo_ct'] + self.BD_RMA['MO']) 
            self.BD_RMA['AS_ER'] = 2*(self.BD_RMA['As_ct'] - self.BD_RMA['ARS']) / (self.BD_RMA['As_ct'] +self.BD_RMA['ARS']) 
            self.BD_RMA['ZN_ER'] = 2*(self.BD_RMA['Zn_ct'] - self.BD_RMA['ZN']) / (self.BD_RMA['Zn_ct'] + self.BD_RMA['ZN'])
            self.BD_RMA['CU_ER'] = self.BD_RMA['CU_ER'].apply(lambda x: x if x > 0.1 or x < -0.1 else None)
            self.BD_RMA['AG_ER'] = self.BD_RMA['AG_ER'].apply(lambda x: x if x > 0.1 or x < -0.1 else None)
            self.BD_RMA['MO_ER'] = self.BD_RMA['MO_ER'].apply(lambda x: x if x > 0.1 or x < -0.1 else None)
            self.BD_RMA['AS_ER'] = self.BD_RMA['AS_ER'].apply(lambda x: x if x > 0.1 or x < -0.1 else None)
            self.BD_RMA['ZN_ER'] = self.BD_RMA['ZN_ER'].apply(lambda x: x if x > 0.1 or x < -0.1 else None)
            self.SesgoCu = (1-(np.std(self.BD_RMA['Cu_ct'])/np.std(self.BD_RMA['CU'])))*100
            self.SesgoAg = (1-(np.std(self.BD_RMA['Ag_ct'])/np.std(self.BD_RMA['AG'])))*100
            self.SesgoMo = (1-(np.std(self.BD_RMA['Mo_ct'])/np.std(self.BD_RMA['MO'])))*100
            self.SesgoAs = (1-(np.std(self.BD_RMA['As_ct'])/np.std(self.BD_RMA['ARS'])))*100
            self.SesgoZn = (1-(np.std(self.BD_RMA['Zn_ct'])/np.std(self.BD_RMA['ZN'])))*100
            self.BD_RMA_SS = self.BD_RMA.dropna(subset=['CUOX'])
            self.BD_RMA_SS = self.BD_RMA_SS.dropna(subset=['CuSS_ct'])
            self.SesgoCuss = (1-(np.std(self.BD_RMA_SS['CuSS_ct'])/np.std(self.BD_RMA_SS['CUOX'])))*100
            self.BD_RMA['CuProm'] = (self.BD_RMA['Cu_ct']+self.BD_RMA['CU'])/2
            self.BD_RMA['AgProm'] = (self.BD_RMA['Ag_ct']+self.BD_RMA['AG'])/2
            self.BD_RMA['MoProm'] = (self.BD_RMA['Mo_ct']+self.BD_RMA['MO'])/2
            self.BD_RMA['AsProm'] = (self.BD_RMA['As_ct']+self.BD_RMA['ARS'])/2
            self.BD_RMA['ZnProm'] = (self.BD_RMA['Zn_ct']+self.BD_RMA['ZN'])/2
            self.BD_RMA_SS['CuSSProm'] = (self.BD_RMA_SS['CUOX']+self.BD_RMA_SS['CuSS_ct'])/2
            self.BD_RMA['CuRD'] = 200*(self.BD_RMA['CU']-self.BD_RMA['Cu_ct'])/(self.BD_RMA['CU']+self.BD_RMA['Cu_ct'])
            self.BD_RMA['AgRD'] = 200*(self.BD_RMA['AG']-self.BD_RMA['Ag_ct'])/(self.BD_RMA['AG']+self.BD_RMA['Ag_ct'])
            self.BD_RMA['MoRD'] = 200*(self.BD_RMA['MO']-self.BD_RMA['Mo_ct'])/(self.BD_RMA['MO']+self.BD_RMA['Mo_ct'])
            self.BD_RMA['AsRD'] = 200*(self.BD_RMA['ARS']-self.BD_RMA['As_ct'])/(self.BD_RMA['ARS']+self.BD_RMA['As_ct'])
            self.BD_RMA['ZnRD'] = 200*(self.BD_RMA['ZN']-self.BD_RMA['Zn_ct'])/(self.BD_RMA['ZN']+self.BD_RMA['Zn_ct'])
            self.BD_RMA.to_csv("RMA2.csv")
            self.BD_RMA_SS['CuSSRD'] = 200*(self.BD_RMA_SS['CUOX']-self.BD_RMA_SS['CuSS_ct'])/(self.BD_RMA_SS['CUOX']+self.BD_RMA_SS['CuSS_ct'])
            self.BD_RMA_SS['CUSS_ER'] = 2*(self.BD_RMA_SS['CuSS_ct'] - self.BD_RMA_SS['CUOX']) / (self.BD_RMA_SS['CuSS_ct'] + self.BD_RMA_SS['CUOX'])
            self.BD_RMA_SS['CUSS_ER'] = self.BD_RMA_SS['CUSS_ER'].apply(lambda x: x if x > 0.1 or x < -0.1 else None)
            self.BD_RMA = self.BD_RMA.sort_values('CuProm')
            self.BD_RMA['MediaMovCuProm'] = self.BD_RMA['CuProm'].rolling(window=2).mean()
            self.BD_RMA = self.BD_RMA.sort_values('AgProm')
            self.BD_RMA['MediaMovAgProm'] = self.BD_RMA['AgProm'].rolling(window=2).mean()
            self.BD_RMA = self.BD_RMA.sort_values('MoProm')
            self.BD_RMA['MediaMovMoProm'] = self.BD_RMA['MoProm'].rolling(window=2).mean()
            self.BD_RMA = self.BD_RMA.sort_values('AsProm')
            self.BD_RMA['MediaMovAsProm'] = self.BD_RMA['AsProm'].rolling(window=2).mean()
            self.BD_RMA = self.BD_RMA.sort_values('ZnProm')
            self.BD_RMA['MediaMovZnProm'] = self.BD_RMA['ZnProm'].rolling(window=2).mean()
            self.BD_RMA = self.BD_RMA.sort_values('SAMPLE_ID')
            self.BD_RMA_SS = self.BD_RMA_SS.sort_values('CuSSProm')
            self.BD_RMA_SS['MediaMovCuSSProm'] = self.BD_RMA_SS['CuSSProm'].rolling(window=2).mean()
            self.BD_RMA_SS = self.BD_RMA_SS.sort_values('SAMPLE_ID')
            self.BD_RMA.to_csv("RMA3.csv")
            self.TextImp2.setText("Creacion de Columnas y Variables Completado")
            self.GRMA.setEnabled(True)
            self.QQPRMA.setEnabled(True)
            self.RDRMA.setEnabled(True)
            self.CUSSL.setEnabled(True)
            self.REPRMA.setEnabled(True)
        except Exception as e:
            mensaje = f"Error: {e}. Revisar si el archivo seleccionado tiene los analitos completos"
            QMessageBox.critical(self, "Error en el proceso", mensaje)
    def RMAPlot(self):
        try:
            #Cu
            self.BD_RMA = self.BD_RMA.dropna(subset=['CU'])
            X = sm.add_constant(self.BD_RMA['CU'])  
            modelo = sm.OLS(self.BD_RMA['Cu_ct'], X)
            resultados = modelo.fit()
            fig, ax = plt.subplots(figsize=(9, 8))
            scatter = ax.scatter(self.BD_RMA['CU'], self.BD_RMA['Cu_ct'], color='#339966', edgecolor='black', s=30,  marker='^',label='Check\nSamples')
            ax.plot(self.BD_RMA['CU'], resultados.predict(X), color='red', label='RMA\nTrend')
            ax.plot([0, self.BD_RMA['CU'].max()], [0, self.BD_RMA['CU'].max()], color='#5089BC', label='Y=X')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            
                
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id:.0f}\nCu SGS: {x:.3f}\nCu Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
                        
            ax.set_title(
                'External Controls (CHINALCO SGS vs CERTIMIN)', 
                fontsize=14, 
                pad=20,
                bbox=dict(facecolor='#FFFF99', edgecolor='black', boxstyle='round,pad=0.5')
            )
            ax.set_xlabel('CHINALCO SGS Cu%')
            ax.set_ylabel('CERTIMIN Cu%')
            ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), edgecolor="black")
            ax.annotate(
                f"Sesgo",
                xy=(1.03, 0.39),
                xycoords='axes fraction',
                fontsize=15,
                color='black',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.annotate(
                f"{self.SesgoCu:.2f}%",
                xy=(1.03, 0.35),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('CU RMA PLOT.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Ag
            self.BD_RMA = self.BD_RMA.dropna(subset=['AG'])
            X = sm.add_constant(self.BD_RMA['AG'])  
            modelo = sm.OLS(self.BD_RMA['Ag_ct'], X)
            resultados = modelo.fit()
            fig, ax = plt.subplots(figsize=(9, 8))
            scatter = ax.scatter(self.BD_RMA['AG'], self.BD_RMA['Ag_ct'], color='#339966', edgecolor='black', s=30,  marker='^',label='Check\nSamples')
            ax.plot(self.BD_RMA['AG'], resultados.predict(X), color='red', label='RMA\nTrend')
            ax.plot([0, self.BD_RMA['AG'].max()], [0, self.BD_RMA['AG'].max()], color='#5089BC', label='Y=X')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id:.0f}\nAg SGS: {x:.3f}\nAg Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title(
                'External Controls (CHINALCO SGS vs CERTIMIN)', 
                fontsize=14, 
                pad=20,
                bbox=dict(facecolor='#FFFF99', edgecolor='black', boxstyle='round,pad=0.5')
            )
            ax.set_xlabel('CHINALCO SGS Ag g/t')
            ax.set_ylabel('CERTIMIN Ag g/t')
            ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), edgecolor="black")
            ax.annotate(
                f"Sesgo",
                xy=(1.03, 0.39),
                xycoords='axes fraction',
                fontsize=15,
                color='black',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.annotate(
                f"{self.SesgoAg:.2f}%",
                xy=(1.03, 0.35),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('AG RMA PLOT.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #AS
            self.BD_RMA = self.BD_RMA.dropna(subset=['ARS'])
            X = sm.add_constant(self.BD_RMA['ARS'])  
            modelo = sm.OLS(self.BD_RMA['As_ct'], X)
            resultados = modelo.fit()
            fig, ax = plt.subplots(figsize=(9, 8))
            scatter = ax.scatter(self.BD_RMA['ARS'], self.BD_RMA['As_ct'], color='#339966', edgecolor='black', s=30,  marker='^',label='Check\nSamples')
            ax.plot(self.BD_RMA['ARS'], resultados.predict(X), color='red', label='RMA\nTrend')
            ax.plot([0, self.BD_RMA['ARS'].max()], [0, self.BD_RMA['ARS'].max()], color='#5089BC', label='Y=X')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id:.0f}\nAs SGS: {x:.3f}\nAs Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
                
            ax.set_title(
                'External Controls (CHINALCO SGS vs CERTIMIN)', 
                fontsize=14, 
                pad=20,
                bbox=dict(facecolor='#FFFF99', edgecolor='black', boxstyle='round,pad=0.5')
            )
            ax.set_xlabel('CHINALCO SGS As%')
            ax.set_ylabel('CERTIMIN As%')
            ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), edgecolor="black")
            ax.annotate(
                f"Sesgo",
                xy=(1.03, 0.39),
                xycoords='axes fraction',
                fontsize=15,
                color='black',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.annotate(
                f"{self.SesgoAs:.2f}%",
                xy=(1.03, 0.35),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('AS RMA PLOT.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #MO
            self.BD_RMA = self.BD_RMA.dropna(subset=['MO'])
            X = sm.add_constant(self.BD_RMA['MO'])  
            modelo = sm.OLS(self.BD_RMA['Mo_ct'], X)
            resultados = modelo.fit()
            fig, ax = plt.subplots(figsize=(9, 8))
            scatter = ax.scatter(self.BD_RMA['MO'], self.BD_RMA['Mo_ct'], color='#339966', edgecolor='black', s=30,  marker='^',label='Check\nSamples')
            ax.plot(self.BD_RMA['MO'], resultados.predict(X), color='red', label='RMA\nTrend')
            ax.plot([0, self.BD_RMA['MO'].max()], [0, self.BD_RMA['MO'].max()], color='#5089BC', label='Y=X')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id:.0f}\nMo SGS: {x:.3f}\nMo Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title(
                'External Controls (CHINALCO SGS vs CERTIMIN)', 
                fontsize=14, 
                pad=20,
                bbox=dict(facecolor='#FFFF99', edgecolor='black', boxstyle='round,pad=0.5')
            )
            ax.set_xlabel('CHINALCO SGS Mo%')
            ax.set_ylabel('CERTIMIN Mo%')
            ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), edgecolor="black")
            ax.annotate(
                f"Sesgo",
                xy=(1.03, 0.39),
                xycoords='axes fraction',
                fontsize=15,
                color='black',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.annotate(
                f"{self.SesgoMo:.2f}%",
                xy=(1.03, 0.35),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('MO RMA PLOT.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            # Zn
            self.BD_RMA = self.BD_RMA.dropna(subset=['ZN'])
            X = sm.add_constant(self.BD_RMA['ZN'])  
            modelo = sm.OLS(self.BD_RMA['Zn_ct'], X)
            resultados = modelo.fit()
            fig, ax = plt.subplots(figsize=(9, 8))
            scatter = ax.scatter(self.BD_RMA['ZN'], self.BD_RMA['Zn_ct'], color='#339966', edgecolor='black', s=30,  marker='^',label='Check\nSamples')
            ax.plot(self.BD_RMA['ZN'], resultados.predict(X), color='red', label='RMA\nTrend')
            ax.plot([0, self.BD_RMA['ZN'].max()], [0, self.BD_RMA['ZN'].max()], color='#5089BC', label='Y=X')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id:.0f}\nZn SGS: {x:.3f}\nZn Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title(
                'External Controls (CHINALCO SGS vs CERTIMIN)', 
                fontsize=14, 
                pad=20,
                bbox=dict(facecolor='#FFFF99', edgecolor='black', boxstyle='round,pad=0.5')
            )
            ax.set_xlabel('CHINALCO SGS Zn%')
            ax.set_ylabel('CERTIMIN Zn%')
            ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), edgecolor="black")
            ax.annotate(
                f"Sesgo",
                xy=(1.03, 0.39),
                xycoords='axes fraction',
                fontsize=15,
                color='black',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.annotate(
                f"{self.SesgoZn:.2f}%",
                xy=(1.03, 0.35),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('ZN RMA PLOT.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #CUSS
            self.BD_RMA_SS = self.BD_RMA_SS.dropna(subset=['CUOX'])
            X = sm.add_constant(self.BD_RMA_SS['CUOX'])  
            modelo = sm.OLS(self.BD_RMA_SS['CuSS_ct'], X)
            resultados = modelo.fit()
            fig, ax = plt.subplots(figsize=(9, 8))
            scatter = ax.scatter(self.BD_RMA_SS['CUOX'], self.BD_RMA_SS['CuSS_ct'], color='#339966', edgecolor='black', s=30,  marker='^',label='Check\nSamples')
            ax.plot(self.BD_RMA_SS['CUOX'], resultados.predict(X), color='red', label='RMA\nTrend')
            ax.plot([0, self.BD_RMA_SS['CUOX'].max()], [0, self.BD_RMA_SS['CUOX'].max()], color='#5089BC', label='Y=X')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA_SS.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id:.0f}\nCuSS SGS: {x:.3f}\nCuSS Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title(
                'External Controls (CHINALCO SGS vs CERTIMIN)', 
                fontsize=14, 
                pad=20,
                bbox=dict(facecolor='#FFFF99', edgecolor='black', boxstyle='round,pad=0.5')
            )
            ax.set_xlabel('CHINALCO SGS CuSS%')
            ax.set_ylabel('CERTIMIN CuSS%')
            ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), edgecolor="black")
            ax.annotate(
                f"Sesgo",
                xy=(1.03, 0.39),
                xycoords='axes fraction',
                fontsize=15,
                color='black',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.annotate(
                f"{self.SesgoCuss:.2f}%",
                xy=(1.03, 0.35),
                xycoords='axes fraction',
                fontsize=15,
                color='red',
                fontweight='bold',
                bbox=dict(boxstyle="round,pad=0.3", edgecolor="white", facecolor="white")
            )
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('CUSS RMA PLOT.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar los archivos importados"
            QMessageBox.critical(self, "Error Inesperado", mensaje)
            
    def QQPLOTRMA(self):
        try:
            #Cu
            quant_CUCt = np.percentile(self.BD_RMA['Cu_ct'], np.linspace(0, 100, 101))
            quant_CUSGS = np.percentile(self.BD_RMA['CU'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quant_CUSGS, quant_CUSGS, color='red', label='Y=X')
            scatter = ax.scatter(quant_CUSGS, quant_CUCt, color='blue', marker='x', label='CU Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Cu SGS: {x:.3f}\nCu Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('CU% QQ-Plot', fontsize=14, pad=20)
            ax.set_xlabel('SGS Toromocho')
            ax.set_ylabel('Certimin')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('CU% QQ-Plot RMA.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Ag
            quant_AGCt = np.percentile(self.BD_RMA['Ag_ct'], np.linspace(0, 100, 101))
            quant_AGSGS = np.percentile(self.BD_RMA['AG'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quant_AGSGS, quant_AGSGS, color='red', label='Y=X')
            scatter = ax.scatter(quant_AGSGS, quant_AGCt, color='blue', marker='x', label='AG Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Ag SGS: {x:.3f}\nAg Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('AG% QQ-Plot', fontsize=14, pad=20)
            ax.set_xlabel('SGS Toromocho')
            ax.set_ylabel('Certimin')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot AG RMA.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #mo
            quant_MOCt = np.percentile(self.BD_RMA['Mo_ct'], np.linspace(0, 100, 101))
            quant_MOSGS = np.percentile(self.BD_RMA['MO'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quant_AGSGS, quant_AGSGS, color='red', label='Y=X')
            scatter = ax.scatter(quant_AGSGS, quant_AGCt, color='blue', marker='x', label='MO Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Mo SGS: {x:.3f}\nMo Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('MO% QQ-Plot', fontsize=14, pad=20)
            ax.set_xlabel('SGS Toromocho')
            ax.set_ylabel('Certimin')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot MO RMA.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #As
            quant_ASCt = np.percentile(self.BD_RMA['As_ct'], np.linspace(0, 100, 101))
            quant_ASSGS = np.percentile(self.BD_RMA['ARS'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quant_ASSGS, quant_ASSGS, color='red', label='Y=X')
            scatter = ax.scatter(quant_ASSGS, quant_ASCt, color='blue', marker='x', label='AS Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'As SGS: {x:.3f}\nAs Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('AS% QQ-Plot', fontsize=14, pad=20)
            ax.set_xlabel('SGS Toromocho')
            ax.set_ylabel('Certimin')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot AS RMA.png', dpi=300, bbox_inches='tight')
            plt.show()
                    
            #Zn 
            quant_ZNCt = np.percentile(self.BD_RMA['Zn_ct'], np.linspace(0, 100, 101))
            quant_ZNSGS = np.percentile(self.BD_RMA['ZN'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quant_ZNSGS, quant_ZNSGS, color='red', label='Y=X')
            scatter = ax.scatter(quant_ZNSGS, quant_ZNCt, color='blue', marker='x', label='ZN Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'Zn SGS: {x:.3f}\nZn Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('ZN% QQ-Plot', fontsize=14, pad=20)
            ax.set_xlabel('SGS Toromocho')
            ax.set_ylabel('Certimin')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot ZN RMA.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            #Cuox
            quant_CUSSCt = np.percentile(self.BD_RMA_SS['CuSS_ct'], np.linspace(0, 100, 101))
            quant_CUSSSGS = np.percentile(self.BD_RMA_SS['CUOX'], np.linspace(0, 100, 101))
            fig, ax = plt.subplots(figsize=(8, 8))
            ax.plot(quant_CUSSSGS, quant_CUSSSGS, color='red', label='Y=X')
            scatter = ax.scatter(quant_CUSSSGS, quant_CUSSCt, color='blue', marker='x', label='CuSS Q-Q Plot')
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'CuSS SGS: {x:.3f}\nCuSS Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('CuSS% QQ-Plot', fontsize=14, pad=20)
            ax.set_xlabel('SGS Toromocho')
            ax.set_ylabel('Certimin')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('Q-Q Plot CuSS RMA.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar los archivos importados"
            QMessageBox.critical(self, "Error inesperado", mensaje)
            
    def RDPLOTRMA(self):
        try:
            #Cu
            self.BD_RMA = self.BD_RMA.sort_values('CuProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RMA['MediaMovCuProm'], self.BD_RMA['CuRD'], color='blue', label='RD Plot Cu')
            ax.plot(self.BD_RMA['CuProm'], self.BD_RMA['CuRD'], color='red', label='Media Movil')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCu_prom: {x:.3f}\nRD Cu: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RMA Cu RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RMA Cu RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RMA = self.BD_RMA.sort_values('SAMPLE_ID')
            
            #Ag
            self.BD_RMA = self.BD_RMA.sort_values('AgProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RMA['MediaMovAgProm'], self.BD_RMA['AgRD'], color='blue', label='RD Plot Ag')
            ax.plot(self.BD_RMA['AgProm'], self.BD_RMA['AgRD'], color='red', label='Media Movil')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAg_prom: {x:.3f}\nRD Ag: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RMA Ag RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RMA Ag RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RMA = self.BD_RMA.sort_values('SAMPLE_ID')
            
            #Mo
            self.BD_RMA = self.BD_RMA.sort_values('MoProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RMA['MediaMovMoProm'], self.BD_RMA['MoRD'], color='blue', label='RD Plot Mo')
            ax.plot(self.BD_RMA['MoProm'], self.BD_RMA['MoRD'], color='red', label='Media Movil')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nMo_prom: {x:.3f}\nRD Mo: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RMA Mo RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RMA Mo RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RMA = self.BD_RMA.sort_values('SAMPLE_ID')
            
            #As
            self.BD_RMA = self.BD_RMA.sort_values('AsProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RMA['MediaMovAsProm'], self.BD_RMA['AsRD'], color='blue', label='RD Plot As')
            ax.plot(self.BD_RMA['AsProm'], self.BD_RMA['AsRD'], color='red', label='Media Movil')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nAs_prom: {x:.3f}\nRD As: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RMA As RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RMA As RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RMA = self.BD_RMA.sort_values('SAMPLE_ID')
            
            #Zn
            self.BD_RMA = self.BD_RMA.sort_values('ZnProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RMA['MediaMovZnProm'], self.BD_RMA['ZnRD'], color='blue', label='RD Plot Zn')
            ax.plot(self.BD_RMA['ZnProm'], self.BD_RMA['ZnRD'], color='red', label='Media Movil')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nZn_prom: {x:.3f}\nRD Zn: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RMA Zn RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RMA Zn RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RMA = self.BD_RMA.sort_values('SAMPLE_ID')
            
            #Cuox
            self.BD_RMA_SS = self.BD_RMA_SS.sort_values('CuSSProm')
            fig, ax = plt.subplots(figsize=(10, 8))
            scatter1 = ax.scatter(self.BD_RMA_SS['MediaMovCuSSProm'], self.BD_RMA_SS['CuSSRD'], color='blue', label='RD Plot Mo')
            ax.plot(self.BD_RMA_SS['CuSSProm'], self.BD_RMA_SS['CuSSRD'], color='red', label='Media Movil')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                index = sel.index
                sample_id = self.BD_RMA_SS.iloc[index]['SAMPLE_ID']
                sel.annotation.set(text=f'SAMPLE_ID: {sample_id}\nCuSS_prom: {x:.3f}\nRD CuSS: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            ax.set_title('RMA CuSS RD-Plot', fontsize=14, pad=20)
            ax.set_xlabel('Promedio')
            ax.set_ylabel('% RD')
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)
            ax.grid(color='gray', linestyle='-', linewidth=0.5)
            fig.tight_layout()
            plt.savefig('RMA CuSS RD-Plot.png', dpi=300, bbox_inches='tight')
            plt.show()
            self.BD_RMA_SS = self.BD_RMA_SS.sort_values('SAMPLE_ID')
        except Exception as e:
            mensaje = f"Error: {e}. Revisar los archivos importados"
            QMessageBox.critical(self, "Error inesperado", mensaje)
            
    def CUSSPLOT(self):
        try:
            #Lineas CUSS CUCI
            fig, ax = plt.subplots(figsize=(12, 8))
            line1, = ax.plot(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CuSS_ct'], color='orange', marker='o', label='CuSS% Certimin')
            line2, = ax.plot(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CuCi_ct'], color='red', marker='o', label='CuCi% Certimin')
            line3, = ax.plot(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CUOX'], color='blue', marker='o', label='CuSS% SGS Toromocho')        
            scatter1 = ax.scatter(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CuSS_ct'], color='orange')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCuSS_Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)        
            scatter2 = ax.scatter(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CuCi_ct'], color='red')
            cursor1 = mplcursors.cursor(scatter2, hover=True)
            @cursor1.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCuCi_Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)        
            scatter3 = ax.scatter(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CUOX'], color='blue')
            cursor2 = mplcursors.cursor(scatter3, hover=True)
            @cursor2.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCuSS_Toro: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)            
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.07), shadow=True, ncol=3)        
            ax.set_title('CuCi vs CUSS')
            ax.set_xlabel('Muestras')
            ax.set_ylabel('%Cu Valores')        
            ax.grid(True)
            plt.savefig('LinePlot CUSS CUCI.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            # Linea CUSS InterLab
            fig, ax = plt.subplots(figsize=(12, 8))
            line1, = ax.plot(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CuSS_ct'], color='orange', marker='o', label='CuSS% Certimin')
            line3, = ax.plot(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CUOX'], color='blue', marker='o', label='CuSS% SGS Toromocho')        
            scatter1 = ax.scatter(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CuSS_ct'], color='orange')
            cursor = mplcursors.cursor(scatter1, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCuSS_Cert: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)
            scatter3 = ax.scatter(self.BD_RMA_SS['SAMPLE_ID'], self.BD_RMA_SS['CUOX'], color='blue')
            cursor2 = mplcursors.cursor(scatter3, hover=True)
            @cursor2.connect("add")
            def on_add(sel):
                x, y = sel.target
                sel.annotation.set(text=f'SAMPLE_ID: {x:.0f}\nCuSS_Toro: {y:.3f}')
                sel.annotation.get_bbox_patch().set_alpha(0.7)            
            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.07), shadow=True, ncol=3)        
            ax.set_title('CUSS CERTIMIN vs SGS Toromocho')
            ax.set_xlabel('Muestras')
            ax.set_ylabel('%Cu Valores')        
            ax.grid(True)
            plt.savefig('LinePlot CUSS SGS vs CERT.png', dpi=300, bbox_inches='tight')
            plt.show()
        except Exception as e:
            mensaje = f"Error: {e}. Revisar los archivos importados"
            QMessageBox.critical(self, "Error inesperado", mensaje)
            
    def REPORTDFRMA(self):
        self.GRI.setEnabled(True)
        self.BD_RMA.to_csv("RMAentrada.csv")
        Report_RMA = pd.DataFrame({ 'Analito': ['Cu %', 'Ag gr/tn', 'Mo %', 'As %', 'Zn %', 'CuSS %'],
        'RMA\n# Pares': [
            self.BD_RMA['CU'].count(),
            self.BD_RMA['AG'].count(),
            self.BD_RMA['MO'].count(),
            self.BD_RMA['ARS'].count(),
            self.BD_RMA['ZN'].count(),
            self.BD_RMA_SS['CuSS_ct'].count()
        ],
        'RMA\nSesgo (%)' : [self.SesgoCu, self.SesgoAg, self.SesgoMo, self.SesgoAs, self.SesgoZn, self.SesgoCuss],
        'RMA\nClase' : ['','','','','',''],
        'Promedio\nSGS-MCP': [
            self.BD_RMA['CU'].mean(),
            self.BD_RMA['AG'].mean(),
            self.BD_RMA['MO'].mean(),
            self.BD_RMA['ARS'].mean(),
            self.BD_RMA['ZN'].mean(),
            self.BD_RMA_SS['CUOX'].mean()
        ],
        'Promedio\nCertimin': [
            self.BD_RMA['Cu_ct'].mean(),
            self.BD_RMA['Ag_ct'].mean(),
            self.BD_RMA['Mo_ct'].mean(),
            self.BD_RMA['As_ct'].mean(),
            self.BD_RMA['Zn_ct'].mean(),
            self.BD_RMA_SS['CuSS_ct'].mean()
        ],
        'Certimin\ndentro 5%\nSGS' : ['','','','','',''],
        'Outlier (ER>0.1)\n#' : [
            self.BD_RMA['CU_ER'].count(),
            self.BD_RMA['AG_ER'].count(),
            self.BD_RMA['MO_ER'].count(),
            self.BD_RMA['AS_ER'].count(),
            self.BD_RMA['ZN_ER'].count(),
            self.BD_RMA_SS['CUSS_ER'].count()
        ],
        'Outlier (ER>0.1)\nPorcentaje %' : [
            self.BD_RMA['CU_ER'].count()*100/self.BD_RMA['CU'].count(),
            self.BD_RMA['AG_ER'].count()*100/self.BD_RMA['AG'].count(),
            self.BD_RMA['MO_ER'].count()*100/self.BD_RMA['MO'].count(),
            self.BD_RMA['AS_ER'].count()*100/self.BD_RMA['ARS'].count(),
            self.BD_RMA['ZN_ER'].count()*100/self.BD_RMA['ZN'].count(),
            self.BD_RMA_SS['CUSS_ER'].count()*100/self.BD_RMA_SS['CuSS_ct'].count()
        ],
        'Outlier (ER>0.1)\Clase' : ['','','','','',''],
        'Clase\nGeneral' : ['','','','','','']})
        
        conditionsRMA = [
        (Report_RMA['RMA\nSesgo (%)'] < 5) & (Report_RMA['RMA\nSesgo (%)'] > -5),
        (Report_RMA['RMA\nSesgo (%)'] >= -10) & (Report_RMA['RMA\nSesgo (%)'] < -5),
        (Report_RMA['RMA\nSesgo (%)'] > 5) & (Report_RMA['RMA\nSesgo (%)'] <= 10),
        (Report_RMA['RMA\nSesgo (%)'] == 99)
    ]
    
        valuesRMA = ['ACEPTABLE', 'CUESTIONABLE', 'CUESTIONABLE', 'No Certificado']
        Report_RMA['RMA\nClase'] = np.select(conditionsRMA, valuesRMA, default='INACEPTABLE')
        conditionsRMAP = [
        (Report_RMA['Promedio\nCertimin'] < (Report_RMA['Promedio\nSGS-MCP'] + Report_RMA['Promedio\nSGS-MCP'] * 0.05)) &
        (Report_RMA['Promedio\nCertimin'] > (Report_RMA['Promedio\nSGS-MCP'] - Report_RMA['Promedio\nSGS-MCP'] * 0.05))
    ]
        valuesRMAP = ['ACEPTABLE']
        Report_RMA['Certimin\ndentro 5%\nSGS'] = np.select(conditionsRMAP, valuesRMAP, default='INACEPTABLE')
        Report_RMA['Outlier (ER>0.1)\Clase'] = np.where(Report_RMA['Outlier (ER>0.1)\nPorcentaje %'] > 10, 'INACEPTABLE', 'ACEPTABLE')
        def determinar_clase_general(row):
            inaceptable_count = (row == 'INACEPTABLE').sum()
            aceptable_count = (row == 'ACEPTABLE').sum()
            
            if inaceptable_count >= 2:
                return 'INACEPTABLE'
            elif aceptable_count >= 2:
                return 'ACEPTABLE'
            else:
                return 'INACEPTABLE'

        Report_RMA['Clase\nGeneral'] = Report_RMA.apply(determinar_clase_general, axis=1)
        numcol = Report_RMA.select_dtypes(include=['float', 'int']).columns
        Report_RMA[numcol] = Report_RMA[numcol].round(3)  
        self.fillRMA(Report_RMA)
        self.BD_RMA.to_csv("RMAsalida.csv")
        try:
            
            dataframesB = {
                "RMA SGS-CERT": Report_RMA}    
            filename = "Reporte_QAQC_Ext.xlsx"
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                for sheet_name, df in dataframesB.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            self.TextImp.setText('Reporte Generado y Exportado, ingresar en las pestañas y/o revisar carpeta')
        except Exception as e:
            mensaje = f"Error: {e}. Revisar que el archivo Reporte_QAQC_Ext.xlsx no este abierto"
            QMessageBox.critical(self, "No se creo el reporte", mensaje)
        

        
    
        
        
    def fillRMA(self, df):
        self.DFRMA.setRowCount(df.shape[0])
        self.DFRMA.setColumnCount(df.shape[1])
        self.DFRMA.setHorizontalHeaderLabels(df.columns)
        
        columns_to_color = ['RMA\nClase', 'Certimin\ndentro 5%\nSGS', 'Outlier (ER>0.1)\Clase', 'Clase\nGeneral']
        
        # Configurar fondo gris para los nombres de las columnas
        for j in range(df.shape[1]):
            header_item = self.DFRMA.horizontalHeaderItem(j)
            if header_item is not None:
                header_item.setBackground(QColor('lightgray'))
        
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(df.iat[i, j]))
                
                # Aplicar color al texto basado en la columna "Clase"
                if df.columns[j] in columns_to_color:
                    if df.iat[i, j] == 'ACEPTABLE':
                        item.setForeground(QColor('green'))
                    elif df.iat[i, j] == 'INACEPTABLE':
                        item.setForeground(QColor('red'))
                    elif df.iat[i, j] == 'CUESTIONABLE':
                        item.setForeground(QColor('orange'))
                
                # Configurar fondo gris para todas las filas de la primera columna
                if j == 0:
                    item.setBackground(QColor('lightgray'))
                
                self.DFRMA.setItem(i, j, item)
        
        self.DFRMA.resizeColumnsToContents()
    
        # Aplicar estilo para los encabezados
        header = self.DFRMA.horizontalHeader()
        header.setStyleSheet("QHeaderView::section { background-color: lightgray; }")
        
    def REPORTRAT(self):
        try:
            self.Ano = self.VAno.text()
            self.Mes = self.VMes.text()

        except Exception as e:
            mensaje = f"Error: {e}. Valor no ingresado"
            QMessageBox.critical(self, "Inserte un Valor", mensaje)
        self.Estandar14 = (self.BD_Cont['CONTROLES'] == 'STD-14').sum()
        self.Pulpa = (self.BD_Cont['CONTROLES'] == 'PULPA').sum()
        self.Estandar15 = (self.BD_Cont['CONTROLES'] == 'STD-15').sum()
        self.DuplicadoCampo = (self.BD_Cont['CONTROLES'] == 'CAMPO ORIGIN').sum()
        self.Estandar13 = (self.BD_Cont['CONTROLES'] == 'STD-13').sum()
        self.BlancoFino = (self.BD_Cont['CONTROLES'] == 'BLK FINO').sum()
        self.Blanco = (self.BD_Cont['CONTROLES'] == 'BLANCO').sum()
        self.Rechazo = (self.BD_Cont['CONTROLES'] == 'RECHA').sum()
        self.MuestrasTotal = self.BD_Total['SAMPLE_ID'].count()
        self.ControlExterno = self.BD_RMA['CU'].count()
        self.TotalControles = self.Estandar14+self.Pulpa+self.Estandar15+self.DuplicadoCampo+self.Estandar13+self.BlancoFino+self.Blanco+self.Rechazo+self.ControlExterno
        RIDataframe = pd.DataFrame({
        'Año/Mes' : ['Total Muestras', 'Tipo de Control', 'Dup de Campo', 'Rechazo Grueso','Dup de Pulpa', 'STD', 'Control Externo', 'Blanco Fino', 'Blanco Grueso', 'Total Controles'],
        self.Ano : [self.MuestrasTotal, 'N° Muestras', self.DuplicadoCampo, self.Rechazo, self.Pulpa, self.Estandar13+self.Estandar14+self.Estandar15, self.ControlExterno, self.BlancoFino, self.Blanco, self.TotalControles ],
        self.Mes : [' ', 'RI (%)', 
                    (self.DuplicadoCampo*100/self.MuestrasTotal).round(2), 
                    (self.Rechazo*100/self.MuestrasTotal).round(2), 
                    (self.Pulpa*100/self.MuestrasTotal).round(2), 
                    ((self.Estandar13+self.Estandar14+self.Estandar15)*100/self.MuestrasTotal).round(2), 
                    (self.ControlExterno*100/self.MuestrasTotal).round(2), 
                    (self.BlancoFino*100/self.MuestrasTotal).round(2),
                    (self.Blanco*100/self.MuestrasTotal).round(2),
                    (self.TotalControles*100/self.MuestrasTotal).round(2) ]
        })
        self.fillRI(RIDataframe)

    def fillRI(self, df):
        # Configurar el número de filas y columnas
        self.RATINS.setRowCount(df.shape[0])
        self.RATINS.setColumnCount(df.shape[1])
    
        # Establecer los encabezados de las columnas
        self.RATINS.setHorizontalHeaderLabels(df.columns)
    
        # Obtener el header y aplicar estilo
        header = self.RATINS.horizontalHeader()
        header_font = QFont()
        header_font.setBold(True)
        header.setFont(header_font)
        header.setStyleSheet("QHeaderView::section { background-color: lightgray; }")
    
        # Definir estilos para las celdas
        cell_font = QFont()
        cell_font.setBold(True)
        gray_background = QBrush(QColor(200, 200, 200))
    
        # Rellenar el QTableWidget con los datos del DataFrame
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[i, j]))
    
                # Estilo para las dos primeras filas
                if i < 2:
                    item.setFont(cell_font)
                    item.setBackground(gray_background)
    
                # Estilo para la primera columna
                if j == 0:
                    item.setBackground(gray_background)
    
                # Estilo para la última fila
                if i == df.shape[0] - 1:
                    item.setFont(cell_font)
    
                self.RATINS.setItem(i, j, item)
    
        # Ajustar el tamaño de las columnas para que se ajusten al contenido
        self.RATINS.resizeColumnsToContents()
            
if __name__ == "__main__":
    app =  QtWidgets.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())