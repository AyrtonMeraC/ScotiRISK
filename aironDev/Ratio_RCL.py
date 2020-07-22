from aironDev import General_Register
import numpy as np
import openpyxl
import pandas as pd
from pandas import ExcelWriter
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

def Calculo_RCL(data,carga1,nombre,fecha):
    t88=Tabla88_Disp(os. getcwd() + '/aironDev/Tabla88.xlsx')
    t87=Tabla87_Disp(os. getcwd() + '/aironDev/Tabla87.xlsx')
    list_carga1=[]

    LCR=[]
    for i in t87:
        if(not(i in LCR)):
            LCR.append(t87[i]['LCR'])
    
    piv1=[]
    piv2=[]

    for i in LCR:
        piv1.append(i)
        info={
            'PesoI':0,
            'USDI':0,
            'Otras_1I':0,
            'Otras_2I':0,
            'Individual':0,
            'PesoC':0,
            'USDC':0,
            'Otras_1C':0,
            'Otras_2C':0,
            'Consolidado':0,
        }
        piv2.append(info)
    hash1 = {k:v for k, v in zip(piv1,piv2)}
    
    for i in data:
        cod=t87[i['Categoria']]['LCR']

        if(i['Nivel_Consolidacion']==1):
            if(i['Vencimiento_Contractual']==1):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['1_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoI']=hash1[cod]['PesoI']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDI']=hash1[cod]['USDI']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1I']=hash1[cod]['Otras_1I']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2I']=hash1[cod]['Otras_2I']+resultado
                hash1[cod]['Individual'] = hash1[cod]['Individual'] + resultado

            elif(i['Vencimiento_Contractual']==2):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['2_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoI']=hash1[cod]['PesoI']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDI']=hash1[cod]['USDI']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1I']=hash1[cod]['Otras_1I']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2I']=hash1[cod]['Otras_2I']+resultado
                hash1[cod]['Individual'] = hash1[cod]['Individual'] + resultado
            
            elif(i['Vencimiento_Contractual']==3):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoI']=hash1[cod]['PesoI']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDI']=hash1[cod]['USDI']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1I']=hash1[cod]['Otras_1I']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2I']=hash1[cod]['Otras_2I']+resultado
                hash1[cod]['Individual'] = hash1[cod]['Individual'] + resultado
            
            elif(i['Vencimiento_Contractual']==4):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoI']=hash1[cod]['PesoI']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDI']=hash1[cod]['USDI']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1I']=hash1[cod]['Otras_1I']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2I']=hash1[cod]['Otras_2I']+resultado
                hash1[cod]['Individual'] = hash1[cod]['Individual'] + resultado
            
            elif(i['Vencimiento_Contractual']==5):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoI']=hash1[cod]['PesoI']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDI']=hash1[cod]['USDI']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1I']=hash1[cod]['Otras_1I']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2I']=hash1[cod]['Otras_2I']+resultado
                hash1[cod]['Individual'] = hash1[cod]['Individual'] + resultado
            
            elif(i['Vencimiento_Contractual']==6):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoI']=hash1[cod]['PesoI']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDI']=hash1[cod]['USDI']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1I']=hash1[cod]['Otras_1I']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2I']=hash1[cod]['Otras_2I']+resultado
                hash1[cod]['Individual'] = hash1[cod]['Individual'] + resultado
            
            elif(i['Vencimiento_Contractual']==7):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoI']=hash1[cod]['PesoI']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDI']=hash1[cod]['USDI']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1I']=hash1[cod]['Otras_1I']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2I']=hash1[cod]['Otras_2I']+resultado
                hash1[cod]['Individual'] = hash1[cod]['Individual'] + resultado
        
        
        if(i['Nivel_Consolidacion']==2):

            if(i['Vencimiento_Contractual']==1):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['1_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoC']=hash1[cod]['PesoC']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDC']=hash1[cod]['USDC']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1C']=hash1[cod]['Otras_1C']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2C']=hash1[cod]['Otras_2C']+resultado
                hash1[cod]['Consolidado'] = hash1[cod]['Consolidado'] + resultado

            elif(i['Vencimiento_Contractual']==2):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['2_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoC']=hash1[cod]['PesoC']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDC']=hash1[cod]['USDC']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1C']=hash1[cod]['Otras_1C']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2C']=hash1[cod]['Otras_2C']+resultado
                hash1[cod]['Consolidado'] = hash1[cod]['Consolidado'] + resultado
            
            elif(i['Vencimiento_Contractual']==3):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoC']=hash1[cod]['PesoC']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDC']=hash1[cod]['USDC']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1C']=hash1[cod]['Otras_1C']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2C']=hash1[cod]['Otras_2C']+resultado
                hash1[cod]['Consolidado'] = hash1[cod]['Consolidado'] + resultado
            
            elif(i['Vencimiento_Contractual']==4):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoC']=hash1[cod]['PesoC']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDC']=hash1[cod]['USDC']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1C']=hash1[cod]['Otras_1C']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2C']=hash1[cod]['Otras_2C']+resultado
                hash1[cod]['Consolidado'] = hash1[cod]['Consolidado'] + resultado
            
            elif(i['Vencimiento_Contractual']==5):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoC']=hash1[cod]['PesoC']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDC']=hash1[cod]['USDC']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1C']=hash1[cod]['Otras_1C']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2C']=hash1[cod]['Otras_2C']+resultado
                hash1[cod]['Consolidado'] = hash1[cod]['Consolidado'] + resultado
            
            elif(i['Vencimiento_Contractual']==6):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoC']=hash1[cod]['PesoC']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDC']=hash1[cod]['USDC']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1C']=hash1[cod]['Otras_1C']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2C']=hash1[cod]['Otras_2C']+resultado
                hash1[cod]['Consolidado'] = hash1[cod]['Consolidado'] + resultado
            
            elif(i['Vencimiento_Contractual']==7):
                resultado=i['Flujo_Efectivo'] * t88[i['Categoria']]['3a7_RCL']
                if(i['Moneda']=='000'):
                    hash1[cod]['PesoC']=hash1[cod]['PesoC']+resultado
                elif(i['Moneda']=='013'):
                    hash1[cod]['USDC']=hash1[cod]['USDC']+resultado
                elif(i['Moneda']=='777'):
                    hash1[cod]['Otras_1C']=hash1[cod]['Otras_1C']+resultado
                elif(i['Moneda']=='888'):
                    hash1[cod]['Otras_2C']=hash1[cod]['Otras_2C']+resultado
                hash1[cod]['Consolidado'] = hash1[cod]['Consolidado'] + resultado

    for i in hash1:
        info={
            'Recalculo Auditorio':i,
            'CLP_Individual':hash1[i]['PesoI'],
            'USD_Individual':hash1[i]['USDI'],
            'Otras 1_Individual':hash1[i]['Otras_1I'],
            'Otras 2_Individual':hash1[i]['Otras_2I'],
            'Total Moneda 1':hash1[i]['Individual'],
            'CLP_Consolidado':hash1[i]['PesoC'],
            'USD_Consolidado':hash1[i]['USDC'],
            'Otras 1_Consolidado':hash1[i]['Otras_1C'],
            'Otras 2_Consolidado':hash1[i]['Otras_2C'],
            'Total Moneda 2':hash1[i]['Consolidado'],
        }
        list_carga1.append(info)

    cabecera=Cabecera_Cuadratura(carga1)
    limite=Ingresos_Lim(hash1['Ingresos'],hash1['Ingresos individual de riesgo < A3 o equivalente. (2)'],hash1['Egreso'],cabecera)
    list_carga1.append(limite)
    egresos=Egresos_Netos(limite,hash1['Egreso'])
    list_carga1.append(egresos)
    dif_caratula=Diferencia_Caratula(cabecera[1:],egresos)
    list_carga1.append(dif_caratula)
    General_Register.reporte_Vista(cabecera,list_carga1,nombre+'_'+fecha+'_RATIO')

def Egresos_Netos(limite,dato):
    info={
        'Recalculo Auditorio':'Egresos Netos',
        'CLP_Individual':dato['PesoI']-limite['CLP_Individual'],
        'USD_Individual':dato['USDI']-limite['USD_Individual'],
        'Otras 1_Individual':dato['Otras_1I']-limite['Otras 1_Individual'],
        'Otras 2_Individual':dato['Otras_2I']-limite['Otras 2_Individual'],
        'Total Moneda 1':dato['Individual']-limite['Total Moneda 1'],
        'CLP_Consolidado':dato['PesoC']-limite['CLP_Consolidado'],
        'USD_Consolidado':dato['USDC']-limite['USD_Consolidado'],
        'Otras 1_Consolidado':dato['Otras_1C']-limite['Otras 1_Consolidado'],
        'Otras 2_Consolidado':dato['Otras_2C']-limite['Otras 2_Consolidado'],
        'Total Moneda 2':dato['Consolidado']-limite['Total Moneda 2'],
    }
    return info

def Diferencia_Caratula(cabecera,egresos):
    data={
        'Recalculo Auditorio':'Diferencia contra carátula ',
    }
    for i in egresos:
        if(egresos[i]!='Egresos Netos'):
            data[i]=cabecera[0][i]-egresos[i]
    return data

def Ingresos_Lim(data1,data2,data3,cabecera):
    info={
        'Recalculo Auditorio':'Ingresos con Lim 75%',
    }

    info2={
        'Recalculo Auditorio':'Ingresos con Lim 75%',
    }

    info['CLP_Individual']=data1['PesoI']+data2['PesoI']
    info['USD_Individual']=data1['USDI']+data2['USDI']
    info['Otras 1_Individual']=data1['Otras_1I']+data2['Otras_1I']
    info['Otras 2_Individual']=data1['Otras_2I']+data2['Otras_2I']
    info['Total Moneda 1']=data1['Individual']+data2['Individual']
    info['CLP_Consolidado']=data1['PesoC']+data2['PesoC']
    info['USD_Consolidado']=data1['USDC']+data2['USDC']
    info['Otras 1_Consolidado']=data1['Otras_1C']+data2['Otras_1C']
    info['Otras 2_Consolidado']=data1['Otras_2C']+data2['Otras_2C']
    info['Total Moneda 2']=data1['Consolidado']+data2['Consolidado']

    info2['CLP_Individual']=(data3['PesoI'] * 0.75)
    info2['USD_Individual']=(data3['USDI'] * 0.75)
    info2['Otras 1_Individual']=(data3['Otras_1I'] * 0.75)
    info2['Otras 2_Individual']=(data3['Otras_2I'] * 0.75)
    info2['Total Moneda 1']=(data3['Individual'] * 0.75)
    info2['CLP_Consolidado']=(data3['PesoC'] * 0.75)
    info2['USD_Consolidado']=(data3['USDC'] * 0.75)
    info2['Otras 1_Consolidado']=(data3['Otras_1C'] * 0.75)
    info2['Otras 2_Consolidado']=(data3['Otras_2C'] * 0.75)
    info2['Total Moneda 2']=(data3['Consolidado'] * 0.75)

    for i in info:
        if(info[i] > info2[i]):
            info[i]=info2[i]
    
    return info
            

def Cabecera_Cuadratura(carga1):
    list_cabecera=[]
    info={
        'Recalculo Auditorio':'Activos_Liquidos',
        'CLP_Individual':carga1[0]['Activos_Liquidos'],
        'USD_Individual':carga1[1]['Activos_Liquidos'],
        'Otras 1_Individual':carga1[2]['Activos_Liquidos'],
        'Otras 2_Individual':carga1[3]['Activos_Liquidos'],
        'Total Moneda 1':carga1[4]['Activos_Liquidos'],
        'CLP_Consolidado':carga1[5]['Activos_Liquidos'],
        'USD_Consolidado':carga1[6]['Activos_Liquidos'],
        'Otras 1_Consolidado':carga1[7]['Activos_Liquidos'],
        'Otras 2_Consolidado':carga1[8]['Activos_Liquidos'],
        'Total Moneda 2':carga1[9]['Activos_Liquidos'],
    }
    list_cabecera.append(info)
    info={
        'Recalculo Auditorio':'Egresos_Netos',
        'CLP_Individual':carga1[0]['Egresos_Netos'],
        'USD_Individual':carga1[1]['Egresos_Netos'],
        'Otras 1_Individual':carga1[2]['Egresos_Netos'],
        'Otras 2_Individual':carga1[3]['Egresos_Netos'],
        'Total Moneda 1':carga1[4]['Egresos_Netos'],
        'CLP_Consolidado':carga1[5]['Egresos_Netos'],
        'USD_Consolidado':carga1[6]['Egresos_Netos'],
        'Otras 1_Consolidado':carga1[7]['Egresos_Netos'],
        'Otras 2_Consolidado':carga1[8]['Egresos_Netos'],
        'Total Moneda 2':carga1[9]['Egresos_Netos'],
    }
    list_cabecera.append(info)
    return list_cabecera

def Tabla88_Disp(excel):
    doc = openpyxl.load_workbook(excel)
    doc.get_sheet_names()
    hoja = doc.get_sheet_by_name('Hoja1')
    hoja.rows
    list_Cu=[]
    list_CM=[]
    for filas in hoja.rows:
        data2={
            '1_RCL':filas[1].value,
            '2_RCL':filas[2].value,
            '3a7_RCL':filas[3].value,
            '1_RFEN':filas[4].value,
            '2_RFEN':filas[5].value,
            '3_RFEN':filas[6].value,
            '4_RFEN':filas[7].value,
            '5_RFEN':filas[8].value,
            '6_RFEN':filas[9].value,
            '7_RFEN':filas[10].value,
        }
        list_Cu.append(str(filas[0].value))
        list_CM.append(data2)
    hash = {k:v for k, v in zip(list_Cu,list_CM)}
    return hash

def Tabla87_Disp(excel):
    doc = openpyxl.load_workbook(excel)
    doc.get_sheet_names()
    hoja = doc.get_sheet_by_name('Hoja1')
    hoja.rows
    list_Cu=[]
    list_CM=[]
    for filas in hoja.rows:
        data={
            'LCR':filas[1].value,
            'NSFR':filas[2].value,
            'Pais Domicilio':filas[3].value,
            'Tipo Flujo':filas[4].value,
            'Categoria':filas[5].value,
        }
        list_Cu.append(str(filas[0].value))
        list_CM.append(data)
    hash = {k:v for k, v in zip(list_Cu,list_CM)}
    return hash
