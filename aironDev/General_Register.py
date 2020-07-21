import numpy as np
import openpyxl
import requests
import pandas as pd
from pandas import ExcelWriter
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

class Data:
    def __init__(self):
        self.info=[]

class Reporte:
    def __init__(self,banco,codId,nombre,anho,mes):
        self.banco=banco
        self.codID=codId
        self.nombre=nombre
        self.anho=anho
        self.mes=mes
        self.registro=[]
        
class Catalogo:
    def __init__(self):
        self.cod=[]
        self.nombre=[]
        
class Persona:
    def __init__(self,nombre,rut,dv):
        self.nombre=nombre
        self.rut=rut
        self.dv=dv

def Deco_Catalog(cod,cat):
    i=0
    for i in range(len(cat.cod)):
        if(cod==cat.cod[i]):
            return cat.nombre[i]


"""# Carga de Archivo .DAT"""

def Charge_Info(filename):
    file = open(filename, 'r')
    data=Data()
    for linea in file:
        data.info.append(linea)
    return data


"""# Carga Catalogo Excel"""

def Carga_Excel(excel):
    doc = openpyxl.load_workbook(excel)
    doc.get_sheet_names()
    hoja = doc.get_sheet_by_name('Hoja1')
    hoja.rows
    cat=Catalogo()
    for filas in hoja.rows:
        cat.cod=np.append(cat.cod,filas[0].value)
        cat.nombre=np.append(cat.nombre,filas[1].value)
    return cat

def Carga_Parametros(excel):
    doc = openpyxl.load_workbook(excel)
    hoja = doc['Hoja1']
    list_registro=[]
    list_1=[]
    list_2=[]
    largo=0
    for i in hoja.rows:
        largo=len(i)
    
    for count in range(largo):
        list_registro=[]
        for i in hoja.rows:
            if(i[count].value!=None):
                list_registro.append(i[count].value)
        list_1.append(list_registro[0])
        list_2.append(list_registro[1:])
    hash = {k:v for k, v in zip(list_1, list_2)}
    return hash

def reporte_excel(list_p,count,nombreReporte):
    if(len(list_p)<1000000):
        writer = ExcelWriter(nombreReporte+'_'+str(count)+'.xlsx')
        df = pd.DataFrame(data=list_p)
        df.to_excel(writer, 'Reporte', index=False)
        writer.save()
    else:
        writer1 = ExcelWriter(nombreReporte+'_Parte1_'+str(count)+'.xlsx')
        df = pd.DataFrame(data=list_p[:1000000])
        df.to_excel(writer1, 'Reporte', index=False)
        writer1.save()
        writer2 = ExcelWriter(nombreReporte+'_Parte2_'+str(count)+'.xlsx')
        df = pd.DataFrame(data=list_p[1000000:])
        df.to_excel(writer2, 'Reporte', index=False)
        writer2.save()

def reporte_Vista(list_i,list_c,nombreReporte):
    if(len(list_i)<1048576):
        writer = ExcelWriter(nombreReporte+'.xlsx')
        df = pd.DataFrame(data=list_i)
        df2 = pd.DataFrame(data=list_c)
        df.to_excel(writer, 'Reporte', index=False)
        df2.to_excel(writer, 'Reporte', index=False, startrow=len(list_i)+2)
        writer.save()

def Cheq_ASCII(i):
    for x in i: 
        cod=ord(x)
        if((cod<48 or cod>57) and (cod<65 or cod>90) and (cod!=38) and (cod!=44) and (cod!=47) and (cod!=32) and (cod!=10)):
            return ("Contiene Caracter ASCII NO VALIDO: "+x)
    return 0

"""Posterior a que incluya la funcion de validacion mediante SII, esta funcion variará
"""
def Validar_Rut(rut,dv):
    j=1
    suma=0
    r=int(rut)
    if(r<50000000):
        for i in rut:
            if(j==8):
                j=2
            piv=int(i)
            suma=(piv*j)+suma
            j=j+1
        module=suma/11
        total=suma-(11*module)
        digito=11-total

        if(digito==11):
            val='0'
        elif(digito==10):
            val='K'
        else:
            val=digito
    
        if(val!=dv):
            return rut
        else:
            return 0
    else:
        return 0

def Validacion_Nombre(nombre):
    abrevia=Carga_Excel('aironDev/Abreviaturas.xlsx')
    i=nombre.find('/')
    j1=0
    if(i!=-1):
        j1=i+1
        i=nombre[j1:].find('/')
        if(i==-1): #Falta otro /
            return -2
        elif(nombre[j1]==" "):
            return-3
        else:
            j1=j1+i+1
            i=nombre[j1:].find('/')
            #En esta validacion veo si existe algún otro "/" adicional erroneo
            if(i!=-1):
                return-5
            else:
                return 1
    else:
        for a in range(len(abrevia.nombre)):
            if((abrevia.nombre[a] in nombre) or (abrevia.cod[a] in nombre)):
                return 1
        return 2

def EliminarRutRep(list_rut):
    ruts=np.array
    ruts=[]
    flag=0
    
    for i in range(len(list_rut)):
        
        if(len(ruts)!=0):
            for j in range(len(ruts)):
                if(i!=j):
                    if(list_rut[i].rut==ruts[j].rut):
                        flag=1
            if(flag==0):
                ruts.append(list_rut[i])
        else:
            ruts.append(list_rut[i])
        
        flag=0
    
    return ruts

def Formato_Nombre(nombre):
    i=nombre.find('/')
    j1=0
    j2=0
    if(i!=-1):
        j1=i
        apeP=nombre[:j1]
        j1=j1+1
        i=nombre[j1:].find('/')
        apeM=nombre[j1:j1+i]
        j1=j1+i+1
        i=nombre[j1:].find(' ')
        j2=j1+i+1
        i=nombre[j2:].find(' ')
        name=nombre[j1:j2+i]
        name=name+" "+apeP+" "+apeM
        return name
    else:
        name=nombre
        list_n=[]
        nombre_r=""
        flag=0
        while(flag==0):
            i=name.find(' ')
            j1=i
            largo=len(name)
            if(largo!=j1+1 and name[j1+1]!=' ' and j1!=-1):
                list_n.append(name[:j1])
                name=name[j1+1:]
            else:
                flag=1
                i=name.find(' ')
                if(i!=-1):
                    list_n.append(name[:i])
                else:
                    list_n.append(name)
                    
        for i in list_n:
            for j in i:
                cod=ord(j)
                if((cod>47 and cod<58) or (cod>64 and cod<91) or (cod==38) or (cod==44) or (cod==47) or (cod==32) or (cod==10)):
                    nombre_r=nombre_r+j
            nombre_r=nombre_r+' '
        return(nombre_r)

def Validacion_General(persona):
    
    nombre=Formato_Nombre(persona.nombre)
    persona_2=Validar_Rut_Consulta(persona.rut,persona.dv)

    if(persona_2!=None):
        persona_2.nombre=Formato_Nombre(persona_2.nombre)
        
        if(nombre!=persona_2.nombre and persona_2!=None and persona_2.nombre!='&'):
            errores=(Comparar_Nombres(nombre, persona_2.nombre))
            if(errores!=-1):
                data={
                    'Rut':persona.rut,
                    'Rut_DV':persona.dv,
                    'Nombre_Ingresado':errores['Nombre_Ingresado'],
                    'Nombre_Esperado':errores['Nombre_Esperado'],
                }
                if(not('&' in data['Nombre_Esperado'])):
                    return data
                else:
                    return -1
            else:
                return -1
    else:
        print("No se pudo encontrar por el Rut ",persona.nombre)
        data={
            'Rut':persona.rut,
            'Rut_DV':persona.dv,
            'Nombre_Ingresado':nombre,
            'Nombre_Esperado':'Nombre No Encontrado',
        }
        return data

def Validar_Rut_Consulta(rut,dv):
    payload = {'RUT': rut, 'DV': dv}
    try:
        r = requests.post('https://zeus.sii.cl/cvc_cgi/nar/nar_consulta/get/',data=payload)
        todo=r.text
        large=len('<td width="300"><font class="texto">')
        i=todo.find('<td width="300"><font class="texto">')
        if(i!=-1):
            j=todo.find('</font></td></tr>')
            nombre=(todo[i+large:j])
            person=Persona(nombre, rut, dv)
            return(person)
    except:
        print("No se Pudo Establecer una Conexion Con el Servidor")
        person=Persona("Sin Conexion", rut, dv)
        return person

def Comparar_Nombres(nombre1, nombre2):
    countA=0
    countB=0
    limS1=0
    limS2=0
    limI1=0
    limI2=0
    contador_general=0
    conta_palabras=0
    abrevia=Carga_Excel('aironDev/Abreviaturas.xlsx')
    
    while(countA!=len(nombre1) and countB!=len(nombre2)):
        if(nombre1[countA]==' ' and nombre2[countB]==' '):
            conta_palabras=conta_palabras+1
            limI1=limS1
            limI2=limS2
            limS1=countA
            limS2=countB
            if(nombre1[limI1+1:limS1]=='S' and (len(nombre1[limI1:limS1])!=len(nombre2[limI2:limS2]))):
                count=Comparar_Palabras(nombre1[limI1:limS1+2],nombre2[limI2:limS2],abrevia)+1
                contador_general=contador_general+count
            if(nombre2[limI2+1:limS2]=='SA'):
                contador_general=contador_general-1
                conta_palabras=conta_palabras+1
            count=Comparar_Palabras(nombre1[limI1:limS1],nombre2[limI2:limS2],abrevia)
            contador_general=contador_general+count
            countA=countA+1
            countB=countB+1
        else:
            if(nombre1[countA]!=' ' and nombre2[countB]==' '):
                countA=countA+1
            else:
                if(nombre2[countB]!=' ' and nombre1[countA]==' '):
                    countB=countB+1
                else:
                    if(nombre1[countA]!=' ' and nombre2[countB]!=' '):
                        countA=countA+1
                        countB=countB+1
    
    if(conta_palabras != contador_general):
        data={
            'Nombre_Ingresado':nombre1,
            'Nombre_Esperado':nombre2,
        }
        return (data)
    else:
        return -1

def Comparar_Palabras(p1,p2,abrevia):
    countA=0
    countB=0
    contador_general=0
    for i in range(len(abrevia.nombre)):
        if((p1[1:] == abrevia.nombre[i]) or (p2 == abrevia.cod[i])):
            return 1
        else:
            if((p1[1:] == abrevia.cod[i]) or (p2 == abrevia.nombre[i])):
                return 1
    
    while(countA!=len(p1) and countB!=len(p2)):
        if(len(p1) > len(p2)):
            if(p1[countA] in p2):
                contador_general=contador_general+1
            countA=countA+1
        else:
            if(p2[countB] in p1):
                contador_general=contador_general+1
            countB=countB+1
    
    
    if(contador_general==len(p1) or contador_general==len(p2)):
        return 1
        
    return 0



