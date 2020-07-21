from aironDev import General_Balance
from aironDev import situacion_liquidez
from aironDev import razon_liquidez
from aironDev import General_Register
import os
import shutil
import time
import threading

class Data:
    def __init__(self):
        self.info=[]

def Menu1():
    os.system ("cls") 
    direc=input('Ingrese Directorio: ')
    list_B=[]
    list_fecha=[]
    list_carga=[]
    list_carga2=[]
    mapeo=""
    list_dir=os.listdir(direc)

    for i in list_dir:
        if('balance.xlsx' in i):
            num=i.find('_balance.xlsx')
            fecha=i[:num]
            print(fecha)
            list_fecha.append(fecha)
            list_B.append(direc+'\\'+i)
        elif('mapeo.xlsx' in i):
            mapeo=direc+'\\'+i
    
    list_83=General_Balance.Tabla83_Disp()
    list_73=General_Balance.Tabla73_Disp()
    for x in list_83:
        data={
            'Codigo':x,
            'Nombre':list_83[x]['Nombre'],
        }
        list_carga.append(data)
    
    for x in list_73:
        data={
            'Codigo':x,
            'Nombre':list_73[x]['Nombre'],
        }
        list_carga2.append(data)

    for i in range(len(list_B)):
        list_piv1=[]
        lista_piv1=list_carga
        list_piv2=[]
        lista_piv2=list_carga2
        rM=General_Balance.Carga_M(mapeo)
        rB=General_Balance.Carga_B(list_B[i])
        M49=General_Balance.Carga_MC49(mapeo)
        list_saldos,list_problemas,list_saldos49=General_Balance.Carga_Balance(rB,rM,M49)
        count=0
        for x in lista_piv1:
            x[list_fecha[i]]=list_saldos[count]
            count=count+1
            list_piv1.append(x)
        list_carga=list_piv1
        count=0
        for x in lista_piv2:
            x[list_fecha[i]]=list_saldos49[count]
            count=count+1
            list_piv2.append(x)
        list_carga2=list_piv2
    General_Register.reporte_excel(list_carga,1,'Archivo_Balance_C46')
    General_Register.reporte_excel(list_carga2,1,'Archivo_Balance_C49')
    General_Register.reporte_excel(list_problemas,1,'Archivo_Balance_Problemas')
    Mover_Excels(direc,"Balance")

def Menu2():
    list_B=[]
    list_carga=[]
    os.system ("cls") 
    direc=input('Ingrese Directorio: ')
    list_dir=os.listdir(direc)
    for i in list_dir:
        if('.dat' in i):
            print(i)
            list_B.append(direc+'\\'+i)
    
    list_83=General_Balance.Tabla83_Disp()
    for x in list_83:
        data={
            'Codigo':x,
            'Nombre':list_83[x]['Nombre'],
        }
        list_carga.append(data)

    for i in list_B:
        list_piv=[]
        lista_piv=list_carga
        data=General_Register.Charge_Info(i)
        letra=data.info[0][4]
        if(letra=='C'):
            nombre=str(data.info[0][4:7])
            fecha=str(data.info[0][7:11])+str(data.info[0][11:13])+str(data.info[0][13:15])
        else:
            nombre=str(data.info[0][3:6])
            fecha=str(data.info[0][6:10])+str(data.info[0][10:12])+str(data.info[0][12:14])
        list_carga1,list_carga2,list_cuadratura,list_montos,list_Individual,list_Consolidado = situacion_liquidez.Carga_C46(data.info[1:])
        General_Register.reporte_excel(list_carga1,1,nombre+'_'+fecha)
        General_Register.reporte_excel(list_carga2,2,nombre+'_'+fecha)
        General_Register.reporte_excel(list_cuadratura,3,nombre+'_'+fecha+'_CUADRATURA')
        General_Register.reporte_Vista(list_Individual,list_Consolidado,nombre+'_'+fecha+'_RESUMEN')
        count=0
        for x in lista_piv:
            x[str(fecha)]=list_montos[count]
            count=count+1
            list_piv.append(x)
        list_carga=list_piv
    General_Register.reporte_excel(list_carga,4,nombre+'_'+fecha+'_VISTA')
    Mover_Excels(direc,"C46")

def Menu3():
    list_B=[]
    list_carga1=[]
    list_carga2=[]
    os.system ("cls") 
    direc=input('Ingrese Directorio: ')
    list_dir=os.listdir(direc)
    for i in list_dir:
        if('.dat' in i):
            print(i)
            list_B.append(direc+'\\'+i)
    
    list_73=razon_liquidez.Tabla73_Disp()
    for x in list_73:
        data={
            'Codigo':x['Codigo'],
            'Nombre':x['Origen_Flujo'],
        }
        list_carga1.append(data)
        list_carga2.append(data)

    for i in list_B:

        list_piv1=[]
        lista_piv1=list_carga1

        list_piv2=[]
        lista_piv2=list_carga2

        data=General_Register.Charge_Info(i)
        letra=data.info[0][4]

        if(letra=='C'):
            nombre=str(data.info[0][4:7])
            key=str(data.info[0][7:11])+str(data.info[0][11:13])
            fecha=str(data.info[0][7:11])+str(data.info[0][11:13])+str(data.info[0][13:15])
        else:
            nombre=str(data.info[0][3:6])
            key=str(data.info[0][6:10])+str(data.info[0][10:12])
            fecha=str(data.info[0][6:10])+str(data.info[0][10:12])+str(data.info[0][12:14])
        
        list_c1,list_c2,list_cuadratura,list_Individual,list_Consolidado = razon_liquidez.Carga_C49(data.info[1:],key,nombre,fecha)
        General_Register.reporte_excel(list_c1,1,nombre+'_'+fecha)
        General_Register.reporte_excel(list_c2,2,nombre+'_'+fecha)
        General_Register.reporte_excel(list_cuadratura,0,nombre+'_'+fecha+'_CUADRATURA')
        count=0
        
        for x in lista_piv1:
            x[str(fecha)]=list_Individual[count][str(key)]
            count=count+1
            list_piv1.append(x)
        list_carga1=list_piv1

        count=0
        for x in lista_piv2:
            x[str(fecha)]=list_Consolidado[count][str(key)]
            count=count+1
            list_piv2.append(x)
        list_carga2=list_piv2

    General_Register.reporte_Vista(list_carga1,list_carga2,nombre+'_'+fecha+'_CONTRACTUAL')
    Mover_Excels(direc,"C49")

def Mover_Excels(dirname,nombre):
    list_B=[]
    os.mkdir(dirname+"\\"+nombre)
    directorio=os.getcwd()
    list_B=os.listdir(directorio)
    for i in list_B:
        if('.xlsx' in i):
            print(i)
            shutil.move(i,dirname+"\\"+nombre)

def Menu_Principal():
    print("1-Generar Vista Balance")
    print("2-Generar Vista Contractual C46")
    print("3-Generar Vista Contractual C49")
    opcion=input("Ingrese Opcion: ")
    if(opcion=='1'):
        Menu1()
    elif(opcion=='2'):
        Menu2()
    elif(opcion=='3'):
        Menu3()
    else:
        print("Opcion No Valida")

Menu_Principal()
