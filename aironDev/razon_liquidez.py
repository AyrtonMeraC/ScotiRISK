from aironDev import General_Register
from aironDev import Ratio_RCL
import os
excel='aironDev/Catalogo.xlsx'

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

lista=General_Register.Carga_Parametros('aironDev/CatalogoC.xlsx')

def Tabla73_Disp():
    Tabla73=General_Register.Carga_Excel('aironDev/Tabla73.xlsx')
    list_73=[]
    for count in range(len(Tabla73.cod)):
        codigo=int(Tabla73.cod[count])
        COD_73={
            'Codigo':str(codigo),
            'Origen_Flujo':Tabla73.nombre[count],
        }
        list_73.append(COD_73)
    return list_73

def Carga_C49(data,key,nombre,fecha):
    list_cuadratura=[]
    list_carga1=[]
    list_carga2=[]
    list_consolidado=Tabla73_Disp()
    list_individual=Tabla73_Disp()
    individual=0
    c_local=0
    c_global=0

    for x in list_individual:
        x[str(key)]=0
    
    for x in list_consolidado:
        x[str(key)]=0

    for x in data:
        Tipo_Registro=x[:2]
        C49_data={
            'Tipo_Registro':x[:2],
        }
        if(Tipo_Registro=='01'):
            C49=lista['C49_1']
            C49_data={
                'Tipo_Registro':x[:C49[0]],
                'Fecha':x[C49[0]:C49[1]],
                'Nivel_Consolidacion':int(x[C49[1]:C49[2]]),
                'Moneda':int(x[C49[2]:C49[3]]),
                'Activos_Liquidos':int(x[C49[3]:C49[4]]),
                'Egresos_Netos':int(x[C49[4]:C49[5]]),
                'Fuentes_Financiamiento_Estable':int(x[C49[5]:C49[6]]),
                'Financiamiento_Estable_Requerido':int(x[C49[6]:C49[7]]),
            }
            list_carga1.append(C49_data)
            if(C49_data['Nivel_Consolidacion']==1):
                individual=individual+1
            elif(C49_data['Nivel_Consolidacion']==2):
                c_local=c_local+1
            elif(C49_data['Nivel_Consolidacion']==3):
                c_global=c_global+1

        elif(Tipo_Registro=='02'):
            C49=lista['C49_2']
            C49_data={
                'Tipo_Registro':x[:C49[0]],
                'Fecha':x[C49[0]:C49[1]],
                'Nivel_Consolidacion':int(x[C49[1]:C49[2]]),
                'Categoria':x[C49[2]:C49[3]],
                'Vencimiento_Contractual':int(x[C49[3]:C49[4]]),
                'Pais':x[C49[4]:C49[5]],
                'Moneda':x[C49[5]:C49[6]],
                'Tipo_Flujo':int(x[C49[6]:C49[7]]),
                'Flujo_Efectivo':int(x[C49[7]:C49[8]])*int(x[C49[8]:C49[9]]+'1'),
                'Filler':x[C49[9]:C49[10]],
            }
            list_carga2.append(C49_data)
            if(C49_data['Nivel_Consolidacion']==1):
                individual=individual+1
                for i in list_individual:
                    if(C49_data['Categoria']==str(i['Codigo'])):
                        if(C49_data['Tipo_Flujo']==1):
                            i[str(key)]=i[str(key)]+C49_data['Flujo_Efectivo']

            elif(C49_data['Nivel_Consolidacion']==2):
                c_local=c_local+1
                for i in list_consolidado:
                    if(C49_data['Categoria']==str(i['Codigo'])):
                        if(C49_data['Tipo_Flujo']==1):
                            i[str(key)]=i[str(key)]+C49_data['Flujo_Efectivo']
                            
            elif(C49_data['Nivel_Consolidacion']==3):
                c_global=c_global+1
    Cuadratura={
        'Numero Registros Consolidacion Individual':individual,
        'Numero Registros Consolidado Local':c_local,
        'Numero Registros Consolidado Global':c_global,
    }
    list_cuadratura.append(Cuadratura)
    Ratio_RCL.Calculo_RCL(list_carga2,list_carga1,nombre,fecha)
    return list_carga1,list_carga2,list_cuadratura,list_individual,list_consolidado