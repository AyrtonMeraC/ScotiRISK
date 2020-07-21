from aironDev import General_Register
import os
excel='aironDev/Catalogo.xlsx'

class Data:
    def __init__(self):
        self.info=[]

class Reporte:
    def __init__(self,banco,codId,nombre,anho,mes,dia):
        self.banco=banco
        self.codID=codId
        self.nombre=nombre
        self.dia=dia
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

def Tabla83_Disp():
    Tabla83=General_Register.Carga_Excel('aironDev/Tabla83.xlsx')
    list_83=[]
    for count in range(len(Tabla83.cod)):
        COD_83={
            'Codigo':Tabla83.cod[count],
            'Origen_Flujo':Tabla83.nombre[count],
        }
        list_83.append(COD_83)
    return list_83

def Carga_C46(data):
    count_total=0
    count_R1=0
    count_R2=0
    list_carga1=[]
    list_carga2=[]
    list_cuadratura=[]
    list_montos=[]
    list_83=Tabla83_Disp()
    for a in list_83:
        list_montos.append(0)
    for x in data :
        count_total=count_total+1
        Tipo_Registro=x[:2]
        C46_data={
            'Tipo_Registro':x[:2],
        }
        if(Tipo_Registro=='01'):
            C46=lista['C46_1']
            C46_data={
                'Tipo_Registro':x[:C46[0]],
                'Nivel_Consolidacion':x[C46[0]:C46[1]],
                'Tipo_Monto_Control':x[C46[1]:C46[2]],
                'Monto':int(x[C46[2]:C46[3]]),
                'Filler':x[C46[3]:C46[4]],
            }
            count_R1=count_R1+1
            list_carga1.append(C46_data)
        elif(Tipo_Registro=='02'):
            C46=lista['C46_2']
            C46_data={
                'Tipo_Registro':x[:C46[0]],
                'Nivel_Consolidacion':int(x[C46[0]:C46[1]]),
                'Tipo_Monto_Base':int(x[C46[1]:C46[2]]),
                'Tipo_Flujo':int(x[C46[2]:C46[3]]),
                'Banda_Temporal':x[C46[3]:C46[4]],
                'Moneda_Pago':x[C46[4]:C46[5]],
                'Origen_Flujo':int(x[C46[5]:C46[6]]),
                'Monto_Flujo':int(x[C46[6]:C46[7]])
            }
            count_R2=count_R2+1
            if(C46_data['Nivel_Consolidacion']==1 and C46_data['Tipo_Monto_Base']==3):
                count=0
                for a in list_83:
                    if(C46_data['Origen_Flujo']==a['Codigo']):
                        if(C46_data['Tipo_Flujo']==1):
                            list_montos[count]=list_montos[count]+C46_data['Monto_Flujo']
                        elif(C46_data['Tipo_Flujo']==2):
                            list_montos[count]=list_montos[count]-C46_data['Monto_Flujo']
                    count=count+1
            list_carga2.append(C46_data)
    
    Cuadratura={
        'Numero Registros Informados':count_total,
        'Numero Registros Codigo 01':count_R1,
        'Numero Registros Codigo 02':count_R2,
    }

    list_Individual=Tabla83_Disp()
    for x in list_Individual: 
        x['Contractual flujos a 30 días']=0
        x['Contractual flujos 30 días ME']=0
        x['Contractual flujos 90 días']=0
        x['Ajustada flujos a 30 días']=0
        x['Ajustada flujos 30 días ME']=0
        x['Ajustada flujos 90 días']=0
    
    list_Consolidado=Tabla83_Disp()
    for x in list_Consolidado: 
        x['Contractual flujos a 30 días']=0
        x['Contractual flujos 30 días ME']=0
        x['Contractual flujos 90 días']=0
        x['Ajustada flujos a 30 días']=0
        x['Ajustada flujos 30 días ME']=0
        x['Ajustada flujos 90 días']=0
        
    for x in list_carga2:
        
        if(x['Nivel_Consolidacion']==1):
            if(x['Tipo_Monto_Base']==1):
                
                if(int(x['Banda_Temporal'])<521):
                    if(int(x['Banda_Temporal'])<311 and int(x['Moneda_Pago'])==3):
                        for y in list_Individual:
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Contractual flujos 30 días ME']=y['Contractual flujos 30 días ME']+x['Monto_Flujo']
                                    y['Contractual flujos a 30 días']=y['Contractual flujos a 30 días']+x['Monto_Flujo']
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Contractual flujos 30 días ME']=y['Contractual flujos 30 días ME']-x['Monto_Flujo']
                                    y['Contractual flujos a 30 días']=y['Contractual flujos a 30 días']-x['Monto_Flujo']
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']-x['Monto_Flujo']
                    elif(int(x['Banda_Temporal'])<311):
                        for y in list_Individual:
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Contractual flujos a 30 días']=y['Contractual flujos a 30 días']+x['Monto_Flujo']
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Contractual flujos a 30 días']=y['Contractual flujos a 30 días']-x['Monto_Flujo']
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']-x['Monto_Flujo']
                    else:
                        for y in list_Individual: 
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']-x['Monto_Flujo']
                                    
            elif(x['Tipo_Monto_Base']==2):
                if(int(x['Banda_Temporal'])<521):
                    
                    if(int(x['Banda_Temporal'])<311 and int(x['Moneda_Pago'])==3):
                        for y in list_Individual:
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Ajustada flujos 30 días ME']=y['Ajustada flujos 30 días ME']+x['Monto_Flujo']
                                    y['Ajustada flujos a 30 días']=y['Ajustada flujos a 30 días']+x['Monto_Flujo']
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Ajustada flujos 30 días ME']=y['Ajustada flujos 30 días ME']-x['Monto_Flujo']
                                    y['Ajustada flujos a 30 días']=y['Ajustada flujos a 30 días']-x['Monto_Flujo']
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']-x['Monto_Flujo']
                    elif(int(x['Banda_Temporal'])<311):
                        for y in list_Individual:
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Ajustada flujos a 30 días']=y['Ajustada flujos a 30 días']+x['Monto_Flujo']
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Ajustada flujos a 30 días']=y['Ajustada flujos a 30 días']-x['Monto_Flujo']
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']-x['Monto_Flujo']
                    else:
                        for y in list_Individual: 
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']-x['Monto_Flujo']
        elif(x['Nivel_Consolidacion']==2):
            if(x['Tipo_Monto_Base']==1):
                
                if(int(x['Banda_Temporal'])<521):
                    if(int(x['Banda_Temporal'])<311 and int(x['Moneda_Pago'])==3):
                        for y in list_Consolidado:
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Contractual flujos 30 días ME']=y['Contractual flujos 30 días ME']+x['Monto_Flujo']
                                    y['Contractual flujos a 30 días']=y['Contractual flujos a 30 días']+x['Monto_Flujo']
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Contractual flujos 30 días ME']=y['Contractual flujos 30 días ME']-x['Monto_Flujo']
                                    y['Contractual flujos a 30 días']=y['Contractual flujos a 30 días']-x['Monto_Flujo']
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']-x['Monto_Flujo']
                    elif(int(x['Banda_Temporal'])<311):
                        for y in list_Consolidado:
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Contractual flujos a 30 días']=y['Contractual flujos a 30 días']+x['Monto_Flujo']
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Contractual flujos a 30 días']=y['Contractual flujos a 30 días']-x['Monto_Flujo']
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']-x['Monto_Flujo']
                    else:
                        for y in list_Consolidado: 
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Contractual flujos 90 días']=y['Contractual flujos 90 días']-x['Monto_Flujo']
            elif(x['Tipo_Monto_Base']==2):
                if(int(x['Banda_Temporal'])<521):
                    if(int(x['Banda_Temporal'])<311 and int(x['Moneda_Pago'])==3):
                        for y in list_Consolidado:
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Ajustada flujos 30 días ME']=y['Ajustada flujos 30 días ME']+x['Monto_Flujo']
                                    y['Ajustada flujos a 30 días']=y['Ajustada flujos a 30 días']+x['Monto_Flujo']
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Ajustada flujos 30 días ME']=y['Ajustada flujos 30 días ME']-x['Monto_Flujo']
                                    y['Ajustada flujos a 30 días']=y['Ajustada flujos a 30 días']-x['Monto_Flujo']
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']-x['Monto_Flujo']
                    elif(int(x['Banda_Temporal'])<311):
                        for y in list_Consolidado:
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Ajustada flujos a 30 días']=y['Ajustada flujos a 30 días']+x['Monto_Flujo']
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Ajustada flujos a 30 días']=y['Ajustada flujos a 30 días']-x['Monto_Flujo']
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']-x['Monto_Flujo']
                    else:
                        for y in list_Consolidado: 
                            if(x['Origen_Flujo']==int(y['Codigo'])):
                                if(x['Tipo_Flujo']==1):
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']+x['Monto_Flujo']
                                elif(x['Tipo_Flujo']==2):
                                    y['Ajustada flujos 90 días']=y['Ajustada flujos 90 días']-x['Monto_Flujo']
    list_cuadratura.append(Cuadratura)
    return list_carga1,list_carga2,list_cuadratura,list_montos,list_Individual,list_Consolidado