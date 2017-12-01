import math,win32com.client,win32con,win32api

try:
    Visum
except:
    Visum=win32com.client.Dispatch("Visum.Visum")
    Visum.LoadVersion("D:/VDF_test.ver")

def GetVol(CR,CrType,tc,qmax,t0,q): 
            Data=CR[CrType]
            if CrType=="HCM":                
                a=Data[0]
                b=Data[1]
                c=Data[2]
                pierw=(((tc/t0)-1)/a)**(1/float(b))
                Vol=qmax*c*pierw
                
                 
            if CrType=="HCM2":
                a=Data[0]
                b1=Data[1]
                b2=Data[2]
                c=Data[3]
                if q>qmax:
                    b = b2 
                else: b=b1
                pierw=(((tc/t0)-1)/a)**(1/float(b))
                Vol=qmax*c*pierw
                
                                                
            if CrType=="HCM3":
                a=Data[0]
                b=Data[1]
                c=Data[2]
                d=Data[3]
                if q>qmax:                    
                    pierw=((((tc-(q-qmax)*d)/t0)-1)/a)**(1/float(b))
                else: 
                    pierw=(((tc/t0)-1)/a)**(1/float(b))
                Vol=qmax*c*pierw 
                 
                 
            if CrType=="CONICAL":
                a=Data[0]
                b=(2*a-1)/(2*a-2)
                c=Data[1]
                
            if CrType=="CONICAL_MARGINAL":        
                a=Data[0]
                cf=Data[1] 
                
            if CrType=="EXPONENTIAL":
                a=Data[0]
                b=Data[1]
                c=Data[2]
                d=Data[3]
                sc=Data[4]
                if q>qmax:                    
                    Vol=(tc-t0-math.exp(a*sc)/b+d*sc)*qmax*c/d
                else: 
                    Vol=(math.log(b*(tc-t0))/a)*qmax*c
                 
                
            if CrType=="INRETS":
                a=Data[0]
                c=Data[1]
                if q>qmax:
                    Vol=qmax*c*((0.1*tc/t0)/(1.1-a))**(0.5)
                else: 
                    Vol=1.1*qmax*c*(1-tc/t0)/(a-tc/t0)
                  
                
                
            if CrType=="LOGISTIC":
                a=Data[0]
                b=Data[1]
                c=Data[2]
                d=Data[3]
                f=Data[4]
                
                Vol= qmax*c*(b-math.log((a/(tc-t0)-1)/f))/d
                
                        
            if CrType=="QUADRATIC":
                a=Data[0]
                b=Data[1]
                c=Data[2]
                d=Data[3]
                u=(a+t0-tc)
                Vol=((b**2-4*d*u)**(0.5)-b)*qmax*c/(2*d)
                
                
               
            if CrType=="SIGMOIDAL_MMF_LINKS":
                a=Data[0]
                b=Data[1]
                cf=Data[2]
                d=Data[3]
                f=Data[4]
                 
            if CrType=="SIGMOIDAL_MMF_NODES":
                a=Data[0]
                b=Data[1]
                cf=Data[2]
                d=Data[3]
                f=Data[4] 
                
            if CrType=="Akcelik":
                a=Data[0]
                b=Data[1]
                cf=Data[2]
                d=Data[3]
                
                
            if CrType=="Lohse":
                a=Data[0]
                b=Data[1]
                c=Data[2]
                sc=Data[3]
                if q<qmax:
                    pierw=(((tc/t0)-1)/a)**(1/float(b))
                    Vol=qmax*c*pierw                    
                else:
                    Vol=qmax*c*((tc-t0*(1+a*sc**b))/(a*b*t0*sc**(b-1))+sc)
                
            
            if CrType=="Linear bottle-neck":
                pass
                #Data=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor")]
                
                
            if CrType=="Akcelik2":
                a=Data[0]
                b=Data[1]
                cf=Data[2]
                sc=Data[3]
                
            if CrType=="TMODEL_LINKS":
                [a1,a2,b1,b2,cf,d1,d2,f1,f2,sc]=Data
                
            if CrType=="TMODEL_NODES":
                [a1,a2,b1,b2,cf,d1,d2,f1,f2,sc]=Data
            
            try:    
                if Vol>0:
                    print Vol-q
                return Vol
            except:
                pass
            


def Make_CR_Dict(Crs):
    def AddCR(CrType,Container):
            Data=False
            if CrType=="HCM":                
                Data=[Container.AttValue("hcm_a"),Container.AttValue("hcm_b"),Container.AttValue("capacityFactor")]
                 
            if CrType=="HCM2":                
                Data=[Container.AttValue("hcm2_a"),
                      Container.AttValue("hcm2_b1"),
                      Container.AttValue("hcm2_b2"),
                      Container.AttValue("capacityFactor")]
                                                
            if CrType=="HCM3":
                Data=[Container.AttValue("hcm3_a"),
                      Container.AttValue("hcm3_b"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("hcm3_d")]
            
            if CrType=="CONICAL":
                Data=[Container.AttValue("conical_a"),
                      Container.AttValue("capacityFactor")]
            
            if CrType=="CONICAL_MARGINAL":        
                Data=[Container.AttValue("conical_marginal_a"),
                      Container.AttValue("capacityFactor")]
               
            if CrType=="EXPONENTIAL":
                Data=[Container.AttValue("exponential_a"),
                                        Container.AttValue("exponential_b"),
                                        Container.AttValue("capacityFactor"),
                                        Container.AttValue("exponential_d"),
                                        Container.AttValue("exponential_satcrit")]
                  
            if CrType=="INRETS":
                Data=[Container.AttValue("inrets_a"),
                      Container.AttValue("capacityFactor")]    
                
            if CrType=="LOGISTIC":
                Data=[Container.AttValue("logistic_a"),
                      Container.AttValue("logistic_b"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("logistic_d"),
                      Container.AttValue("logistic_f")]
            
            if CrType=="QUADRATIC":
                Data=[Container.AttValue("quadratic_a"),
                      Container.AttValue("quadratic_b"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("quadratic_d")]
               
            if CrType=="SIGMOIDAL_MMF_LINKS":
                Data=[Container.AttValue("sigmoidal_mmf_a"),
                      Container.AttValue("sigmoidal_mmf_b"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("sigmoidal_mmf_d"),
                      Container.AttValue("sigmoidal_mmf_f")]
                  
            if CrType=="SIGMOIDAL_MMF_NODES":
                Data=[Container.AttValue("sigmoidal_mmf_a"),
                      Container.AttValue("sigmoidal_mmf_b"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("sigmoidal_mmf_d"),
                      Container.AttValue("sigmoidal_mmf_f")]
                 
            if CrType=="Akcelik":
                Data=[Container.AttValue("akcelik_a"),
                      Container.AttValue("akcelik_b"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("akcelik_d")]
                 
            if CrType=="Lohse":
                Data=[Container.AttValue("lohse_a"),
                      Container.AttValue("lohse_b"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("lohse_satcrit")]
              
            if CrType=="Linear bottle-neck":
                Data=[Container.AttValue("capacityFactor")]
                
            if CrType=="Akcelik2":
                Data=[Container.AttValue("akcelik2_a"),
                      Container.AttValue("akcelik2_b"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("akcelik2_d")]
                
            if CrType=="TMODEL_LINKS":
                Data=[Container.AttValue("tmodel_a1"),
                      Container.AttValue("tmodel_a2"),
                      Container.AttValue("tmodel_b1"),
                      Container.AttValue("tmodel_b2"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("tmodel_d1"),
                      Container.AttValue("tmodel_d2"),
                      Container.AttValue("tmodel_f1"),
                      Container.AttValue("tmodel_f2"),
                      Container.AttValue("tmodel_satcrit")]
                
            if CrType=="TMODEL_NODES":
                Data=[Container.AttValue("tmodel_a1"),
                      Container.AttValue("tmodel_a2"),
                      Container.AttValue("tmodel_b1"),
                      Container.AttValue("tmodel_b2"),
                      Container.AttValue("capacityFactor"),
                      Container.AttValue("tmodel_d1"),
                      Container.AttValue("tmodel_d2"),
                      Container.AttValue("tmodel_f1"),
                      Container.AttValue("tmodel_f2"),
                      Container.AttValue("tmodel_satcrit")]
            
            return Data
    CRnos={}
    k=0
    CR={}
    for Cr in Crs:
        CrType=str(Visum.Procedures.Functions.CrFunctions.CrFunction(Cr[1]).AttValue("CrFunctionType"))
        CRnos[k]=CrType
        k+=1
        try:            
            CR[CrType]
        except:         
            CR[CrType]=AddCR(CrType,Visum.Procedures.Functions.CrFunctions.CrFunction(Cr[1]))
    return CR,CRnos
            
def DumpToVisum(Vols):
    Param="i2_inverse_VDF_Vol"
    try: 
        Visum.Net.Links.GetMultiAttValues(Param)  #check if UDAs exists
    except:    
        Visum.Net.Links.AddUserDefinedAttribute(Param, Param, Param, 2,0) #else create UDAs
    Visum.Net.Links.SetMultiAttValues(Param,Vols)

    

    

def GetData():
    typ=1
    Crs=Visum.Net.Links.GetMultiAttValues("CrNo",True)
    typ=win32api.MessageBox(None,
                    "Welcome to Inverse VDF by \nwww.intelligent-infrastructure.eu\n\nIn the next step, You will be asked to specify input parameter for Links.\n",
                    "Inverse VDF by www.intelligent-infrastructure.eu",win32con.MB_ICONINFORMATION)
    Inputs=Visum.Net.Links.AskAttribute()
    Inputs=Visum.Net.Links.GetMultiAttValues(Inputs,True)
    typ=win32api.MessageBox(None,
                        "Is your Input data a speed data (YES), or travel time data (NO) \n",
                        "Type of input data",win32con.MB_YESNO | win32con.MB_ICONQUESTION)
    
    if typ==7:
         T0s=Inputs
    else:
        Lengths=Visum.Net.Links.GetMultiAttValues("Length",True)
        j=1
        T0s=[]
        for length in Lengths:
           T0s.append([j,3600*Lengths[j-1][1]/Inputs[j-1][1]]) 
           j+=1
    Times=Visum.Net.Links.GetMultiAttValues("TCUR_PRTSYS(C)",True)
    Caps=Visum.Net.Links.GetMultiAttValues("CapPrT",True)
    Vols=Visum.Net.Links.GetMultiAttValues("VolVehPrT(AP)",True)
    No=Visum.Net.Links.GetMultiAttValues("No",True)
    j=0
    CR,Crnos=Make_CR_Dict(Crs)
    NewVols=[]
    for Time in Times:        
        NewVols.append([j+1,GetVol(CR,Crnos[j],Time[1],Caps[j][1],T0s[j][1],Vols[j][1])])
        j+=1
    DumpToVisum(NewVols)
    typ=win32api.MessageBox(None,
                    "Done, \nresults saved as "+Param,
                    "Inverse VDF by www.intelligent-infrastructure.eu",
                    win32con.MB_ICONINFORMATION) 

GetData()
Visum.SaveVersion("D:/abc.ver")
  