def Read_CSV(): 
        fields=["Process"
                "SIGR",
                "SDAY",
                "DAYT",
                "SimInitHHMMSS",
                "SimSpan",
                "SimIntD",
                "OWMQ",
                "EquAss",
                "DTAnITE",
                "WarmIte",
                "DtaInitHHMMSS",
                "DtaSpan",
                "DtaIntD",
                "StopBlockTime",
                "MoveBlockTime",
                "HardEscapeBlock",
                "SneakPerLane",
                "AllowSneaking",
                "SplitModelCoef",
                "ElasticDemand",
                 "ExtendDemandProfiles",
                "BaseSplitOnOriginsDTA",
                "KeepFlowOnOrigins",
                "TPWD",
                "SplitBin",
                "ExportSplit",
                "MAX_NUM_POINTS",
                "SIMnPTH",
                "DTAnPTH",
                "CostResInterpolation",
                "ANGS",
                "EXTC",
                "MinPathSpeed",
                "MaxRam",
                "TURR",
                "ResIntD",
                "ResByInterval"]             
        csvfile=open('C:\TRE_cmdl.csv', 'rb')
        param={}
        while 1:
            line = csvfile.readline()            
            if not line:
                break           
            
            newline = line.rstrip('\r\n')
            newline=newline.replace('"','')
            if newline in fields:                                
                param[newline]=csvfile.readline().rstrip('\r\n')
        return param
    
print Read_CSV()
