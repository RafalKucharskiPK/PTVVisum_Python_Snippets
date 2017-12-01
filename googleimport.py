# -*- coding: iso-8859-15 -*-
'''
VISUM add-in Import Google Transit Feed

Internationalisation for German and English has been implemented.

Notice:
* for debugging the add-in please set VISUMINTEGRATED = False

Add-in description: Help.htm / Hilfe.htm

Date: 31.07.2009
Author: Klaus Nökel, Yvonne Huebner
Contact: klaus.noekel@ptv.de, yvonne.huebner@ptv.de
Company: PTV AG'''
import csv
import os
import os.path
import wx
from VisumPy.helpers import GetMulti, SetMulti, HHMMSS2secs, secs2HHMMSS
from datetime import date
from cPickle import loads
#import win32com.client as com # debug

# TODO: Unicode strings / UTF-8 files
# TODO: Shapes
# TODO: Blocks

# Idea: TSys filter
# Idea: Trip ID and / or service ID filter (or date filter)


class MyDialect(csv.excel):
    skipinitialspace = True


def ReportError(msg, exception):
        #Reports an error in a message box with the 'msg' entry.#
    wx.MessageBox(message = msg + " " + str(exception),
                  caption =_("Error"),
                  style=wx.ICON_ERROR|wx.OK)

def convert_ddmmyyyy2yyyymmdd(ddmmyyyy):
        # Convert the date from format ddmmyyyy to yyyymmdd
        # as it is used in googleimport.py
        ddmmyyyy = ddmmyyyy.replace('.', '')
        day = ddmmyyyy[:2]
        month = ddmmyyyy[2:4]
        year= ddmmyyyy[4:]
        return year + month + day

def makeDate(yyyymmdd):
    year = int(yyyymmdd[:4])
    month = int(yyyymmdd[4:6])
    day = int(yyyymmdd[6:])
    return date(year,month,day)

def makeSafeString(s):
    """ return a string that is safe as an attribute value in VISUM. Replace
    illegal characters with underscore. """
    s = s.replace("$", "_")
    s = s.replace(";", "_")
    s = s.replace(")", "_")
    s = s.replace("(", "_")
    return s

class UserAbort:
    pass

class LRCache:
    """ a class that buffers lineroute / timeprofile definitions. When a lineroute
    is added to the cache it is automatically compared to previously stored entries.
    Duplicates are not stored, instead a dictionary maps original IDs to the ID
    of the representative which is retained. Only representatives are created in
    VISUM and subsequent references from vehjourneys etc. should be redirected to
    the representative. """


    def __init__(self):
        self.hashTP = dict() # hash value --> [TP definition]
        self.hashLR = dict() # hash value --> [LR definition]
        self.ReprTP = dict() # tripID --> TP name
        self.ReprLR = dict() # tripID --> lrkeys

    def Add(self, tripID, tpkeys, lri, tpi):

        def findTPRepr(entry, entries):
            tpkeys, lri, tpi = entry
            reprkeys = None
            for x in entries:
                tpkeys1, lri1, tpi1 = x
                if tpkeys[0] == tpkeys1[0] and lri == lri1 and tpi == tpi1:
                    reprkeys = tpkeys1
                    break
            return reprkeys

        def findLRRepr(entry, entries):
            tpkeys, lri, tpi = entry
            reprkeys = None
            for x in entries:
                tpkeys1, lri1, tpi1 = x
                if tpkeys[0] == tpkeys1[0] and lri == lri1:
                    reprkeys = tpkeys1
                    break
            return reprkeys

        def calchash(lri):
            hash = sum(map(sum,lri))
            return hash % 1000

        hash = calchash(lri)
        entry = (tpkeys, lri, tpi)
        if hash not in self.hashLR:
            self.hashLR[hash] = [entry]
            self.hashTP[hash] = [entry]
            self.ReprLR[tripID] = tpkeys
            self.ReprTP[tripID] = tpkeys
        else:
            reprkeys = findLRRepr(entry, self.hashLR[hash])
            if reprkeys == None:
                self.hashLR[hash].append(entry)
                self.hashTP[hash].append(entry)
                self.ReprLR[tripID] = tpkeys
                self.ReprTP[tripID] = tpkeys
            else:
                self.ReprLR[tripID] = reprkeys
                newkeys = (reprkeys[0], reprkeys[1], reprkeys[2], tpkeys[3])
                entry = (newkeys, lri, tpi)
                reprTP = findTPRepr(entry, self.hashTP[hash])
                if reprTP == None:
                    self.hashTP[hash].append(entry)
                    self.ReprTP[tripID] = newkeys
                else:
                    self.ReprTP[tripID] = reprTP

    def NumTripIDs(self):
        return len(self.ReprLR.keys())

    def NumReprTPs(self):
        return sum(map(len, self.hashTP.itervalues()))

    def NumReprLRs(self):
        return sum(map(len, self.hashLR.itervalues()))

    def GetRepresentativeLR(self, id):
        return self.ReprLR[id]

    def GetRepresentativeTP(self, id):
        return self.ReprTP[id]


class GoogleReader:

    def ProcessCalendar(self, calendarfile, datesfile, useDate, filterDate):
        keepGoing = self.progdlg.Update(0, _("Processing calendar"))

        if not keepGoing[0]: # continue?
            self.progdlg.Update(100) # also necessary
            return -1

        self.dictDay = dict() # serviceid --> validdayno
        self.dictBitvec = dict() # serviceid --> bitvec

        # first pass only to determine calendar period = [min date; max date]
        mindate = date.max
        maxdate = date.min

        if calendarfile != None:
            f = open(calendarfile,"rb")
            reader = csv.DictReader(f, dialect=MyDialect)
            for record in reader:
                maxdate = max(maxdate, makeDate(record["end_date"]))
                mindate = min(mindate, makeDate(record["start_date"]))
            f.close()

        if datesfile != None:
            f = open(datesfile,"rb")
            reader = csv.DictReader(f, dialect=MyDialect)
            for record in reader:
                thedate = makeDate(record["date"])
                maxdate = max(maxdate, thedate)
                mindate = min(mindate, thedate)
            f.close()

        # second pass to define the validdays
        minord = mindate.toordinal()
        maxord = maxdate.toordinal()
        periodlen = maxord - minord + 1
        never = "0" * periodlen

        if calendarfile != None:
            f = open(calendarfile,"rb")
            reader = csv.DictReader(f, dialect=MyDialect)
            for record in reader:
                enddate = makeDate(record["end_date"])
                startdate = makeDate(record["start_date"])
                endord = enddate.toordinal()
                startord = startdate.toordinal()
                startweekday = startdate.weekday() # Monday = 0, Tuesday = 1, ...
                bitvec = list(never) # start with empty vector
                weekvec = []
                for wday in ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]:
                    weekvec.append(record[wday])
                windex = startweekday
                for i in xrange(startord-minord, endord+1-minord):
                    bitvec[i] = weekvec[windex]
                    windex = (windex + 1) % 7
                self.dictBitvec[record["service_id"]] = bitvec
            f.close()

        if datesfile != None:
            f = open(datesfile,"rb")
            reader = csv.DictReader(f, dialect=MyDialect)
            lastserviceid = ""
            for record in reader:
                serviceid = record["service_id"]
                if serviceid != lastserviceid:
                    if lastserviceid != "":
                        self.dictBitvec[lastserviceid] = bitvec
                    bitvec = self.dictBitvec.get(serviceid, list(never))
                    lastserviceid = serviceid
                thedate = makeDate(record["date"])
                theord = thedate.toordinal()
                bitvec[theord-minord] = record["exception_type"]
            f.close()
            self.dictBitvec[lastserviceid] = bitvec

        # create CalendarPeriod + ValidDays
        calperiod = Visum.Net.CalendarPeriod
        calperiod.AttValue("TYPE")  # dummy statement added to avoid failure in next line
        calperiod.SetAttValue("TYPE",3)
        calperiod.SetAttValue("VALIDUNTIL","31.12.2029") # to avoid range check failure
        calperiod.SetAttValue("VALIDFROM", mindate.strftime("%d.%m.%Y"))
        calperiod.SetAttValue("VALIDUNTIL",maxdate.strftime("%d.%m.%Y"))
        
        keyNumbers = GetMulti(Visum.Net.ValidDaysCont,"No")
        highestNumber = 1;
        for i in range(len(keyNumbers)):
            if keyNumbers[i]> highestNumber:
                highestNumber = keyNumbers[i]
        
        no = highestNumber + 1  # start with 2 (when this is the first day that will be added into VISUM) to avoid conflict with default "daily"

        for serviceid, bitvec in self.dictBitvec.iteritems():
            vday = Visum.Net.AddValidDays(no)
            vday.SetAttValue("CODE", makeSafeString(serviceid))
            vday.SetAttValue("NAME", makeSafeString(serviceid))
            vday.SetAttValue("DAYVECTOR", "".join(bitvec))
            self.dictDay[serviceid] = no
            no += 1

        # calculate filteridx of bit vector
        if (useDate):
            filterDateConv = makeDate(filterDate)
            dateord = filterDateConv.toordinal()
            if minord <= dateord and dateord <= maxord:
                self.filteridx = dateord - minord
            else:
                ReportError(_("The filter date %s is not in the range of start %s and end date %s: ") % (str(filterDateConv), str(mindate), str(maxdate)), "")
                return -1
        else:
            self.filteridx = -1
            return 1

    def ProcessStops(self, stopfile): ### dodawanie stops
        keepGoing = self.progdlg.Update(0, _("Processing stops")) ### gui ...

        if not keepGoing[0]: # continue?
            self.progdlg.Update(100) # also necessary
            return -1

        filesize = os.stat(stopfile).st_size

        self.dictStop = dict() # stopid --> no
        stops = Visum.Net.Stops ## slownik id -> nos
        try:
            stops.AddUserDefinedAttribute("FSZNO","FSZNo","FareZoneNo",5) # string
            stops.AddUserDefinedAttribute("ORIG_ID","Orig_ID","Original ID",5) # string
        except:
            pass
        
        f = open(stopfile,"rb")
        reader = csv.DictReader(f, dialect = MyDialect)
        
        loop = 0       
        keyNumbers = GetMulti(Visum.Net.Stops,"No")
        highestNumber = 0;
        
        for i in range(len(keyNumbers)):
            if keyNumbers[i]> highestNumber:
                highestNumber = keyNumbers[i]

        # is set to true only if at least one new stop point will be created
        # (if anything changed we spare some VISUM operations done at the end of this procedure)
        anythingChanged = False
                        
        for record in reader:
            loop += 1
            if (loop % 100 == 0):
                cont,skip = self.progdlg.Update(min(99, f.tell() * 100 / filesize))
                if not cont: return # do not return -1 ! it won't work.
                
            id = record["stop_id"]
            
            # Is set to true if a stop point with same number already exists in VISUM 
            #(no new stop point will be created)
            stopAlreadyExisted = False
            
            
            try:
                no = int(id)
                if no in keyNumbers:              # check if stopPoint alreadyExists
                    stopAlreadyExisted = True     # nothing else to do 
                else:
                    stop = Visum.Net.AddStop(no)  #create new stop point
                    anythingChanged = True        
                    keyNumbers.append(no)            # add the number to list with key numbers
                    if no > highestNumber:        # update highestNumber if needed
                        highestNumber = no
            except:
                no = highestNumber + 1
                highestNumber = no
                stop = Visum.Net.AddStop(no)  
                anythingChanged = True 
                stopAlreadyExisted = False
                
            self.dictStop[id] = no

            # set values of new stopPoint (if the stoppoint already existed nothing will be changed)
            if not stopAlreadyExisted:
                stop.SetAttValue("NAME", makeSafeString(record["stop_name"]))
                if "stop_code" in record:
                    stop.SetAttValue("CODE", makeSafeString(record["stop_code"]))
                else:
                    stop.SetAttValue("CODE", makeSafeString(id))
                stop.SetAttValue("XCOORD", record["stop_lon"])
                stop.SetAttValue("YCOORD", record["stop_lat"])
                stop.SetAttValue("ORIG_ID", makeSafeString(record["stop_id"]))
                if "zone_id" in record:
                    stop.SetAttValue("FSZNO", makeSafeString(record["zone_id"]))
    
                node = Visum.Net.AddNode(no)
                stopArea = Visum.Net.AddStopArea(no, stop, node)
                stopPoint = Visum.Net.AddStopPointOnNode(no, stopArea, node)
                
        if anythingChanged:   
            xcoord = GetMulti(Visum.Net.Stops, "XCOORD")
            ycoord = GetMulti(Visum.Net.Stops, "YCOORD")
            SetMulti(Visum.Net.Nodes, "XCOORD", xcoord)
            SetMulti(Visum.Net.Nodes, "YCOORD", ycoord)
            SetMulti(Visum.Net.StopAreas, "XCOORD", xcoord)
            SetMulti(Visum.Net.StopAreas, "YCOORD", ycoord)
            SetMulti(Visum.Net.StopPoints, "CODE", GetMulti(Visum.Net.Stops, "CODE"))
            SetMulti(Visum.Net.StopPoints, "NAME", GetMulti(Visum.Net.Stops, "NAME"))


    def ProcessAgency(self, agencyfile):
        
        self.dictAgencyNo = dict() # agency id --> operator no
        reader = csv.DictReader(open(agencyfile,"rb"), dialect = MyDialect)
        keyNumbers = GetMulti(Visum.Net.Operators,"No")
        highestNumber = 0;
        for i in range(len(keyNumbers)):
            if keyNumbers[i] > highestNumber:
                highestNumber = keyNumbers[i]
        
        no = highestNumber + 1
        
        for record in reader:
            op = Visum.Net.AddOperator(no)
            op.SetAttValue("NAME", makeSafeString(record["agency_name"]))
            if "agency_id" in record:
                self.dictAgencyNo[record["agency_id"]] = no
            no += 1

    def CreateTSys(self):
        """ create the predefined Google Transit tsys"""
        googleTSys = [("0", "Tram"),
                      ("1", "Subway"),
                      ("2", "Rail"),
                      ("3", "Bus"),
                      ("4", "Ferry"),
                      ("5", "Cable car"),
                      ("6", "Gondola"),
                      ("7", "Funicular")]
        self.dictTSys = dict() # Google route type --> VISUM tsys object
        for code, name in googleTSys:
            try:
                tsys = Visum.Net.AddTSystem(code, "PUT")
                tsys.SetAttValue("NAME", name)
            except:
                tsys = Visum.Net.TSystems.ItemByKey(code)
            self.dictTSys[code] = tsys
            

    def ProcessLines(self, routefile):
        keepGoing =self.progdlg.Update(0, _("Processing routes"))
        if not keepGoing[0]: # continue?
            self.progdlg.Update(100)
            return -1

        filesize = os.stat(routefile).st_size

        self.dictLine = dict()
        f = open(routefile,"rb")
        reader = csv.DictReader(f, dialect = MyDialect)
        try:
            Visum.Net.Lines.AddUserDefinedAttribute("SHORTNAME","ShortName","ShortName",5) # string
            Visum.Net.Lines.AddUserDefinedAttribute("LONGNAME","LongName","LongName",5) # string
        except:
            pass
        
        loop = 0
        for record in reader:
            loop += 1
            if (loop % 10 == 0):
                cont,skip = self.progdlg.Update(min(99, f.tell() * 100 / filesize))
                if not cont: return # do not return -1 ! it won't work.
            try:
                line = Visum.Net.AddLine(makeSafeString(record["route_id"]), self.dictTSys[record["route_type"]])
                self.dictLine[record["route_id"]] = line
                line.SetAttValue("SHORTNAME", makeSafeString(record["route_short_name"]))
                line.SetAttValue("LONGNAME", makeSafeString(record["route_long_name"]))
                opNo = self.dictAgencyNo.get(record.get("agency_id",""),0)
                line.SetAttValue("OPERATORNO", opNo)
            except:
                pass

    def ProcessTrips(self, tripfile, stoptimefile, freqfile, useDate):
        keepGoing = self.progdlg.Update(0, _("Processing trips"))
        if not keepGoing[0]: # continue?
            self.progdlg.Update(100) # also necessary
            return -1

        filesize = os.stat(stoptimefile).st_size

        self.dictTP = dict() # ???
        dictDep = dict() # tripid --> departures
        dictNumItems = dict() # tripid --> highest item index
        self.dictTrip = dict() # tripid --> trip file record
        reader = csv.DictReader(open(tripfile,"rb"), dialect=MyDialect)
        for record in reader:
            self.dictTrip[record["trip_id"]] = record

        LRI = []
        TPI = []
        laststop = -1
        f = open(stoptimefile,"rb")
        reader = csv.DictReader(f, dialect=MyDialect)
        curTripID = ""
        self.cache = LRCache()
        loop = 0
        for record in reader:
            loop += 1
            if (loop % 100 == 0):
                cont,skip = self.progdlg.Update(min(99, f.tell() * 100 / filesize))
                if not cont: return

            if record["trip_id"] != curTripID:
                if curTripID != "":
                    dictNumItems[curTripID] = index-1 # previous trip!
                    self.cache.Add(curTripID, lrkeys, LRI, TPI)
                curTripID = record["trip_id"]
                dictDep[curTripID] = [record["departure_time"]]
                startDep = HHMMSS2secs(record["departure_time"])
                trip = self.dictTrip[curTripID]
                dir = trip.get("direction_id", 0)
                if dir == "":
                    dir = 0
                else:
                    dir = int(dir)
                dircode = "<>"[dir]
                lrkeys = (makeSafeString(trip["route_id"]), makeSafeString(curTripID), dircode, makeSafeString(curTripID))
                index = 1
                LRI = []
                TPI = []
                laststop = -1

            stop = self.dictStop[record["stop_id"]]
            if stop == laststop:
                # some GTF raw data include the
                # same stop  multiple times in succession. In these cases ignore all but the first occurrence
                continue
            laststop = stop

            if record["arrival_time"] == "":
                arr = lastDep
            else:
                arr = secs2HHMMSS(max(0, HHMMSS2secs(record["arrival_time"]) - startDep))
            if record["departure_time"] == "":
                dep = arr
            else:
                dep = secs2HHMMSS(max(0, HHMMSS2secs(record["departure_time"]) - startDep))

            lastDep = dep
            if record.get("pickup_type", "0") != "1" or TPI == []:
                board = 1
            else:
                board = 0
            if record.get("drop_off_type", "0") != "1":
                alight = 1
            else:
                alight = 0
            # TODO: postlength
            postlength = 0
            LRI.append([index, stop, stop, 1, postlength])
            TPI.append([index, alight, board, arr, dep, index])

            index += 1
        if curTripID != "":
            dictNumItems[curTripID] = index-1 # previous trip!
            self.cache.Add(curTripID, lrkeys, LRI, TPI)

        # process frequencies
        if freqfile != None:
            dictFreq = dict() # tripid --> [dep]
            f = open(freqfile,"rb")
            reader = csv.DictReader(f, dialect=MyDialect)
            for record in reader:
                tripid = record["trip_id"]
                starttime = record["start_time"]
                endtime = record["end_time"]
                headway = int(record["headway_secs"])
                depsec = range(HHMMSS2secs(starttime), HHMMSS2secs(endtime)+1, headway)
                deps = map(secs2HHMMSS, depsec)
                dictFreq[tripid] = dictFreq.get(tripid, []) + deps
            f.close()
            dictDep.update(dictFreq) # frequencies win, if present

        # write net file
        keepGoing = self.progdlg.Update(0, _("Write network file"))
        if not keepGoing[0]: # continue?
            self.progdlg.Update(100) # also necessary
            return -1

        #netname = r"D:\noe\Develop\Visum Misc\GoogleTransit\test.net" # for debugging purposes
        netname = os.tempnam()+".net"
        f = open(netname, "wb")
        writer = csv.writer(f, delimiter=";")
        # header
        f.writelines(["$VISION\n",
                      "$VERSION:VERSNR;FILETYPE;LANGUAGE;UNIT\n",
                      "3.000;Net;E;KM\n"])
        f.write("$LINEROUTE:LINENAME;NAME;DIRECTIONCODE;ISCIRCLELINE\n")
        for values in self.cache.hashLR.itervalues():
            for lr in values:
                x = [lr[0][0], lr[0][1], lr[0][2], 0] # not tp name!
                writer.writerow(x)
        f.write("$LINEROUTEITEM:LINENAME;LINEROUTENAME;DIRECTIONCODE;INDEX;NODENO;STOPPOINTNO;ISROUTEPOINT;POSTLENGTH\n")
        for values in self.cache.hashLR.itervalues():
            for lr in values:
                lrkeys = [lr[0][0], lr[0][1], lr[0][2]] # not tp name!
                for item in lr[1]:
                    writer.writerow(lrkeys + item)
        f.write("$TIMEPROFILE:LINENAME;LINEROUTENAME;DIRECTIONCODE;NAME\n")
        writer.writerows(self.cache.ReprTP.itervalues())
        f.write("$TIMEPROFILEITEM:LINENAME;LINEROUTENAME;DIRECTIONCODE;TIMEPROFILENAME;INDEX;ALIGHT;BOARD;ARR;DEP;LRITEMINDEX\n")
        for values in self.cache.hashTP.itervalues():
            for tp in values:
                tpkeys = list(tp[0])
                for item in tp[2]:
                    writer.writerow(tpkeys + item)
        f.write("$VEHJOURNEY:NO;NAME;DEP;LINENAME;LINEROUTENAME;DIRECTIONCODE;TIMEPROFILENAME;FROMTPROFITEMINDEX;TOTPROFITEMINDEX\n")
        no = 1
        for tripid, deps in dictDep.iteritems():
            tpkeys = self.cache.GetRepresentativeTP(tripid)
            trip = self.dictTrip[tripid]
            serviceid = trip["service_id"]
            if self.CheckFilter(useDate, serviceid) == True:
                for dep in deps:
                    x = [no, tripid, dep] + list(tpkeys) + [1, dictNumItems[tripid]]
                    writer.writerow(x)
                    no += 1
            else:
                continue

        f.write("$VEHJOURNEYSECTION:VEHJOURNEYNO;NO;VALIDDAYSNO;FROMTPROFITEMINDEX;TOTPROFITEMINDEX\n")
        no = 1
        for tripid, deps in dictDep.iteritems():
            tpkeys = self.cache.GetRepresentativeTP(tripid)
            # Calculate
            trip = self.dictTrip[tripid]
            serviceid = trip["service_id"]
            if self.CheckFilter(useDate, serviceid) == True:
                vdayno = self.dictDay[serviceid]
                for dep in deps:
                    x = [no, no, vdayno, 1, dictNumItems[tripid]]
                    writer.writerow(x)
                    no += 1
            else:
                continue

        f.close()

        # read net
        keepGoing = self.progdlg.Update(0, _("Load network file into VISUM"))
        if not keepGoing[0]: # continue?
            self.progdlg.Update(100) # also necessary
            return -1

        searchparatsys = Visum.CreateNetReadRouteSearchTSys()
        searchparatsys.InsertOrOpenLink(99)
        #searchparatsys.SearchShortestPath(0, True, True, 99, 2, 0)
        searchpara = Visum.CreateNetReadRouteSearch()
        searchpara.SetForAllTSys(searchparatsys)
        Visum.LoadNet(netname, True, searchpara)
        os.remove(netname)

        # set veh journey operator no = line operator no
        ops = GetMulti(Visum.Net.VehicleJourneys, "TimeProfile\\LineRoute\\Line\\OperatorNo")
        SetMulti(Visum.Net.VehicleJourneys, "OperatorNo", ops)

    def SetProjection(self):
        """ set projection to WGS84 """
        WKT_WGS84 = r'GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137,298.257223563]],PRIMEM["Greenwich",0],UNIT["Degree",0.017453292519943295]]'
        Visum.Net.NetParameters.SetProjection(WKT_WGS84, False)

    def SetLinkAttributes(self):
        """ set all link lengths and walk times to an arbitrary value > 0. """
        Visum.Net.Links.SetAllAttValues("LENGTH", 1)
        for tsys in Visum.Net.TSystems:
            if tsys.AttValue("TYPE") == "PUTWALK":
                Visum.Net.Links.SetAllAttValues("t_PUTSYS(%s)" % tsys.AttValue("CODE"), 1800) # 30min

    def CheckFilter(self, filterActive, serviceid):
        """ checks if a datefilter is active and if so checks
        if the service takes place on the given date """
        if not filterActive:
            return True
        else:
            bitvec = self.dictBitvec[serviceid]
            if self.filteridx == -1:
                return False
            else:
                # entry is a string!
                return bitvec[self.filteridx] == '1'

    def Main(self, path, useDate, filterDate):
        self.progdlg = wx.ProgressDialog("Google Transit import", _("Initializing"),
                                         parent=None, style=wx.PD_APP_MODAL | wx.PD_CAN_ABORT | wx.PD_AUTO_HIDE)
        #global Visum
        #Visum = com.Dispatch("Visum.Visum.110") #debug
        self.SetProjection()
        self.CreateTSys()
        agencyfile = os.path.join(path, "agency.txt")
        self.ProcessAgency(agencyfile)
        datesfile = os.path.join(path, "calendar_dates.txt")
        if not os.path.exists(datesfile):
            datesfile = None
        calendarfile = os.path.join(path, "calendar.txt")
        if not os.path.exists(calendarfile):
            calendarfile = None
        if self.ProcessCalendar(calendarfile, datesfile, useDate, filterDate) == -1:
            self.progdlg.Destroy()
            return
        stopfile = os.path.join(path, "stops.txt")
        if self.ProcessStops(stopfile) == -1:
            self.progdlg.Destroy()
            #ReportError("in processStops", "")#debug
            return
        #ReportError("After processStops", "") #debug
        routefile = os.path.join(path, "routes.txt")
        if self.ProcessLines(routefile) == -1:
            self.progdlg.Destroy()
            return
        tripfile = os.path.join(path, "trips.txt")
        stoptimefile = os.path.join(path, "stop_times.txt")
        freqfile = os.path.join(path, "frequencies.txt")
        if not os.path.exists(freqfile):
            freqfile = None
        if self.ProcessTrips(tripfile, stoptimefile, freqfile, useDate) == -1:
            self.progdlg.Destroy()
            return
        self.SetLinkAttributes()
        self.progdlg.Destroy()

def Run():
    defaultParam = { "SourceDirectory": r"D:\Projekte\2_Spec_GoogleTransitFeedImport\GoogleTransit\Austin_feed",
                     "UseFilterDate": True,
                     "FilterDate": "30.01.2007"}

    param = defaultParam

    if Parameter.Data != "":
        tmppara = Parameter.Data.encode("iso-8859-15")
        storedparam = loads(tmppara)
        param.update(storedparam)

    importer = GoogleReader()
    convertDate = convert_ddmmyyyy2yyyymmdd(param["FilterDate"])
    importer.Main(param["SourceDirectory"], param["UseFilterDate"], convertDate)
    Visum.Graphic.DisplayEntireNetwork()


import gettext
gettext.install("GoogleTransit",".\\locale")
languagemap = { "DEU" : 'de_de' }
code = languagemap.get(Visum.GetCurrentLanguage(), 'en_us') # if language not available, use English
translation = gettext.translation("GoogleTransit", "./locale", languages=[code])
translation.install()

if Parameter.OK:
    wx.InitAllImageHandlers()
    Run()
