#!/usr/bin/env python
# -*- coding: utf-8 -*-

lsnap = 100
i = 0
Iterator = Visum.Net.Stops.Iterator
while Iterator.Valid:
    przystanek = Iterator.Item
    i += 1
    stopID = przystanek.AttValue("No")
    if int(przystanek.AttValue("Name")[-3:])  #sprawdź, czy kończy się na cyfrę, czy nie oddzielenie 'Plac Starzynskiego' od 'Plac Starzynskiego 03':
        xS = przystanek.AttValue("XCoord")
        yS = przystanek.AttValue("YCoord")
        StopArea = Visum.Net.AddStopArea(stopID, przystanek, Node, xS, yS)
    StopPoint = Visum.Net.AddStopPointOnNode(stopID, StopArea, Node)
    [Link, res, dist, xL, yL, rel] = Visum.Net.GetNearestLink(xS, yS, lsnap, False)
    if res:
        StopPoint = Visum.Net.AddStopPointOnLink(stopID,StopArea,Link.AttValue("FromNodeNo"),Link.AttValue("ToNodeNo"),True)
    else:
        print "nie znalazl odcinka dla przystanku" + str(stopID)

    Iterator.Next()

Iterator = Visum.Net.Stops.Iterator
while Iterator.Valid:
    przystanek = Iterator.Item
    i += 1
    stopID = przystanek.AttValue("No")
    xS = przystanek.AttValue("XCoord")
    yS = przystanek.AttValue("YCoord")
    StopArea = Visum.Net.AddStopArea(stopID, przystanek, Node, xS, yS)
    StopPoint = Visum.Net.AddStopPointOnNode(stopID, StopArea, Node)
    [Link, res, dist, xL, yL, rel] = Visum.Net.GetNearestLink(xS, yS, lsnap, False)
    if res:
        StopPoint = Visum.Net.AddStopPointOnLink(stopID,StopArea,Link.AttValue("FromNodeNo"),Link.AttValue("ToNodeNo"),True)
    else:
        print "nie znalazl odcinka dla przystanku" + str(stopID)

    Iterator.Next()

Iterator = Visum.Net.Stops.Iterator
while Iterator.Valid:



