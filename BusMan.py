
def Create_Stops(no,Node):
    XCoord=Visum.Net.Nodes.ItemByKey(Node).AttValue("XCoord")
    YCoord=Visum.Net.Nodes.ItemByKey(Node).AttValue("YCoord")
    Visum.Net.AddStop(no,XCoord,YCoord)
    Visum.Net.AddStopArea(no, no, Node, XCoord, YCoord)
    Visum.Net.AddStopPointOnNode(no, no, Node)
    
def Add_Operator(Name):
    no=len(Visum.Net.Operators.GetMultiAttValues("Name", False))+1
    Visum.Net.AddOperator(no)
    Visum.Net.Operators.ItemByKey(no).SetAttValue("Name",Name)
    
def Add_LineRoute(Name,LineName,StopPoints,Backward):
    
    RouteSearchParams = Visum.CreateNetReadRouteSearchTSys()
    RouteSearchParams.SearchShortestPath(1, True, True, 2, 1, 2) 
    
    if Backward == False:
        Route= Visum.CreateNetElements()
        for StopPoint in StopPoints:
            Route.Add(Visum.Net.StopPoints.ItemByKey(StopPoint))            
        Visum.Net.AddLineRoute(Name, LineName, direction1, Route, RouteSearchParams)
            
    else:
        RouteForward = Visum.CreateNetElements()
        RouteBackward = Visum.CreateNetElements()
        for i in range(len(StopPoints)):
            RouteForward.Add(Visum.Net.StopPoints.ItemByKey(StopPoints[i]))            
            RouteBackward.Add(Visum.Net.StopPoints.ItemByKey(StopPoints[-i-1]))
        Visum.Net.AddLineRoute(Name, LineName, direction1, RouteForward, RouteSearchParams)
        Visum.Net.AddLineRoute(Name, LineName, direction2, RouteBackward, RouteSearchParams)
            
    

        
    
    
    
    
    
Create_Stops(1,1)
Create_Stops(2,7)
Create_Stops(3,61)
Create_Stops(4,81)

Visum.Net.AddTSystem("Autobus", "PUT")
Visum.Net.AddTSystem("Trawmaj", "PUT")
Add_Operator("MPK")

Visum.Net.AddLine("179","B")


direction1 = Visum.Net.Directions.ItemByKey(">")
direction2 = Visum.Net.Directions.ItemByKey("<")

Add_LineRoute("179","179",[1,2,3,4],True)







    
