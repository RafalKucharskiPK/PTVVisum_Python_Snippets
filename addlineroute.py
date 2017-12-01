#-*- coding: utf-8 -*-

def VisumInit(path=None):
    """       
    VISUM INIT
    """
    import win32com.client        
    Visum = win32com.client.Dispatch('Visum.Visum.125')
    if path != None: Visum.LoadVersion(path)
    return Visum

    
                    
def Add_Line():
    
    routesearchparameters = Visum.CreateNetReadRouteSearchTSys() # RK: I assume it's needed for proper handling of AddLineRoute
    d=Visum.CreateNetReadRouteSearch() #RK: Is it actually neccessary?
    d.SetForTSys("B",routesearchparameters) #RK: Is it actually neccessary? - should I put it below or above \/
    routesearchparameters.SearchShortestPath(3,#    SearchShortestPath ( [in] Enum ShortestPathCriterionT, 
                                             True,#[in] VARIANT_BOOL includeBlockedLinksInRouting, 
                                             True,#[in] VARIANT_BOOL includeBlockedTurnsInRouting, 
                                             2,#[in] double MaxDeviationFactor, #RK: How to interpret it? Is it max(LineRouteLength/DirectLineRouteDist)
                                             2,#[in] Enum DoIfNotFound, #RK: Procedure works only when 'DoIfNotFound' == 2 ! 
                                             99#[in] VARIANT LinkTypeIfInsert 
                                             ) #RK: I was also trying to tune optional parameters - can you give me also hint on how to use them to get best results?
            
    direction1 = Visum.Net.Directions.ItemByKey(">") 
    direction2 = Visum.Net.Directions.ItemByKey("<") 
    
    TSys=Visum.Net.TSystems.ItemByKey("B")
    linia="Jakas tam linia"
    Line=Visum.Net.AddLine(linia,"B") #RK: What's the difference between the one on the left and: Line=Visum.Net.AddLine(linia,TSys)
            
    Route=Visum.CreateNetElements()
    SP=Visum.Net.StopPoints.ItemByKey(1) 
    Route.Add(SP)
    SP=Visum.Net.StopPoints.ItemByKey(2) 
    Route.Add(SP)           
                
    Visum.Net.AddLineRoute(linia,Line,direction1,Route,routesearchparameters) #RK: WHat'd be the difference between those two?
    Visum.Net.AddLineRoute(linia,Line,direction1,Route)
    
    #RK: Error's that I have: a) Cannot evaluate route items b) Failed to create line route
    
    #RK: Network details: All elements openet for TSys 'B' : Links, Turns, StopPoints. All links with reasonable TPutTSys. Recalculated network: direct dist ~length.
    #RK: When I try to add the same line via Visum windows it's fine, no errors, no problems (with all tested 'Parameters' in AddLineRoute dialog box.

try:
    Visum 
except:        
    Visum=VisumInit("C:\stops6.ver")     

Add_Line(slownik_linie,slownik_stop_points)
