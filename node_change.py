def get_Nodes(Attr):
    return [node[1] for node in Visum.Net.Nodes.GetMultiAttValues(Attr)]


def Set_Nodes(old_Nodes):
    
    z=float(len(old_Nodes))
    for i in range(len(old_Nodes)):
        Visum.Net.Nodes.ItemByKey(old_Nodes[i]).SetAttValue("No",30000+old_Nodes[i])

Nodes=get_Nodes("No")
Set_Nodes(get_Nodes("No"))