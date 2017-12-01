Nodes=Visum.Net.Nodes.GetMultiAttValues("No")
Nodes=[node[1] for node in Nodes]
for node in Nodes:
	Visum.Net.Nodes.ItemByKey(node).SetAttValue("No",node+500000)
i=1
for node in Nodes:
	Visum.Net.Nodes.ItemByKey(node+500000).SetAttValue("No",i)
	i+=1