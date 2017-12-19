list1 = Visum.Net.Links.GetMultipleAttributes(["No","FromNodeNo"])
list2 = Visum.Net.Links.GetMultipleAttributes(["No","ToNodeNo"])

l = list()

for i, line in enumerate(list1):
    if line[0] < 1000:
        l.append(line)
    else:
        l.append(list2[i])

l = tuple(l)





