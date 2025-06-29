import pandas as pd
import matplotlib as plt
df = pd.read_csv("updated_export.csv")
dfe = df.columns
list = ["Make", "Model", "City", "Status", "Finance Rate"]
dfd = df[list]
list_of_makes =  dfd["Model"].unique().tolist()
list_of_makes =  dfd["Model"].unique().tolist()
list_of_makes
#number = dfd["Make"].value_counts()["BMW"]
#number
list_of_makes = ["LEAF", "Rogue", "ARIYA", "Pathfinder", "Kicks", "Sentra", "Titan", "Frontier", "Murano", "Altima"]
data = []
for a in list_of_makes:
    number = dfd["Model"].value_counts()[a].item()
    d = {a: number}
    data.append(d)
data

dfee = pd.DataFrame()
list_values = []
for i in list_of_makes:
    number = dfd["Model"].value_counts()[i].item()
    list_values.append(number)
dfgg = pd.DataFrame(data)


dfgg = pd.DataFrame({'count' : list_values},
                    index=list_of_makes)

dfgg.plot.pie(y="count", figsize=(16,16), labeldistance=1.1, textprops={'fontsize': 30})