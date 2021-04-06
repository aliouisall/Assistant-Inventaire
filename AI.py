# Importation des différentes bibliothèques
import pandas as pd
import argparse
import xlsxwriter
import re
from natsort import natsorted, ns

# Définition du nombre maximal de lignes à afficher
pd.options.display.max_rows = 30000

# Définition des différentes options et arguments
parser = argparse.ArgumentParser()
parser.add_argument("-f", "--file", dest="inputFile", help="permet de désigner un input de type fichier", metavar="FILE")

args = parser.parse_args()

# Ouverture et lecture du fichier
df = pd.read_excel(args.inputFile)


string = ["T134", "T1152", "T89"]
#print(natsorted(df["itemcallnumber"]))

def naturalSorting(list_):
    # decorate
    tmp = [ (int(re.search('\d+', i).group(0)), i) for i in list_ ]
    tmp.sort()
    # undecorate
    return [ i[1] for i in tmp ]

#print(naturalSorting(string))

# Extraction du bloc des T
barcode = df['barcode']
itemcallnumber = df['itemcallnumber']
#print(df[['barcode', 'itemcallnumber', 'title', 'location']])
thesis = pd.DataFrame(naturalSorting(df["itemcallnumber"]))
#print(thesis)
#thesis.merge(pd.DataFrame(df[['barcode', 'itemcallnumber']]), left_on = 'itemcallnumber')

# Création et initialisation du fichier excel trié
# df_new = pd.DataFrame({'Barcode': [], 'Itemcallnumber': [], 'Title': [], 'Location': []})
df_new = thesis
writer = pd.ExcelWriter('Tri.xlsx', engine ='xlsxwriter')
df_new.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()