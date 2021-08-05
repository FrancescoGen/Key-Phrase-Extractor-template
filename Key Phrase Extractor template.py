#Librerie da importare
import pandas as pd
import numpy as np
import pke    
import spacy 

Percorsi_df = pd.read_excel(r'NAME.xlsx')
print (Percorsi_df.head())

List_Percorsi = Percorsi_df.values.tolist() #Turn df into list

###
new_column = []

for i in range(0, len(List_Percorsi)):
    if type(List_Percorsi[i][9]) != str:
        new_column.append(["nan"])
    else:
        extractor = pke.unsupervised.TopicRank()
        extractor.load_document(input=List_Percorsi[i][9], language='it')
        extractor.candidate_selection(pos={"NOUN", "PROPN" "ADJ"}) # keyphrase candidate selection, here sequences of nouns and adjectives # defined by the Universal PoS tagset
        extractor.candidate_weighting() # candidate weighting, here using a random walk algorithm 
        new_column.append(extractor.get_n_best(n=4)) # N-best selection, keyphrases contains the 4 highest scored candidates as # (keyphrase, score) tuples
 

new_column_3 = []
    
for i in new_column:
    if 'nan' in i:
        new_column_3.append(['nan'])
    else:
        new_column_3.append([tup[0] for tup in i])

new_column_4 = []

for i in new_column_3:
    if 'nan' in i:
        new_column_4.append('nan')
    else:
        new_column_4.append(', '.join(i))
            
            
Percorsi_df['key_phrases'] = new_column_4
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Percorsi_1.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
Percorsi_df.to_excel(writer)

writer.save()
