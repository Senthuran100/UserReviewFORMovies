import csv
from imdbpie import Imdb
imdb = Imdb()
import openpyxl
import pandas as pd
review=[]
movie=[]
path='H:\IFS\IMDB\\test.xlsx' # Excel sheet containing the name of the movies 
path1='H:\IFS\IMDB\\test1.xlsx' # Excel sheet containing the result from the IMDB which contain the user review for each movie
df = pd.read_excel(path, sheetname='Sheet1')
for row in df['Movies']:
    try:
     movie.append(row)
     Id = imdb.search_for_title(row)[0]['imdb_id']
     review.append(imdb.get_title_user_reviews(Id)['totalReviews'])
    except IndexError:
        review.append("INVALID")

df = pd.DataFrame({'Movies':movie,'Review': review})
writer = pd.ExcelWriter(path1, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()

