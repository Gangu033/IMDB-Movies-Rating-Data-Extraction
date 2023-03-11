from bs4 import BeautifulSoup
import openpyxl
import requests
from openpyxl import Workbook
wb=openpyxl.Workbook()
sheet=wb.active
sheet.append(['Rank', 'Movies Name', 'Release Year', 'Rating'])

try:
    source=requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()
    soup=BeautifulSoup(source.text,'html.parser')
    movies=soup.find('tbody',class_='lister-list').find_all('tr')
    #print(movies)
    for movie in movies:
        
        movie_name=movie.find('td',class_='titleColumn').find('a').text
        rating=movie.find('td',class_='ratingColumn imdbRating').strong.text
        
        ra=movie.find('td',class_='titleColumn').get_text(strip=True)
        rank=ra.split('.')[0]
        #movie_nm=movie.find('td',class_='titleColumn').get_text(strip=True).split('.')[1].split('(')[0]
        # yr=movie.find('td',class_='titleColumn').get_text(strip=True).split('.')[1].split('(')[1].strip(')')
        year=movie.find('td',class_='titleColumn').span.text.strip('()')
        #print(rating)
        #print(rank)
        #print(movie_nm)
        sheet.append([rank, movie_name, year, rating])
        #print(rank, movie_name, year, rating)
        
except Exception as e:
    print(e)
wb.save('Movies_Rating.xlsx')
wb.close()
