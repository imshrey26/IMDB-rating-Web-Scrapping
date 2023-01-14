from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Rank', 'Movie Name','Year of Release','IMDB Rating'])


try:
    source = requests.get('https://www.imdb.com/chart/top/')  # response object fetches the source code of this HTML web page
    source.raise_for_status() # Captures the error if website is not reachable
    
    soup = BeautifulSoup(source.text,'html.parser') # Collects all the html data from source website in f orm of text and parses

    movies = soup.find('tbody',class_="lister-list") # hits the first class containing all the content 
    movies_list = movies.find_all('tr') # finds list of all tr's under the class lister-list of tbody
    print(f'Total number of movies are {len(movies_list)}')

    for movie in movies_list:

        name = movie.find('td',class_='titleColumn').a.text
       
       #rank = movie.find('td',class_='titleColumn')text Gets all the text 1. The Shawasnk and 1994 but with lot of whitespaces
       
        rank = movie.find('td',class_='titleColumn').get_text(strip=True).split('.')[0]
        
        year = movie.find('td',class_='titleColumn').span.text.strip('()')
        
        rating = movie.find('td',class_='ratingColumn imdbRating').strong.text
        
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])


except Exception as e: # to avoid crashing of the program once there is an error
    print(e)

excel.save('IMDB Movie Ratings.xlsx')


