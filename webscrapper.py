from bs4 import BeautifulSoup
import requests, openpyxl


excel = openpyxl.Workbook()
#print(excel.sheetnames)
sheet = excel.active
sheet.title = "IMDB Scrapped data"
#print(excel.sheetnames)
sheet.append(['Movie Rank','Movie Name','Release Year', 'Rating'])

try:
    source = requests.get('https://www.imdb.com/chart/top')
    source.raise_for_status()

    soup = BeautifulSoup(source.text, 'html.parser')
    
    movies = soup.find('tbody', class_="lister-list").find_all('tr')
    #print(len(movies))==>250

    for movie in movies:


        movie_name = movie.find('td', class_="titleColumn").a.text

        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]

        release_year = movie.find('td', class_="titleColumn").span.text.strip('()')

        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
        

        print(rank, movie_name, release_year, rating)

        sheet.append([rank, movie_name, release_year, rating])

except Exception as e:
    print(e)


excel.save('IMDB Web scrapped data.xlsx')
