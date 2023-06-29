from bs4 import BeautifulSoup
import requests,openpyxl
import mysql.connector

dB=mysql.connector.connect(
    host = "localhost",
    user = "root",
    passwd = "sabari",
    database="scrapimdb"
)
cursor=dB.cursor()

excel=openpyxl.Workbook()
sheet=excel.active
sheet.title="IMDB datum"
sheet.append(["RANK","MOVIE","YEAR","RATING"])


try:
    respose=requests.get("https://www.imdb.com/chart/top/")
    soup=BeautifulSoup(respose.text,'html.parser')
    movies=soup.find('tbody',class_="lister-list").find_all("tr")
    for movie in movies:
        #print(movie)
        movieName=movie.find('td',class_="titleColumn").a.text
        year=movie.find('td',class_="titleColumn").span.text[1:-1]
        movieRating=movie.find('td',class_="ratingColumn").strong.text
        rank=movie.find('td',class_="titleColumn").text.replace(".","").split()[0]
        sheet.append([rank,movieName,year,movieRating])
        values=(rank,movieName,year,movieRating)
        query="insert into top250 (RANKING,MOVIE,YEARS ,RATING) values(%s,%s,%s,%s)"
        cursor.execute(query,values)
        dB.commit()

except Exception as e:
    print(e)

excel.save("IMDBTop_250.xlsx")