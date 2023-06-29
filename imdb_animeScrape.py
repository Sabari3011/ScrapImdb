from bs4 import BeautifulSoup
import openpyxl,requests
import mysql.connector

db=mysql.connector.connect(
    host="localhost",
    user="root",
    passwd="sabari",
    database="scrapimdb"
)
cursor=db.cursor()

excel=openpyxl.Workbook()
sheet=excel.active
sheet.title="Anime"
sheet.append(["SNO","NAME","YEAR","RATING","STORY","DIRECTOR","BUDGET"])

try:
    result=requests.get("https://www.imdb.com/search/title/?genres=animation&sort=user_rating,desc&title_type=feature&num_votes=25000,&pf_rd_m=A2FGELUUNOQJNL&pf_rd_p=f11158cc-b50b-4c4d-b0a2-40b32863395b&pf_rd_r=RA600WAVYT16HZG6RFF3&pf_rd_s=right-6&pf_rd_t=15506&pf_rd_i=top&ref_=chttp_gnr_3")
    soup = BeautifulSoup(result.text, 'html.parser')
    movies=soup.find("div",class_="lister-list").find_all("div",class_="lister-item")
    for movie in movies:
        sno=movie.find("span",class_="lister-item-index").text.replace(".","")
        movieName=movie.find("h3",class_="lister-item-header").a.get_text(strip=True)
        year=movie.find("span",class_="lister-item-year").text[1:-1]
        rating=movie.find("div",class_="ratings-bar").strong.text
        story=movie.find("p").findNext("p").get_text(strip=True)
        director=movie.find("p").findNext("p").findNext("p").a.text
        gross=movie.find("p",class_="sort-num_votes-visible").find_all("span")[-1].get_text(strip=True)
        query="insert into movielist (NAME,YEAR,RATING,STORY,DIRECTOR,BUDGET) values (%s,%s,%s,%s,%s,%s)"
        values=(movieName,year,rating,story,director,gross)
        cursor.execute(query,values)
        db.commit()
        sheet.append([sno,movieName,year,rating,story,director,gross])


except Exception as e:
    print(e)
excel.save("anime.xlsx")