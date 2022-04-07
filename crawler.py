import requests
from bs4 import BeautifulSoup
from xlwt import *

url = "https://www.rottentomatoes.com/top/bestofrt/"

# When requesting access to the content of a webpage, sometimes you will find that a 403 error will appear. This is
# because the server has rejected your access. This is the anti-crawler setting used by the webpage to prevent
# malicious collection of information. At this time, you can access it by simulating the browser header information.
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 '
                  'Safari/537.36 QIHU 360SE'
}
f = requests.get(url, headers=headers)

# Create a BeautifulSoup object and specify the parser as lxml.
soup = BeautifulSoup(f.content, 'lxml')

# After extracting the page content using BeautifulSoup,
# we can use the find method to extract the relevant information.
# The <a> tag defines a hyperlink, which is used to link from one page to another.
# We've written table, because that's the class's name
movies = soup.find("table", {"class": "table"}).find_all("a")

movies_list = []
num = 0

# Creating the Excel file and writing the data in it.

# Setting the encoding and adding the sheet.
workbook = Workbook(encoding="utf-8")
table = workbook.add_sheet("data")

# Create the header of each column in the first row.
table.write(0, 0, "Number")
table.write(0, 1, "Movie URL")
table.write(0, 2, "Movie Name")
table.write(0, 3, "Movie Introduction")
line = 1

# Writes the data in the Excel sheet.
for anchor in movies:
    urls = "https://www.rottentomatoes.com" + anchor['href']
    movies_list.append(urls)
    num += 1

    movie_url = urls
    movie_f = requests.get(movie_url, headers=headers)
    movie_soup = BeautifulSoup(movie_f.content, "lxml")
    movie_content = movie_soup.find("div", {"class": "movie_synopsis clamp clamp-6 js-clamp"})

    print(f"{num}: {urls} \nMovie: {anchor.string.strip()}")
    print(f"Movie info: {movie_content.string.strip()}\n")

    # Write the crawled data into Excel separately from the second row.
    table.write(line, 0, num)
    table.write(line, 1, urls)
    table.write(line, 2, anchor.string.strip())
    table.write(line, 3, movie_content.string.strip())
    line += 1

# Finally, save the Excel file.
workbook.save("Movies Top 100.xls")
