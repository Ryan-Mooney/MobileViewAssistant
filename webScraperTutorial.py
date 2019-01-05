import urllib
from urllib.request import urlopen
import requests
from bs4 import BeautifulSoup

# specify the url
quote_page = 'http://ryan-mooney.com/NKRawsonCreations/'
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
post_data= {'query':'ball'}

# query the website and return the html to the variable ‘page’
page = requests.post(quote_page, headers=headers, data=post_data).content
page=page.decode('utf-8')
print(page)

# parse the html using beautiful soup and store in variable `soup`
soup = BeautifulSoup(page, 'html.parser')

# Take out the <div> of name and get its value
products = soup.findAll('p', {'class': 'product-list-text'})

running_names=''
for each in products:
    name = products.text.strip() # strip() is used to remove starting and trailing
    running_names=running_names+name

print(running_names)