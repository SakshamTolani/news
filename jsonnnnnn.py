import requests
import json
from win32com.client import Dispatch

def voice(sstr):
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(sstr)
if __name__=="__main__":
    voice("News for Today. Lets Go!!")
    url = 'https://newsapi.org/v2/top-headlines?country=in&apiKey=b337d49a717745f592543e93af533124'
    news = requests.get(url).text
    news_dict=json.loads(news)

    arts = news_dict['articles']
    for article in arts:
        print(article['title'])
        voice(article['title'])
        voice("Next")
    print("Thank You for listening. Have a Nice Day!!")
    voice("Thankyou for listening. Have a Nice Day!!")
