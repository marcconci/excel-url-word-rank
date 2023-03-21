from bs4 import BeautifulSoup
import collections
import multiprocessing
import nltk
import pandas as pd
import requests
import xlrd
from nltk.tokenize import RegexpTokenizer

# Load data from Excel
dados = pd.read_excel('data.xlsx')

# Load dictionary from Excel
dicionario = xlrd.open_workbook('dictionary.xls').sheet_by_name('Sheet1')
dic_list = [dicionario.cell_value(i, 0) for i in range(dicionario.nrows)]

# Define a timeout decorator for the word_freq function
def timeout(seconds):
    def decorator(function):
        def wrapper(*args, **kwargs):
            with multiprocessing.pool.ThreadPool(processes=1) as pool:
                result = pool.apply_async(function, args=args, kwds=kwargs)
                try:
                    return result.get(timeout=seconds)
                except multiprocessing.TimeoutError as e:
                    return e
        return wrapper
    return decorator

# Define the word_freq function to count word frequency and match words to the dictionary
@timeout(1)
def word_freq(url):
    try:
        r = requests.get(url)
        soup = BeautifulSoup(r.text, "html5lib")
        text = soup.get_text()
        tokenizer = RegexpTokenizer('\w+')
        tokens = tokenizer.tokenize(text)
        words = [word.lower() for word in tokens]
        sw = nltk.corpus.stopwords.words('english')
        words_ns = [word for word in words if word not in sw]
        words_all = collections.Counter(words_ns)
        rank = sum([1 for key in words_all.items() if key[0] in dic_list])
        return rank
    except:
        pass

# Process each URL in the data and add its rank to the DataFrame
for i, url in enumerate(dados.iloc[:, 0]):
    print(round(100 * i / len(dados)), '%')
    dados.iloc[i, 1] = word_freq(str(url))

# Save the processed data to a new Excel file
dados.to_excel('processed_data.xlsx')
print('Data exported successfully!')