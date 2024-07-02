import nltk
from nltk.corpus import stopwords
import re
from nltk.tokenize import word_tokenize, sent_tokenize
import pandas as pd
import os
import requests
from bs4 import BeautifulSoup

# Load input data
data = pd.read_excel('Input.xlsx')
URL = data['URL']
URL_ID = data['URL_ID']

def extract_data(URL):
    r = requests.get(URL)
    soup = BeautifulSoup(r.content, 'html5lib')

    title_h1 = soup.find('h1')
    title = title_h1.get_text() if title_h1 else "No Title"

    article_p = soup.find_all('p')
    if article_p:
        article_p = article_p[:-3]
        data = ' '.join([p.get_text() for p in article_p])
    else:
        data = "No Content"

    return title, data

# Load positive and negative words
with open('MasterDictionary/positive-words.txt', 'r') as f:
    positive = set(f.read().split())

with open('MasterDictionary/negative-words.txt', 'r') as f:
    negative = set(f.read().split())

def load_stop_words(files):
    stop_words = set()
    for file in files:
        with open(file, 'r') as f:
            for line in f:
                stop_words.add(line.strip())
    return stop_words

files = [
    'StopWords/StopWords_Auditor.txt',
    'StopWords/StopWords_Currencies.txt',
    'StopWords/StopWords_DatesandNumbers.txt',
    'StopWords/StopWords_Generic.txt',
    'StopWords/StopWords_GenericLong.txt',
    'StopWords/StopWords_Geographic.txt',
    'StopWords/StopWords_Names.txt'
]

stop_words = load_stop_words(files)

def clean_and_tokenize(text):
    text = re.sub(r'[^a-zA-Z\s]', ' ', text)
    tokens = word_tokenize(text.lower())
    tokens = [word for word in tokens if word not in stop_words]
    return tokens

def analyze_sentiment(tokens):
    positive_score = sum(1 for word in tokens if word in positive)
    negative_score = sum(1 for word in tokens if word in negative)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score

def calculation(tokens):
    word_count = len(tokens)
    avg_word_len = sum(len(word) for word in tokens) / word_count
    personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', ' '.join(tokens)))
    return word_count, avg_word_len, personal_pronouns



def calculate_average_sentence_length(text):
    sentences = sent_tokenize(text)
    num_sentences = len(sentences)
    words = word_tokenize(text)
    num_words = len(words)
    average_sentence_length = num_words / num_sentences 
    return average_sentence_length

def avg_num_of_words_per_sen(text):
    sentences = sent_tokenize(text)
    num_sentences = len(sentences)
    words = word_tokenize(text)
    num_words = len(words)
    average_words_per_sentence = num_words / num_sentences 
    return average_words_per_sentence


def syllable(tokens):
    vowels = 'aeiou'
    words_syllable = set()

    for word in tokens:
        word = word.lower().strip()  
        if any(vowel in word for vowel in vowels):
            if not (word.endswith('ed') or word.endswith('es')):
                words_syllable.add(word)
            
    return words_syllable

def complex_words(tokens,syllable_words):
    count = 0
    complex_words = []
    for word in tokens:
        if word in syllable_words:
            count += 1
            if count > 2:
                complex_words.append(word)

    return complex_words

def calculate_syllables_per_word(tokens,syllable_list):
    syllables_per_word = len(syllable_list)/len(tokens)
    return syllables_per_word


results = []


for url_id, url in zip(URL_ID, URL):
    title, data = extract_data(url)

    tokens = [word for word in word_tokenize(data.lower()) if word not in stop_words and re.match(r'^[a-zA-Z]+$', word)]

    syllable_list = syllable(tokens)

    positive_score, negative_score, polarity_score, subjectivity_score = analyze_sentiment(tokens)
    word_count, avg_word_len, personal_pronouns = calculation(tokens)
    avg_sen_len = calculate_average_sentence_length(data)
    complex_words_list = complex_words(tokens,syllable_list)
    percentage_of_complex = len(complex_words_list) / word_count
    fog_index = 0.4 * (avg_sen_len + percentage_of_complex)
    avg_num_per_sen = avg_num_of_words_per_sen(data)
    syllable_per_word = calculate_syllables_per_word(tokens,syllable_list)

    results.append({
        'URL_ID': url_id,
        'URL': url,
        'Title': title,
        'Positive Score': positive_score,
        'Negative Score': negative_score,
        'Polarity Score': polarity_score,
        'Subjectivity Score': subjectivity_score,
        'Word Count': word_count,
        'Avg Word Length': avg_word_len,
        'Personal Pronouns': personal_pronouns,
        'Avg Sentence Length': avg_sen_len,
        'Percentage of Complex Words': percentage_of_complex,
        'FOG Index': fog_index,
        'Avg Number of Words per Sentence': avg_num_per_sen,
        'Complex Word Count': len(complex_words_list),
        'Syllables per Word': syllable_per_word
    })


df = pd.DataFrame(results)


df.to_excel('Output.xlsx', index=False)
