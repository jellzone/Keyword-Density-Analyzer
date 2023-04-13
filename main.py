import csv
from collections import Counter
from textblob import TextBlob
import textstat
import openpyxl

# Read text from the first column of an xlsx file
def read_text_from_xlsx(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    text = ""
    for row in sheet.iter_rows(min_col=1, max_col=1, values_only=True):
        if row[0] is not None:
            text += str(row[0]) + " "
    return text.strip()

# Replace the text variable assignment with the function call
text = read_text_from_xlsx('your_file.xlsx')

from itertools import chain
from collections import Counter

# Analyze keyword density for 1-word and 2-5-word phrases
def keyword_density(text, min_words=1, max_words=5, include_words=None, exclude_words=None):
    words = text.lower().split()
    filtered_words = [word for word in words if (include_words is None or word in include_words) and (exclude_words is None or word not in exclude_words)]
    word_count = len(filtered_words)

    # Generate n-grams (phrases of n words) for n between min_words and max_words
    ngrams = {n: [] for n in range(min_words, max_words + 1)}
    for n in range(min_words, max_words + 1):
        ngrams[n].extend(zip(*[filtered_words[i:] for i in range(n)]))

    # Calculate keyword density for n-grams
    keyword_density = {}
    for n in range(min_words, max_words + 1):
        ngram_counts = Counter(ngrams[n])
        ngram_keyword_density = {' '.join(ngram): (count / word_count) * 100 for ngram, count in ngram_counts.items()}
        keyword_density.update({f'{n}-word phrases, {phrase}': density for phrase, density in ngram_keyword_density.items()})
    return keyword_density

# Content analysis
def content_analysis(text):
    word_count = len(text.split())
    char_length = len(text)
    letters = sum(c.isalpha() for c in text)
    sentences = textstat.sentence_count(text)
    syllables = textstat.syllable_count(text)
    avg_words_per_sentence = word_count / sentences
    avg_syllables_per_word = syllables / word_count
    return {
        'Word Count': word_count,
        'Character Length': char_length,
        'Letters': letters,
        'Sentences': sentences,
        'Syllables': syllables,
        'Average Words/Sentence': avg_words_per_sentence,
        'Average Syllables/Word': avg_syllables_per_word,
    }

# Readability analysis
def readability_analysis(text):
    return {
        'Reading Ease': textstat.flesch_reading_ease(text),
        'Grade Level': textstat.flesch_kincaid_grade(text),
        'Gunning Fog': textstat.gunning_fog(text),
        'Coleman Liau Index': textstat.coleman_liau_index(text),
        'Smog Index': textstat.smog_index(text),
        'Automated Reading Index': textstat.automated_readability_index(text),
    }

# Sentiment analysis
def sentiment_analysis(text):
    blob = TextBlob(text)
    sentences = blob.sentences
    sentiment_data = []
    for sentence in sentences:
        sentiment_data.append({
            'Sentence': str(sentence),
            'Sentiment': sentence.sentiment.polarity,
            'Words Quantity': len(sentence.words),
            'Readability': textstat.flesch_reading_ease(str(sentence)),
        })
    return sentiment_data

# Write results to a CSV file
def export_to_csv(data, filename):
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = list(data[0].keys())
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in data:
            writer.writerow(row)

# Perform the analysis
keyword_density_data = keyword_density(text)
content_analysis_data = content_analysis(text)
readability_analysis_data = readability_analysis(text)
sentiment_analysis_data = sentiment_analysis(text)

# Export the results to CSV files
# Combine content analysis and readability analysis data into one dictionary
combined_data = {**content_analysis_data, **readability_analysis_data}

# Export combined data (content analysis and readability analysis) to a CSV file
def export_combined_data_to_csv(data, filename):
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = list(data.keys())
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerow(data)

# Perform the analysis
keyword_density_data = keyword_density(text)
content_analysis_data = content_analysis(text)
readability_analysis_data = readability_analysis(text)
sentiment_analysis_data = sentiment_analysis(text)

# Export the results to CSV files
# Combine content analysis and readability analysis data into one dictionary
combined_data = {**content_analysis_data, **readability_analysis_data}

# Export the combined content and readability analysis data to a CSV file
export_combined_data_to_csv(combined_data, 'content_readability_analysis.csv')

# Export keyword density data to a CSV file
keyword_density_list = [{'Keyword': k, 'Density': v} for k, v in keyword_density_data.items()]
export_to_csv(keyword_density_list, 'keyword_density.csv')

# Export sentiment analysis data to a CSV file
export_to_csv(sentiment_analysis_data, 'sentiment_analysis.csv')
