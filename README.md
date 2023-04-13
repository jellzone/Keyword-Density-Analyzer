# Keyword-Density-Analyzer
This script analyzes the keyword density of a given text, by calculating the frequency of occurrence of one-word and multi-word phrases. It also computes various text statistics, such as word count, character length, sentence count, and readability scores.
# Usage
1. Install Python 3.x on your system, if not already installed.
2. Install the required dependencies by running:
``
pip install -r requirements.txt``
3. Edit the text variable in main.py to your desired input text.
4. Run the script with the command:

``
python main.py``
The script will output the results to the console, and also export a CSV file with the keyword density analysis.
#  Customization
You can customize the behavior of the script by modifying the following variables in main.py:

- text: the input text to be analyzed.
- include_words: a list of words that should be included in the analysis, regardless of frequency.
- exclude_words: a list of words that should be excluded from the analysis, regardless of frequency.
- ngram_sizes: a list of integers representing the number of words in each n-gram phrase to be analyzed.
csv_file_path: the path to the output CSV file.

#  Contributing
Feel free to contribute to this project by submitting bug reports, feature requests, or pull requests.
