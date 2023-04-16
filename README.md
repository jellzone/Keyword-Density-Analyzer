# Reviews-Keyword-Density-Analyzer
This Text Analytics Tool is a Python script that analyzes text data from an Excel file, performing keyword density analysis, content analysis, and sentence analysis. The script uses the Natural Language Toolkit (nltk) and Textstat libraries to analyze the text and generates an output Excel file with separate sheets for keyword density, content analysis, and sentence analysis.

# Usage
1. Install Python 3.x on your system, if not already installed.
2. Install the required dependencies by running:
``
pip install -r requirements.txt``
or
``
pip install openpyxl nltk textstat``
This will install the openpyxl, nltk, and textstat libraries, which are required for the script to work.
3. Edit the text variable in main.py to your desired input text.
4. Run the script with the command:

``
python text_analytics_tool.py
``
The script will output the results to the console, and also export a CSV file with the keyword density analysis.

Learn this url that if you have problem in Mac Os: https://stackoverflow.com/questions/38916452/nltk-download-ssl-certificate-verify-failed

#  Customization
You can customize the behavior of the script by modifying the following variables in main.py:

- text: the input text to be analyzed. the defult is your_file.xlsx, you can change the file name as you wish.
- include_words: a list of words that should be included in the analysis, regardless of frequency.
- exclude_words: a list of words that should be excluded from the analysis, regardless of frequency.
- ngram_sizes: a list of integers representing the number of words in each n-gram phrase to be analyzed.
csv_file_path: the path to the output CSV file.

#  Contributing
Feel free to contribute to this project by submitting bug reports, feature requests, or pull requests.

#  Question
If you have any questions or issues, please submit them in the Issues section of the GitHub repository. We'll do our best to address them promptly.
