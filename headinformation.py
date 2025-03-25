import pandas as pd
import re

NOTES_SEASON = "النقاط  الخاصة بـ:"
SCHOOL_YEAR = "السنة الدراسية :"
CLASS = "الفوج التربوي :"
SUBJECT = "مادة :"
KEYWORDS = [NOTES_SEASON, SCHOOL_YEAR, CLASS, SUBJECT]
remplacehead = ["[notes_of_season]", "[school_year]", "[class]", "[subject]"]


class HeadInformation:
    def __init__(self, path, sheet):
        self.file = pd.read_excel(path, sheet_name=sheet)
        self.sentence = self.file.loc[3].to_list()[0]
        self.sentence_di = {}
        self.getheadinfo()
        self.nameclass = self.sentence_di["[class]"]

# Read a Head of File
# The sentence: information about the ClassSchool
#     def sentence(self):
#         file_line = self.file.loc[3].to_list()[0]

# Function to extract words after the specified keyword
    def extractafterkeyword(self, keyword, text, next_keyword=None):
        """
        1 keyword: This is the specific word or phrase that the function looks for in the given text.
        2 text: This is the full text (or sentence) in which the function searches for the keyword
        3 next_keyword: This argument represents the next keyword in the sequence.
         It is used to determine where to stop the extraction.
        4 s*(.*?)s* : This is useful for extracting specific text while ignoring extra spaces around it.
        """
        if next_keyword:
            # Regular expression to find the word(s) after the keyword and stop at the next keyword
            pattern = rf"{re.escape(keyword)}\s*(.*?)\s*{re.escape(next_keyword)}"
        else:
            # Regular expression to find the word(s) after the keyword until the end of the sentence
            pattern = rf"{re.escape(keyword)}\s*(.*)"

        match = re.search(pattern, text)
        if match:
            return match.group(1).strip()  # # Extract the result and clean it from extra spaces
        else:
            return None

    def getheadinfo(self):
        # Keywords to search for in order
        # Extract words after each keyword and stop at the next keyword
        for i, keyword in enumerate(KEYWORDS):
            next_keyword = KEYWORDS[i + 1] if i + 1 < len(KEYWORDS) else None
            result = self.extractafterkeyword(keyword, self.sentence, next_keyword)
            self.sentence_di[remplacehead[i]] = result

