# This program contains all the Python code used for textual analysis on the IBM Annual Reports for 1990, 2000, 2010, 2019, and 2020
# Data 620
# Ted Fitch
# Last updated 04APR21


# Word Count Starter Program
# Based on Chuck Severance's Python Code - Romeo Section 9.4
# from Python for Informatics 3
# http://do1.dr-chuck.com/pythonlearn/EN_us/pythonlearn.pdf
#
# Original code changes by Carrie Beam and Nivedita Bijlani for UMUC DATA 620
# Further code changes made by Theodore Fitch for UMGC DATA 620 Assignment 12.1
# Last updated Apr 4th, 2021
# Recent changes were made in order to extract text from IBM annual reports, identify each unique word, create word counts for each year's report, determine the length of each word, and output that information to a CSV file for further analysis.

# SECTION 1: CREATE VARIABLES AND LIBRARIES ########################################

# Import libraries to use
import string
import pandas as pd
from nltk.corpus import stopwords
from nltk.stem.wordnet import WordNetLemmatizer

# Create variables to use later in loops
counts = dict()
lmtzr = WordNetLemmatizer()

# Import the corpus of English stop words
stop = stopwords.words('english')

# Uncomment the print statement below to see what words are included by default
# print (stop)

# Append my own stop words to this list based on content of IBM annual reports
stop.append("2021")
stop.append("2020")
stop.append("2019")
stop.append("2018")
stop.append("2017")
stop.append("2016")
stop.append("2015")
stop.append("2014")
stop.append("2013")
stop.append("2012")
stop.append("2011")
stop.append("2010")
stop.append("2009")
stop.append("2008")
stop.append("2007")
stop.append("2006")
stop.append("2005")
stop.append("2004")
stop.append("2003")
stop.append("2002")
stop.append("2001")
stop.append("2000")
stop.append("1999")
stop.append("1998")
stop.append("1997")
stop.append("1996")
stop.append("1995")
stop.append("1994")
stop.append("1993")
stop.append("1992")
stop.append("1991")
stop.append("1990")
stop.append("1989")
stop.append("1988")
stop.append("50")
stop.append("31")
stop.append("20")
stop.append("10")
stop.append("9")
stop.append("8")
stop.append("7")
stop.append("6")
stop.append("5")
stop.append("4")
stop.append("3")
stop.append("2")
stop.append("1")
stop.append("0")
stop.append("—")
stop.append("–")
stop.append("a")
stop.append("b")
stop.append("c")
stop.append("d")
stop.append("e")
stop.append("f")
stop.append("g")
stop.append("h")
stop.append("i")
stop.append("j")
stop.append("k")
stop.append("l")
stop.append("m")
stop.append("n")
stop.append("o")
stop.append("p")
stop.append("q")
stop.append("r")
stop.append("s")
stop.append("t")
stop.append("u")
stop.append("v")
stop.append("w")
stop.append("x")
stop.append("y")
stop.append("z")
stop.append("us")
stop.append("per")
stop.append("also")
stop.append("like")
stop.append("well")
stop.append("use")
stop.append("one")
stop.append("and")
stop.append("the")
stop.append("a")
stop.append("to")
stop.append("of")
stop.append("in")
stop.append("to")
stop.append("for")
stop.append("as")
stop.append("are")
stop.append("on")
stop.append("is")
stop.append("by")
stop.append("that")
stop.append("with")
stop.append("due")
stop.append("may")
stop.append("its")
stop.append("or")
stop.append("'")
stop.append("ibm")
stop.append("one")
stop.append("nonus")
stop.append("used")
stop.append("page")
stop.append("•")
stop.append("three")
stop.append("fourth")

# SECTION 2: 1999 REPORT (This same section is repeated 5x [once for each report] with the only changes being the filenames and variable names.)
########################################

# Open file and create as variable. Open using utf-8 encoding.
dog99 = open ('IBM_Annual_Report_1999.txt', "r", encoding='utf-8')


# Add words to "words" dictionary by searching text doc line by line
for line in dog99:
    line = line.rstrip()
    line = line.translate(line.maketrans('','',string.punctuation))
    line = line.lower()
    words = line.split()
# Remove stop words
    for word in words:
        if word not in stop:
             # Lemmatize - Lemmatizing is similar to stemming but reduces the word to its root as defined in the dictionary
            lmtzr.lemmatize(word)
            # ASCII Fix. Removes any non ASCII characters from word (many random symbols from PDF to TXT conversion)
            word = word.encode("ascii", "ignore")
            word = word.decode()
            # Now add to counts dictionary if it does not exist
            if word not in counts:
                counts[word] = 1
            else:
                counts[word] += 1

# Update dictionary so any numbers are eliminated and eliminate anything with count value 1 in dictionary
counts = {k: v for k, v in counts.items() if not k.isdigit()}

# Make masterlist containing both the counts of each word, and the words [c,w]
word_list99 = [(counts[w], w, "1999", len(w)) for w in counts]

# Sort the masterlist
word_list99.sort()

# Reverse the list so that the largest are first
word_list99.reverse()

# Only keep top 200 words
tword_list99 = word_list99[:200]

# SECTION 3: 2000 REPORT ########################################

# Open file and create as variable. Open using utf-8 encoding.
dog00 = open ('IBM_Annual_Report_2000.txt', "r", encoding='utf-8')

# Add words to "words" dictionary by searching text doc line by line
for line in dog00:
    line = line.rstrip()
    line = line.translate(line.maketrans('','',string.punctuation))
    line = line.lower()
    words = line.split()
# Remove stop words
    for word in words:
        if word not in stop:
            # Lemmatize - Lemmatizing is similar to stemming but reduces the word to its root as defined in the dictionary
            lmtzr.lemmatize(word)
            # ASCII Fix. Removes any non ASCII characters from word
            word = word.encode("ascii", "ignore")
            word = word.decode()
            # Now add to counts dictionary if it does not exist
            if word not in counts:
                counts[word] = 1
            else:
                counts[word] += 1

# Update dictionary so any numbers are eliminated and eliminate anything with count value 1 in dictionary
counts = {k: v for k, v in counts.items() if not k.isdigit()}

# Make masterlist containing both the counts of each word, and the words [c,w]
word_list00 = [(counts[w], w, "2000", len(w)) for w in counts]

# Sort the masterlist
word_list00.sort()

# Reverse the list so that the largest are first
word_list00.reverse()

# Only keep top 200 words
tword_list00 = word_list00[:200]

# SECTION 4: 2010 REPORT ########################################

# Open file and create as variable. Open using utf-8 encoding.
dog10 = open ('IBM_Annual_Report_2010.txt', "r", encoding='utf-8')


# Add words to "words" dictionary by searching text doc line by line
for line in dog10:
    line = line.rstrip()
    line = line.translate(line.maketrans('','',string.punctuation))
    line = line.lower()
    words = line.split()
# Remove stop words
    for word in words:
        if word not in stop:
            # Lemmatize - Lemmatizing is similar to stemming but reduces the word to its root as defined in the dictionary
            lmtzr.lemmatize(word)
            # ASCII Fix. Removes any non ASCII characters from word
            word = word.encode("ascii", "ignore")
            word = word.decode()
            # Now add to counts dictionary if it does not exist
            if word not in counts:
                counts[word] = 1
            else:
                counts[word] += 1

# Update dictionary so any numbers are eliminated and eliminate anything with count value 1 in dictionary
counts = {k: v for k, v in counts.items() if not k.isdigit()}

# Make masterlist containing both the counts of each word, and the words [c,w]
word_list10 = [(counts[w], w, "2010", len(w)) for w in counts]

# Sort the masterlist
word_list10.sort()

# Reverse the list so that the largest are first
word_list10.reverse()

# Only keep top 200 words
tword_list10 = word_list10[:200]

# SECTION 5: 2019 REPORT ########################################

# Open file and create as variable. Open using utf-8 encoding.
dog19 = open ('IBM_Annual_Report_2019.txt', "r", encoding='utf-8')


# Add words to "words" dictionary by searching text doc line by line
for line in dog19:
    line = line.rstrip()
    line = line.translate(line.maketrans('','',string.punctuation))
    line = line.lower()
    words = line.split()
# Remove stop words
    for word in words:
        if word not in stop:
            # Lemmatize - Lemmatizing is similar to stemming but reduces the word to its root as defined in the dictionary
            lmtzr.lemmatize(word)
            # ASCII Fix. Removes any non ASCII characters from word
            word = word.encode("ascii", "ignore")
            word = word.decode()
            # Now add to counts dictionary if it does not exist
            if word not in counts:
                counts[word] = 1
            else:
                counts[word] += 1

# Update dictionary so any numbers are eliminated and eliminate anything with count value 1 in dictionary
counts = {k: v for k, v in counts.items() if not k.isdigit()}

# Make masterlist containing both the counts of each word, and the words [c,w]
word_list19 = [(counts[w], w, "2019", len(w)) for w in counts]

# Sort the masterlist
word_list19.sort()

# Reverse the list so that the largest are first
word_list19.reverse()

# Only keep top 200 words
tword_list19 = word_list19[:200]

# SECTION 6: 2020 REPORT ########################################

# Open file and create as variable. Open using utf-8 encoding.
dog20 = open ('IBM_Annual_Report_2020.txt', "r", encoding='utf-8')

# Add words to "words" dictionary by searching text doc line by line
for line in dog20:
    line = line.rstrip()
    line = line.translate(line.maketrans('','',string.punctuation))
    line = line.lower()
    words = line.split()

# Remove stop words
    for word in words:
        if word not in stop:
            # Lemmatize - Lemmatizing is similar to stemming but reduces the word to its root as defined in the dictionary
            lmtzr.lemmatize(word)
            # ASCII Fix. Removes any non ASCII characters from word
            word = word.encode("ascii", "ignore")
            word = word.decode()
            # Now add to counts dictionary if it does not exist
            if word not in counts:
                counts[word] = 1
            else:
                counts[word] += 1

# Update dictionary so any numbers are eliminated and eliminate anything with count value 1 in dictionary
counts = {k: v for k, v in counts.items() if not k.isdigit()}

# Make masterlist containing both the counts of each word, and the words [c,w]
word_list20 = [(counts[w], w, "2020", len(w)) for w in counts]

# Sort the masterlist
word_list20.sort()

# Reverse the list so that the largest are first
word_list20.reverse()

# Only keep top 100 words
tword_list20 = word_list20[:200]

# SECTION 7: MASTERLIST COMPILATION AND PRINT ########################################

# Create masterlist of truncated (200 word length) lists
word_list = tword_list99 + tword_list00 + tword_list10 + tword_list19 + tword_list20

# Add headers
headers = [('Count', 'Word', 'Year', 'Length')]
word_list = headers + word_list


# Use pandas to add wordslist to CSV - this will be taken over to Tableau for further processing
df = pd.DataFrame(word_list)
df.to_csv('words_ibm_annrep_top200.csv', index=False, header=False)

# Make masterlist of all words
mword_list = word_list99 + word_list00 + word_list10 + word_list19 + word_list20

# Add headers
headers = [('Count', 'Word', 'Year', 'Length')]
mword_list = headers + mword_list

# Use pandas to add master wordslist to CSV - this will be taken over to Tableau for further processing
df = pd.DataFrame(mword_list)
df.to_csv('words_ibm_annrep_master.csv', index=False, header=False)

# Close open txt files so they can be used by other programs
dog99.close()
dog00.close()
dog10.close()
dog19.close()
dog20.close()

#End of script
