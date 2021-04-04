'''
The goal is to get all email contained in a plain text and maintain the email.CSV Database.
Author: Yanko Sarzedas da Costa
Version: 1
'''

from os import listdir, getcwd
from os.path import isfile, join
import pandas as pd

# Get all plain text in a current folder.
current_path = getcwd() + '\\DB'
files = [f for f in listdir(current_path) if isfile(join(current_path, f))]
# Check out if there is a old vervions od CSV emails file.
try:
   position = files.index('emails.csv')
   df = pd.read_csv(current_path + '\\' + 'emails.csv', sep=';')
except:
    position = -1
    df = None

# Filter the plain text.
files = [f for f in files if f.endswith('.txt')]

# Build a valid list ascii code
valid_ascciis = [64] + [i for i in range(65, 91)] + [i for i in range(97, 123)] + [45] + [46] + [95] + [i for i in range(48, 58)]

# Extract from all plane files the emails to export to a CSV file
try:
    # get old version emails list.
    emails = df['email'].tolist()
except:
    emails = []

for file in files:
    print(file)
    f = open(current_path + '\\' + file, 'r', encoding="utf8")
    text = f.read()
    while len(text) > 0:
        # Extract the first parte where probably there is an email.
        start = text.find(' ') + 1
        position = text.find('@', start)
        if position == -1: position = 0
        end1 = text.find('\n', position)
        end2 = text.find(' ', position)
        if end1 != -1:
            end = end1 + 1
        elif end2 != -1:
            end = end2 + 1
        else:
            end = len(text)
        email_text = text[start:end].strip()
        # Extract the second parte that is possible to have an email.
        for email_pos in range(len(email_text)-1, 0, -1):
            if ord(email_text[email_pos]) not in valid_ascciis:
                break
        email_text = email_text[email_pos+1:]
        # Check out if it is an email.
        if not email_text.startswith('@'):
            if email_text.endswith('.'):
                email_text = email_text[:len(email_text)-1]
            if email_text.find('@') != -1 and email_text.find('.') != -1:
                emails.append(email_text.strip())
        text = text[end: ]

# Sort.
emails.sort()
# Remove duplicated emails
emails = list(dict.fromkeys(emails))

# Export an email list into a CSV file.
df = pd.DataFrame({'email':emails})
df.to_csv(current_path + '\\' + 'emails.csv', sep=';', index=False)
print('Tamanho da lista:', len(emails))