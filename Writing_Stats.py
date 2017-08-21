            ##### Writing Stats #####


            ##Helper Functions and Libraries##
#nltk.download()  #only once

from tkinter import *
from tkinter import ttk
import nltk
import collections
import re


          

try:          #from http://etienned.github.io/posts/extract-text-from-word-docx-simply/
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile

 
"""
Module that extract text from MS XML Word document (.docx).
(Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'


def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)
    
    




            ##Writing Stats Main Function##

def writing_stats(text):

  pb = ttk.Progressbar(orient='horizontal', mode='determinate', length=100)
  pb.grid(row=11, pady=(20, 20))

  pb['value'] = 10
  root.update_idletasks()

  counter = 0
  counter1 = 0
  counter2 = 0
  word_count = 0
  letter_count = 0
  letters_per_word = []
  letters_current_word = 0
  words_per_sentence = []
  words_current_sentence = 0
  sentence_count = 0
  sentences_current_paragraph = 0
  paragraph_count = 0
  sentences_per_paragraph = []
  pronouns = []
  nouns = []
  verbs = []
  adverbs = []
  conjunctions = []
  prepositions = []
  interjections = []
  adjectives = []
  articles = []
  first_word = {'pronoun' : 0, 'noun' : 0, 'verb' : 0, 'adverb' : 0, 'conjunction' : 0, 'preposition' : 0, 'interjection' : 0, 'adjective' : 0, 'article' : 0}
  end_marks = {'period' : 0, 'exclamation point' : 0, 'question mark' : 0}
  other_punctuation = {'comma' : 0, 'semicolon' : 0, 'colon' : 0, 'ampersand' : 0, 'hyphen' : 0, \
                       'exclamation point' : 0, 'period' : 0, 'dash' : 0, 'question mark' : 0, 'parenthesis' : 0, 'quotation mark' : 0, 'ellipses' : 0, 'forward slash' : 0, 'period' : 0}
  voice_per_sentence = {'passive' : 0, 'active' : 0, 'unsure': 0}
  passive_current_sentence = 0

  text = text.replace('\n\n', ' @@')
  text = text.replace('\n\t', ' @@')
  text = text.replace('\t', ' @@')
  text = text.replace('\n', ' @@')
  text = text.replace('\u201c', '"')
  text = text.replace('\u201d', '"')
  text = text.replace('*', '')
  text = text.split(' ')
  text = list(filter(None, text))

  print(text)
  
  pb['value']=20
  root.update_idletasks()
        
  for nxt, word in zip(text[1:]+['i'], text):  
    if word != '-' and word!= '&':
      word_count += 1
      words_current_sentence += 1
    
    if word == '-':
      other_punctuation['dash'] += 1
      letter_count -= 1

    if word == '&':
        other_punctuation['ampersand'] += 1
        letter_count -= 1

    if counter == 0:
      pb['value'] = 30
      root.update_idletasks()
      counter += 1
    
    lower = word.lower()
    lower_nxt = nxt.lower()
    while True:
        try:
            tagged_lower_next = nltk.pos_tag([lower_nxt])
            break
        except IndexError:
            tagged_lower_next = 'I'
    
    if (lower == 'was' and (tagged_lower_next[0][1] == 'VBD' or tagged_lower_next[0][1] == 'VBG' or tagged_lower_next[0][1] == 'VBN')) or (lower == 'were' and (tagged_lower_next[0][1] == 'VBD'\
    or tagged_lower_next[0][1] == 'VBG' or tagged_lower_next[0][1] == 'VBN')) or (lower == 'be' and (tagged_lower_next[0][1] == 'VBD' or tagged_lower_next[0][1] == 'VBG' or tagged_lower_next[0][1] == 'VBN'))\
    or (lower == 'been' and (tagged_lower_next[0][1] == 'VBD' or tagged_lower_next[0][1] == 'VBG' or tagged_lower_next[0][1] == 'VBN')) or (lower == 'has' and (tagged_lower_next[0][1] == 'VBD'\
    or tagged_lower_next[0][1] == 'VBG' or tagged_lower_next[0][1] == 'VBN')) or (lower == 'have' and (tagged_lower_next[0][1] == 'VBD' or tagged_lower_next[0][1] == 'VBG' or tagged_lower_next[0][1] == 'VBN'))\
    or (lower == 'is' and (tagged_lower_next[0][1] == 'VBD' or tagged_lower_next[0][1] == 'VBG' or tagged_lower_next[0][1] == 'VBN')):
        passive_current_sentence += 3


    if ((word[len(word) - 1] == '.' or word[len(word) - 2] == '.') and (nxt[0].isupper() or nxt[0] == '"' or \
    nxt[0] == "@" or word == text[len(text) - 1])):
      sentence_count += 1
      sentences_current_paragraph += 1
      words_per_sentence.append(words_current_sentence)
      words_current_sentence = 0
      end_marks['period'] += 1
      if passive_current_sentence > 2:
        voice_per_sentence['passive'] += 1
      if passive_current_sentence < 2:
        voice_per_sentence['active'] += 1
      if passive_current_sentence == 2:
        voice_per_sentence['unsure'] += 1
      passive_current_sentence = 0

    if '.@@' in word:
      sentence_count += 1
      sentences_current_paragraph += 1
      words_per_sentence.append(words_current_sentence)
      words_current_sentence = 1
      end_marks['period'] += 1
      if passive_current_sentence > 2:
        voice_per_sentence['passive'] += 1
      if passive_current_sentence < 2:
        voice_per_sentence['active'] += 1
      if passive_current_sentence == 2:
        voice_per_sentence['unsure'] += 1
      passive_current_sentence = 0

    
    if '?@@' in word:
      sentence_count += 1
      sentences_current_paragraph += 1
      words_per_sentence.append(words_current_sentence)
      words_current_sentence = 1
      end_marks['question mark'] += 1
      if passive_current_sentence > 2:
        voice_per_sentence['passive'] += 1
      if passive_current_sentence < 2:
        voice_per_sentence['active'] += 1
      if passive_current_sentence == 2:
        voice_per_sentence['unsure'] += 1
      passive_current_sentence = 0

    if '!@@' in word:
      sentence_count += 1
      sentences_current_paragraph += 1
      words_per_sentence.append(words_current_sentence)
      words_current_sentence = 1
      end_marks['exclamation point'] += 1
      if passive_current_sentence > 2:
        voice_per_sentence['passive'] += 1
      if passive_current_sentence < 2:
        voice_per_sentence['active'] += 1
      if passive_current_sentence == 2:
        voice_per_sentence['unsure'] += 1
      passive_current_sentence = 0


    if ((word[len(word) - 1] == '?' or word[len(word) - 2] == '?') and (nxt[0].isupper() or nxt[0] == '"' or \
    nxt[0:1] == "@@" or word == text[len(text) - 1])):
      sentence_count += 1
      sentences_current_paragraph += 1
      words_per_sentence.append(words_current_sentence)
      words_current_sentence = 0
      end_marks['question mark'] += 1
      if passive_current_sentence > 2:
        voice_per_sentence['passive'] += 1
      if passive_current_sentence < 2:
        voice_per_sentence['active'] += 1
      if passive_current_sentence == 2:
        voice_per_sentence['unsure'] += 1
      passive_current_sentence = 0
      
    if ((word[len(word) - 1] == '!' or word[len(word) - 2] == '!') and (nxt[0].isupper() or nxt[0] == '"'  \
    or nxt[0:1] == "@@" or word == text[len(text) - 1])):
      sentence_count += 1
      sentences_current_paragraph += 1
      words_per_sentence.append(words_current_sentence)
      words_current_sentence = 0
      end_marks['exclamation point'] += 1
      if passive_current_sentence > 2:
        voice_per_sentence['passive'] += 1
      if passive_current_sentence < 2:
        voice_per_sentence['active'] += 1
      if passive_current_sentence == 2:
        voice_per_sentence['unsure'] += 1
      passive_current_sentence = 0
      
      
    if len(word) > 3 and '...' in word: 
      other_punctuation['ellipses'] += 1

    if "@@" in word and word != '@@' and word != text[0]:
        paragraph_count += 1
        sentences_per_paragraph.append(sentences_current_paragraph)
        sentences_current_paragraph = 0
        letter_count -= 2
        letters_current_word -= 2

    if word == '@@':
        paragraph_count += 1
        sentences_per_paragraph.append(sentences_current_paragraph)
        sentences_current_paragraph = 0
        letter_count -= 2
        letters_current_word -= 2
        text.remove(word)
        sentence_count += 1
        sentences_current_paragraph += 1
        words_per_sentence.append(words_current_sentence)
        words_current_sentence = 1
        end_marks['period'] += 1
        if passive_current_sentence > 2:
          voice_per_sentence['passive'] += 1
        if passive_current_sentence < 2:
          voice_per_sentence['active'] += 1
        if passive_current_sentence == 2:
          voice_per_sentence['unsure'] += 1
        passive_current_sentence = 0

    if counter1 == 0:
      pb['value'] = 40
      root.update_idletasks()
      counter1 += 1

    if nxt == 'i':
      paragraph_count += 1
      sentences_per_paragraph.append(sentences_current_paragraph)
      sentences_current_paragraph = 0

    
    for letter in word:
      letter_count += 1
      letters_current_word += 1
      if letter == ',':
        other_punctuation['comma'] += 1
        letter_count -= 1
        letters_current_word -= 1
      if letter == ';':
        other_punctuation['semicolon'] += 1
        letter_count -= 1
        letters_current_word -= 1
      if letter == ':':
        other_punctuation['colon'] += 1
        letter_count -= 1
        letters_current_word -= 1
      if letter == '-':
        other_punctuation['hyphen'] += 1
        letter_count -= 1
        letters_current_word -= 1
      if letter == '(' or letter == ')':
        other_punctuation['parenthesis'] += 1
        letter_count -= 1
        letters_current_word -= 1
      if letter ==  '"':
        other_punctuation['quotation mark'] += 1
        letter_count -= 1
        letters_current_word -= 1
      if letter == '/':
        other_punctuation['forward slash'] += 1
        letter_count -= 1
        letters_current_word -= 1
      if letter == '.' or letter == '?' or letter == '!':
        letter_count -= 1
        letters_current_word -= 1
      if letter == '!':
        other_punctuation['exclamation point'] += 1
        letter_count -= 1
        letters_current_word -= 1
      if letter == '?':
          other_punctuation['question mark'] += 1
          letter_count -= 1
          letters_current_word -= 1
      if letter == '.':
          other_punctuation['period'] += 1
          letter_count -= 1
          letters_current_word -= 1

    if counter2 == 0:
      pb['value'] = 50
      root.update_idletasks()
      counter2 += 1
    
    lst_letters = list(enumerate(word))
    for letter in lst_letters:
      if letter == lst_letters[len(lst_letters)-1]:
        letters_per_word.append(letters_current_word)
        letters_current_word = 0

  pb['value'] = 70
  root.update_idletasks()

  
  other_punctuation['period'] = other_punctuation['period'] - end_marks['period'] - (3 * other_punctuation['ellipses'])
  other_punctuation['question mark'] = other_punctuation['question mark'] - end_marks['question mark']
  other_punctuation['exclamation point'] = other_punctuation['exclamation point'] - end_marks['exclamation point']  

  if len(words_per_sentence) > 0:
      ave_words_per_sentence = sum(words_per_sentence)/float(len(words_per_sentence))
      ave_letters_per_word = sum(letters_per_word)/float(len(letters_per_word))
  if len(sentences_per_paragraph) > 0:
      ave_sentences_per_paragraph = sum(sentences_per_paragraph)/float(len(sentences_per_paragraph))

  if sentence_count > 0:
      percent_passive = voice_per_sentence['passive']/float(sentence_count)
      percent_active = voice_per_sentence['active']/float(sentence_count)
      percent_unsure = voice_per_sentence['unsure']/float(sentence_count)
      percentages_voice = {'Percent passive' : percent_passive*100, \
                           'Percent active' : percent_active*100, \
                           'Percent unsure' : percent_unsure*100}

  new_text = []
  notword = 'i'
  for word in text:
    word = word.replace('@@', '')
    if word == '':
        word == notword
    new_text.append(word)

  while True:
      try:
          tagged_text = nltk.pos_tag(new_text)
          break
      except IndexError:
          pass
  print(tagged_text)

  pb['value'] = 75
  root.update_idletasks()


  for nxt, word in zip(tagged_text[1:]+['i'], tagged_text):
      if ((word[0][len(word[0]) - 1] == '.' or word[0][len(word[0])-1]=='?' or  \
      word[0][len(word[0]) - 1] == '!') and (nxt[0][0].isupper() or nxt[0][0] == '"' or  \
      nxt[0][0:2] == "@@")):
          if nxt[1] == "CC":
              first_word['conjunction'] += 1
          if nxt[1] == 'WP' or nxt[1] == 'WPS' or nxt[1] == 'PRP' or nxt[1] == 'PRP$':
              first_word['pronoun'] += 1
          if nxt[1] == 'NN' or nxt[1] == 'NNS' or nxt[1] == 'NNP' or nxt[1] == 'NNPS':
              first_word['noun'] += 1
          if nxt[1] == 'VB' or nxt[1] == 'VBD' or nxt[1] == 'VBG' or nxt[1] == 'VBN' or \
          nxt[1] == 'VBP' or nxt[1] == 'VBZ':
              first_word['verb'] += 1
          if nxt[1] == 'RB' or nxt[1] == 'RBR' or nxt[1] == 'RBS' or nxt[1] == 'WRB':
              first_word['adverb'] += 1 
          if nxt[1] == 'IN':
              first_word['preposition'] += 1
          if nxt[1] == 'UH':
              first_word['interjection'] += 1
          if nxt[1] == 'JJ' or nxt[1] == 'JJR' or nxt[1] == 'JJS':
              first_word['adjective'] += 1
          if nxt[1] == 'DT' or nxt[1] == 'WDT' or nxt[1] == 'PDT':
              first_word['article'] += 1
      if word == tagged_text[0]:
          if word[1] == "CC":
              first_word['conjunction'] += 1
          if word[1] == 'WP' or word[1] == 'WPS' or word[1] == 'PRP' or word[1] == 'PRP$':
              first_word['pronoun'] += 1
          if word[1] == 'NN' or word[1] == 'NNS' or word[1] == 'NNP' or word[1] == 'NNPS':
              first_word['noun'] += 1
          if word[1] == 'VB' or word[1] == 'VBD' or word[1] == 'VBG' or word[1] == 'VBN' or \
          word[1] == 'VBP' or word[1] == 'VBZ':
              first_word['verb'] += 1
          if word[1] == 'RB' or word[1] == 'RBR' or word[1] == 'RBS' or word[1] == 'WRB':
              first_word['adverb'] += 1 
          if word[1] == 'IN':
              first_word['preposition'] += 1
          if word[1] == 'UH':
              first_word['interjection'] += 1
          if word[1] == 'JJ' or word[1] == 'JJR' or word[1] == 'JJS':
              first_word['adjective'] += 1
          if word[1] == 'DT' or word[1] == 'WDT' or word[1] == 'PDT':
              first_word['article'] += 1
          
      alpha_word = re.sub("[^a-zA-Z]+", "", word[0])
      alpha_word = alpha_word.lower()
      if word[1] == 'CC':
          conjunctions.append(alpha_word)
      if word[1] == 'WP' or word[1] == 'WPS' or word[1] == 'PRP' or word[1] == 'PRP$':
          pronouns.append(alpha_word)
      if word[1] == 'NN' or word[1] == 'NNS' or word[1] == 'NNP' or word[1] == 'NNPS':
          nouns.append(alpha_word)
      if word[1] == 'VB' or word[1] == 'VBD' or word[1] == 'VBG' or word[1] == 'VBN' or \
      word[1] == 'VBP' or word[1] == 'VBZ':
          verbs.append(alpha_word)
      if word[1] == 'RB' or word[1] == 'RBR' or word[1] == 'RBS' or word[1] == 'WRB':
          adverbs.append(alpha_word)
      if word[1] == 'IN':
          prepositions.append(alpha_word)
      if word[1] == 'UH':
          interjections.append(alpha_word)
      if word[1] == 'JJ' or word[1] == 'JJR' or word[1] == 'JJS':
          adjectives.append(alpha_word)
      if word[1] == 'DT' or word[1] == 'WDT' or word[1] == 'PDT':
          articles.append(alpha_word)
  
  
  count_nouns = collections.Counter(nouns)
  count_nouns.update({' ' : 0, '  ' : 0, '   ' : 0, '    ' : 0, '     ' : 0})
  top_nouns = []
  for i in range(0,5):
      top_nouns.append((max(count_nouns, key=count_nouns.get), max(count_nouns.values())))
      del count_nouns[max(count_nouns, key=count_nouns.get)]


  count_verbs = collections.Counter(verbs)
  count_verbs.update({' ' : 0, '  ' : 0, '   ' : 0, '    ' : 0, '     ' : 0})
  top_verbs = []
  for i in range(0,5):
      top_verbs.append((max(count_verbs, key=count_verbs.get), max(count_verbs.values())))
      del count_verbs[max(count_verbs, key=count_verbs.get)]
  
  count_adverbs = collections.Counter(adverbs)
  count_adverbs.update({' ' : 0, '  ' : 0, '   ' : 0, '    ' : 0, '     ' : 0})
  top_adverbs = []
  for i in range(0,5):
      top_adverbs.append((max(count_adverbs, key=count_adverbs.get), max(count_adverbs.values())))
      del count_adverbs[max(count_adverbs, key=count_adverbs.get)]
 
  
  count_adjectives = collections.Counter(adjectives)
  count_adjectives.update({' ' : 0, '  ' : 0, '   ' : 0, '    ' : 0, '     ' : 0})
  top_adjectives = []
  for i in range(0,5):
      top_adjectives.append((max(count_adjectives, key=count_adjectives.get), max(count_adjectives.values())))
      del count_adjectives[max(count_adjectives, key=count_adjectives.get)]

  pb['value'] = 100
  root.update_idletasks()

  
   
  return  ('Paragraph count: ' + str(paragraph_count),
  'Word count: ' + str(word_count),
  'Letter count: ' + str(letter_count),
  'Sentence count: ' + str(sentence_count),
  'Average letters per word: ' + str(ave_letters_per_word),
  'Average words per sentence: ' + str(ave_words_per_sentence),
  'Average sentences per paragraph: ' + str(ave_sentences_per_paragraph),
  'End mark frequencies: ' + str(end_marks),
  'Punctuation Frequencies (excluding end marks): ' + str(other_punctuation),
  'Percent of sentences identified to be active, passive, and undetermined: ' + str(percentages_voice),
  'Most frequent nouns: ' + str(top_nouns),
  'Most frequent verbs: ' + str(top_verbs),
  'Most frequent adjectives: ' + str(top_adjectives),
  'Most frequent adverbs: ' + str(top_adverbs),
  'Part of speech of first word in each sentence frequencies: ' + str(first_word))
               
     ### Tkinter GUI ###
  
root = Tk()
root.title("Writing Statistics")

def retrieve_input():
    InputValue=str(text_input.get("1.0","end-1c"))
    window = Toplevel(root)
    window.title('Writing Statistics')
    retrn = writing_stats(InputValue)
    for i in range(0, 14):
        ttk.Label(window, wraplength=450, anchor=CENTER, text=str(retrn[i]), width = 100).pack()
    
def retrieve_file_path():
    file_path=str(filepath.get())
    window = Toplevel(root)
    window.title('Writing Statistics')
    retrn = writing_stats(get_docx_text(file_path))
    for i in range(0, 14):
        ttk.Label(window, wraplength=450, anchor=CENTER, text=str(retrn[i]), width = 100).pack()
    

mainframe = ttk.Frame(root, padding='20 20 20 30')
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

filepath = ttk.Entry(mainframe, width=100)
filepath.grid(column=2, row=7, sticky=W)
ttk.Label(mainframe, text='Enter File Path to Word document:  ').grid(column=1, row=7, sticky=E, padx=(100,0))

scrollbar = Scrollbar(mainframe)
scrollbar.grid(column=3, row=9, sticky=E)
text_input = Text(mainframe, height=6, width=75, wrap=WORD, yscrollcommand=scrollbar.set)
text_input.grid(column=2, row=9, sticky=W)
scrollbar.config(command=text_input.yview)
ttk.Label(mainframe, text='or').grid(column=2, row=8, pady=(10,10))
ttk.Label(mainframe, text='Copy and Paste Text:  ').grid(column=1, row=9, sticky=E)

button = ttk.Button(mainframe, text='Get Writing Statistics from Text Box', command=lambda: retrieve_input())
button.grid(row=9, column=4, sticky=(E), padx=(10,0))
button = ttk.Button(mainframe, text='Get Writing Statistics from File Path', command=lambda: retrieve_file_path())
button.grid(row=7, column=4, sticky=(E), padx=(10,0))

ttk.Label(mainframe, wraplength=450, anchor=CENTER, text='Guidlines for best writing statistics results:', width=100).grid(column=2, row=1, pady=(0,10))
ttk.Label(mainframe, wraplength=450, anchor=CENTER, text='- Only file paths ending with .docx (Word documents) can be analyzed. To get a documents file path, right click the file and select "properties" and copy-paste the file path. Make sure to add the file name and extension ".docx" to the end of the file path if it is not already there.', width=100).grid(column=2, row=2, pady=(0,10))
ttk.Label(mainframe, wraplength=450, anchor=CENTER, text='- Any text outside of a paragraph (title, authorship line, etc.) will be counted as its own paragraph and sentence. Remove these lines of text if you do not wish for them to be counted as such.',width=100).grid(column=2, row=3, pady=(0,10))
ttk.Label(mainframe, wraplength=450, anchor=CENTER, text='- Remove blank lines of text if you do not wish them to be counted as paragraphs', width=100).grid(column=2, row=4, pady=(0,10))
ttk.Label(mainframe, wraplength=450, anchor=CENTER, text='- Footnotes will be read only if converted to regular text and put at the end of the document/text', width=100).grid(column=2, row=5, pady=(0,10))
ttk.Label(mainframe, wraplength=450, anchor=CENTER, text='- For text that is copied/pasted, remove any extra spaces at the end of the text or else they will be counted as extra paragraphs', width=100).grid(column=2, row=6, pady=(0,100))


root.mainloop()



        ##Notes##


  # - does not 'read' footnotes - convert to regular text at end of docx before uploading file
  # - Remove Title/cover page/author line/etc. before uploading docx. Extra lines. Title will be read as paragraph and a sentence
  #   of this kind will result in inflated paragraph count and deflated average sentences per
  #   paragraph count
  # - File input can only be docx (Word Document)
  # - For copy and paste text, if there is a space at the end of the text, the program will read it as an extra paragraph
 

  # - Some symbols make POS tagging not work (ex. *)


  #Create error message when string index out of range error
  #graphs for words_per_sentence, sentences_per_paragraph, POS frequency of first word in sentences, most frequent words for each POS 
  


''' Key for POS tags from nltk     * = used, - = unused

Number  Tag     Description
1.	CC	Coordinating conjunction *
2.	CD	Cardinal number - 
3.	DT	Determiner (articles) *
4.	EX	Existential there
5.	FW	Foreign word - 
6.	IN	Preposition or subordinating conjunction *
7.	JJ	Adjective *
8.	JJR	Adjective, comparative *
9.	JJS	Adjective, superlative *
10.	LS	List item marker - 
11.	MD	Modal - 
12.	NN	Noun, singular or mass *
13.	NNS	Noun, plural *
14.	NNP	Proper noun, singular *
15.	NNPS	Proper noun, plural *
16.	PDT	Predeterminer *
17.	POS	Possessive ending - 
18.	PRP	Personal pronoun *
19.	PRP$	Possessive pronoun *
20.	RB	Adverb *
21.	RBR	Adverb, comparative *
22.	RBS	Adverb, superlative *
23.	RP	Particle  -
24.	SYM	Symbol  -
25.	TO	to  -
26.	UH	Interjection *
27.	VB	Verb, base form *
28.	VBD	Verb, past tense *
29.	VBG	Verb, gerund or present participle *
30.	VBN	Verb, past participle *
31.	VBP	Verb, non-3rd person singular present *
32.	VBZ	Verb, 3rd person singular present *
33.	WDT	Wh-determiner (articles) *
34.	WP	Wh-pronoun *
35.	WP$	Possessive wh-pronoun *
36.	WRB	Wh-adverb *

'''



