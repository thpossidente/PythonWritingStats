            ##### Writing Stats #####


            ##Helper Functions and Libraries##



#import('nltk') # doesnt work



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

# text = input('Enter text: ')  Use only in absence of word file input
def writing_stats(text):
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
  first_word = {'pronoun' : 0, 'noun' : 0, 'verb' : 0, 'adverb' : 0, 'conjuction' : 0, 'preposition' : 0, 'interjection' : 0, 'adjective' : 0, 'articles' : 0}
  end_marks = {'period' : 0, 'exclamation point' : 0, 'question mark' : 0}
  other_punctuation = {'comma' : 0, 'semicolon' : 0, 'colon' : 0, 'ampersand' : 0, 'hyphen' : 0, 'exclamation point' : 0, 'period' : 0, 'dash' : 0, 'question mark' : 0, 'parenthesis' : 0, 'quotation mark' : 0, 'ellipses' : 0, 'forward slash' : 0, 'period' : 0}
  voice_per_sentence = {'passive' : 0, 'active' : 0, 'unsure': 0}
  passive_current_sentence = 0

  text = text.replace('\n\n', '@@')
  text = text.replace('\u201c', '"')
  text = text.replace('\u201d', '"')
  text = text.split(' ')
  if text[len(text) - 1] == '':
      del text[len(text) - 1]
          
  print(text)
  
  for nxt, word in zip(text[1:]+['i'], text):  
    if word != '-' and word!= '&':
      word_count += 1
      words_current_sentence += 1

    if "@@" in word:
      paragraph_count += 1
      sentences_per_paragraph.append(sentences_current_paragraph)
      sentences_current_paragraph = 0
      letter_count -= 2
      letters_current_word -= 2
      
    if word == '-':
      other_punctuation['dash'] += 1
      letter_count -= 1

    if word == '&':
        other_punctuation['ampersand'] += 1
        letter_count -= 1
    
    lower = word.lower()
    lower_nxt = nxt.lower()
    
    if lower.strip('.') == 'was' or lower.strip('.') == 'were' or lower.strip('.') == 'be' or lower.strip('.') == 'been':
      passive_current_sentence += 1
  
    if (lower.strip('.') == 'has' and lower_nxt.strip('.') == 'been') or (lower.strip('.') == 'is' and lower_nxt.strip('.') == 'being') or (lower.strip('.') == 'have' and lower_nxt.strip('.') == 'been') or (lower.strip('.') == 'was' and lower_nxt.strip('.') == 'being') or (lower.strip('.') == 'were' and lower_nxt.strip('.') == 'being') or (lower.strip('.') == 'had' and lower_nxt.strip('.') == 'been'):
      passive_current_sentence += 3

    if (word[len(word) - 1] == '.' and (nxt[0].isupper() or nxt[0] == '"' or nxt[0:2] == "@@" or word == text[len(text) - 1])):
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
    
    if (word[len(word) - 1] == '?' and (nxt[0].isupper() or nxt[0] == '"' or nxt[0:2] == "@@" or word == text[len(text) - 1])):
      sentence_count += 1
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
      
    if (word[len(word) - 1] == '!' and (nxt[0].isupper() or nxt[0] == '"' or nxt[0:2] == "@@" or word == text[len(text) - 1])):
      sentence_count += 1
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
      
    if word[len(word) - 1] == '.' and (nxt[0].isupper() == False and nxt[0] != '"') and word != text[len(text) - 1]:
      other_punctuation['period'] += 1
      
    if len(word) > 3 and '...' in word: 
      other_punctuation['ellipses'] += 1
      
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
        
      list_letters = list(enumerate(word))
      if letter == list_letters[len(list_letters) - 1][1]:
        letters_per_word.append(letters_current_word)
        letters_current_word = 0

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
      percentages_voice = {'Percent passive' : percent_passive, 'Percent active' : percent_active, 'Percent unsure' : percent_unsure}
  
  print('\n', 'Paragraph count: ' + str(paragraph_count),'\n',
        'Word count: ' + str(word_count),'\n',
        'Letter count: ' + str(letter_count),'\n',
        'Sentence count: ' + str(sentence_count),'\n',
        'Average letters per word: ' + str(ave_letters_per_word),'\n',
        'Average words per sentence: ' + str(ave_words_per_sentence),'\n',
        'Average sentences per paragraph: ' + str(ave_sentences_per_paragraph),'\n',
        'End mark frequencies: ' + str(end_marks),'\n',
        'Punctuation Frequencies (excluding end marks): ' + str(other_punctuation),'\n',
        'Percent of sentences identified to be active, passive, and undetermined: ' + str(percentages_voice)
       )
  print(words_per_sentence)

writing_stats(get_docx_text("C:\\Users\\Tom\\Downloads\\Test Doc (1).docx"))
#writing_stats(text)   Use if no word doc input (comment out line above)



        ##Notes##

  #repetitive verbs, nouns, adverbs, adjectives
  #passive vs active - make better using POS tagging for is/was/were/be/ + past tence as passive
  #part of speech - using nltk
  #graphs for words_per_sentence and sentences_per_paragraph
  #does not 'read' footnotes - convert to regular text before uploading file
  #need to import nltk, try from __, import ___






