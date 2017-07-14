            ##### Writing Stats #####


            ##Helper Functions and Libraries##
import nltk
import collections
#nltk.download()



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
  pronouns = []
  nouns = []
  verbs = []
  adverbs = []
  conjunctions = []
  prepositions = []
  interjections = []
  adjectives = []
  articles = []
  first_word = {'pronoun' : 0, 'noun' : 0, 'verb' : 0, 'adverb' : 0, 'conjuction' : 0, 'preposition' : 0, 'interjection' : 0, 'adjective' : 0, 'articles' : 0}
  end_marks = {'period' : 0, 'exclamation point' : 0, 'question mark' : 0}
  other_punctuation = {'comma' : 0, 'semicolon' : 0, 'colon' : 0, 'ampersand' : 0, 'hyphen' : 0, 'exclamation point' : 0, 'period' : 0, 'dash' : 0, 'question mark' : 0, 'parenthesis' : 0, 'quotation mark' : 0, 'ellipses' : 0, 'forward slash' : 0, 'period' : 0}
  voice_per_sentence = {'passive' : 0, 'active' : 0, 'unsure': 0}
  passive_current_sentence = 0

  text = text.replace('\n\n', ' @@')
  text = text.replace('\u201c', '"')
  text = text.replace('\u201d', '"')
  text = text.split(' ')
  
  for word in text:
      if word == '':
          text.remove(word)


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

    if nxt == 'i':
      paragraph_count += 1
      sentences_per_paragraph.append(sentences_current_paragraph)
    
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
        
    lst_letters = list(enumerate(word))
    for letter in lst_letters:
      if letter == lst_letters[len(lst_letters)-1]:
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

  new_text = []
  for word in text:
    word = word.replace('@@', '')
    new_text.append(word)

          
  tagged_text = nltk.pos_tag(new_text)
  print(tagged_text)

  for nxt, word in zip(tagged_text[1:]+['i'], tagged_text):
      if word[1] == 'CC':
          conjunctions.append(word[0])
      if word[1] == 'WP' or word[1] == 'WPS' or word[1] == 'PRP' or word[1] == 'PRP$':
          pronouns.append(word[0])
      if word[1] == 'NN' or word[1] == 'NNS' or word[1] == 'NNP' or word[1] == 'NNPS':
          nouns.append(word[0])
      if word[1] == 'VB' or word[1] == 'VBD' or word[1] == 'VBG' or word[1] == 'VBN' or word[1] == 'VBP' or word[1] == 'VBZ':
          verbs.append(word[0])
      if word[1] == 'RB' or word[1] == 'RBR' or word[1] == 'RBS' or word[1] == 'WRB':
          adverbs.append(word[0])
      if word[1] == 'IN':
          prepositions.append(word[0])
      if word[1] == 'UH':
          interjections.append(word[0])
      if word[1] == 'JJ' or word[1] == 'JJR' or word[1] == 'JJS':
          adjectives.append(word[0])
      if word[1] == 'DT' or word[1] == 'WDT' or word[1] == 'PDT':
          articles.append(word[0])
  

  
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



writing_stats(get_docx_text("C:\\Users\\Tom\\Downloads\\1082CreationMyth.docx"))
#writing_stats(text)   Use if no word doc input and you want to be prompted to enter text (comment out line above)



        ##Notes##

  #graphs for words_per_sentence and sentences_per_paragraph
  #does not 'read' footnotes - convert to regular text before uploading file

  #POS counting most frequent in each POS
  #POS first word in sentence
  #passive vs active - make better using POS tagging for is/was/were/be/ + past tence as passive

  


''' Key for POS tags from nltk     * = used, - = unused

Number  Tag     Description
1.	CC	Coordinating conjunction *
2.	CD	Cardinal number
3.	DT	Determiner (articles) *
4.	EX	Existential there
5.	FW	Foreign word
6.	IN	Preposition or subordinating conjunction *
7.	JJ	Adjective *
8.	JJR	Adjective, comparative *
9.	JJS	Adjective, superlative *
10.	LS	List item marker
11.	MD	Modal
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



