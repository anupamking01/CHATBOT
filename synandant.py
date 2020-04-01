import nltk
from win32com.client import Dispatch
from nltk.corpus import wordnet

speak = Dispatch("SAPI.SpVoice")
synonyms = []
antonyms = []

for syn in wordnet.synsets("good"):
	for l in syn.lemmas():
		synonyms.append(l.name())
		if l.antonyms():
			antonyms.append(l.antonyms()[0].name())

print(set(synonyms))
speak.Speak(set(synonyms))

print(set(antonyms))
speak.Speak(set(antonyms))

