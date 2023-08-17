from docx import Document
import genanki
import random

deck_id = random.randrange(1 << 30, 1 << 31)
print(deck_id)

deck_name = input('Deck Name: ')
deck_name += '.apkg'

file_name = input('Docx file name: ')

if file_name.find('.docx') == -1:
    file_name += '.docx'

my_model = genanki.Model(
    deck_id,
    'Simple Model',
    fields=[
        {'name': 'Question'},
        {'name': 'Answer'},
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '{{Question}}',
            'afmt': '{{FrontSide}}<hr id="answer">{{Answer}}',
        },
    ])

my_deck = genanki.Deck(int(deck_id), deck_name)

doc = Document(file_name)

deck = genanki.Deck(deck_id, deck_name)

save = False

title = ""
info = ""
titles = []
infos = []

for paragraph in doc.paragraphs:

    if paragraph.style.name == 'Heading 1':
        continue

    if paragraph.style.name == 'Heading 2':
        if info == '':
            continue

        infos.append(info)
        info = ''
        continue

    info += paragraph.text + '\n'

infos.append(info)


for paragraph in doc.paragraphs:

    if paragraph.style.name == 'Heading 2':
        titles.append(paragraph.text)

print(len(infos))
print(len(titles))

for i in range(len(infos)):
    note = genanki.Note(
        model=my_model,
        fields=[titles[i], infos[i]])
    my_deck.add_note(note)

genanki.Package(my_deck).write_to_file(deck_name)
print('Finished')
