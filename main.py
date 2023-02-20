import requests
import json
from tkinter import *
import openpyxl

# Your API key acquired for "Merriam-Webster's Learner's Dictionary with Audio"
API_KEY = ""

with open("vocabulary.json") as f:
    vocab_list = json.loads(f.read())
vocab_num = len(vocab_list)


def export_xl():
    global vocab_list, vocab_num
    wb = openpyxl.load_workbook("vocabulary.xlsx")
    ws = wb.create_sheet()

    for vocab_dict in vocab_list:
        row = [vocab_dict["hwd"]] + vocab_dict["def"]
        ws.append(row)
    wb.save("vocabulary.xlsx")

    vocab_list = []
    with open("vocabulary.json", "w") as f:
        f.write(json.dumps(vocab_list, ensure_ascii=False))
    vocab_num = 0
    l_export_desc.config(text=f"New {vocab_num} words stored. Would you like to export to xlsx file?")


def search_def(search_word):
    url = f"https://www.dictionaryapi.com/api/v3/references/learners/json/{search_word}?key={API_KEY}"
    res = requests.get(url)
    json_data = res.json()
    def_list = json_data[0]["meta"]["app-shortdef"]["def"]
    new_def_list = []
    for item in def_list:
        new_def_list.append(item.replace("{bc}", ""))
    return new_def_list


def add_vocab():
    new_headword = e_headword.get()
    if e_definition.get().strip() != "":
        new_definition = [e_definition.get()]
    else:
        new_definition = search_def(new_headword)

    new_dict = {
        "hwd": new_headword,
        "def": new_definition
    }

    global vocab_list
    vocab_list.append(new_dict)

    with open("vocabulary.json", "w") as f:
        f.write(json.dumps(vocab_list, ensure_ascii=False, indent=2))

    global vocab_num
    vocab_num += 1
    l_export_desc.config(text=f"New {vocab_num} words stored. Would you like to export to xlsx file?")
    e_headword.delete(0, END)
    e_definition.delete(0, END)



win = Tk()
win.title("Vocabulary List")
win.config(padx=20, pady=20)

l_headword = Label(text="Headword: ")
l_definition = Label(text="Definition: ")
l_note = Label(text="*Without definition, it will be automatically generated from dictionary",
               fg="grey")
e_headword = Entry()
e_definition = Entry()
b_add = Button(text="Add", command=add_vocab)
l_export_desc = Label(text=f"New {vocab_num} words stored. Would you like to export to xlsx file?",
                      fg="grey")
b_export = Button(text="Export", command=export_xl)

l_headword.grid(row=0, column=0, sticky="e")
l_definition.grid(row=1, column=0, sticky="e")
l_note.grid(row=2, column=0, columnspan=2)
e_headword.grid(row=0, column=1, sticky="w")
e_definition.grid(row=1, column=1, sticky="w")
b_add.grid(row=3, column=0, columnspan=2)
l_export_desc.grid(row=4, column=0, columnspan=2, pady=(30, 0))
b_export.grid(row=5, column=0, columnspan=2)


win.mainloop()
