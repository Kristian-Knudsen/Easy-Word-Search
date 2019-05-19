#!/usr/bin/env python
import docx
import glob
import win32com.client
import os
from tkinter import filedialog
from tkinter import *

def selectgui():
    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory()
    return folder_selected

def convfiles():
    # allows for converting .doc to .docx
    word = win32com.client.Dispatch("Word.Application") # opens a word client
    word.visible = 0 # not really sure
    for doc in glob.iglob("*.doc"): # loops through all files in the current folder, and finds .doc files
        doc = doc[:-4] # removes the .doc part of the filename to prevent being renamed with that ending on
        in_file = os.path.abspath(doc) # gets the name to open the file
        wb = word.Documents.Open(in_file) # opens the file
        out_file = os.path.abspath("{}.docx".format(doc)) #  creates a new name for the file
        wb.SaveAs2(out_file, FileFormat=16) # saves the files as .docx file (Format 16)
        wb.Close() # closes the file
    word.Quit() # closes word client

def main():
    os.chdir(selectgui())
    usrconv = input("Do you want to conv files? [y/n]")
    if usrconv == "y" or usrconv == "Y":
        convfiles()

    keyword = input("Search for >> ") # gets user keyword to search for - maybe make this a option for cmd to use also

    results = {}

    for file in glob.glob('./*.docx'): # loops through all .docx files in the folder
        results[file] = 0
        try:
            doc = docx.Document(file)  # opens a word with the current file
            for paragraph in doc.paragraphs:  # gets all paragraphs (all text)
                line = paragraph.text.split(" ")  # splits the lines
                if line == ['']:  # if the line is a space then remove it/pass it
                    pass  # passes
                else:  # if line is not a space
                    for word in line:  # get each word in the line to check for the keyword
                        if keyword in word:  # check if the keyword and the current word is the sam
                            #print("match at word {} in document {}".format(word, file))  # print a successmessage when
                            # a word is found in a file
                            results[file] = results[file] + 1
        except docx.opc.exceptions.PackageNotFoundError:
            print("File faulted - continuing")
            continue

    bkey, bvalue = max(results.items(), key = lambda x:x[1])
    print("The file with the most results is {} with a total of {} hits".format(bkey, bvalue))
    for key in results:
        if key == bkey:
            pass
        else:
            value = results[key]
            if value == 0:
                pass
            else:
                key = key[2:]
                print("Lesser hits is found in {} with {} hits".format(key, value))

while True:
    main()
    stop = input("Do you wish to search for more keywords? [y/n]")
    if stop == "n" or stop == "N":
    	break
