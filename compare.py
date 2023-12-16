# -*- coding: utf-8 -*-
"""
Created on Tue Oct 24 12:47:14 2023

@author: vzocc
"""

import os
import filecmp
import difflib

import win32com.client as win32

import PySimpleGUI as psg
import tkinter as tk

import glob



#*****************************************************************************
# @brief Returns all text files in directory
# @param directory
#
# @return list of files
#***************************************************************************** 
def return_all_files(d):

    files = glob.glob(d + "*.txt")
    files = [f.replace('\\', '/') for f in files]

    files = [f for f in files if 'corrected' not in f ]
    files = [f for f in files if 'orig' not in f ]
    
    return files

    
def open_word_document(file_path):
    word = win32.gencache.EnsureDispatch("Word.Application")
    print(file_path)
    doc = word.Documents.Open(file_path)
    return word, doc


def highlight_differences(original_file, modified_file):
    # Read the content from the original file
    with open(original_file, 'r') as original:
        original_content = original.readlines()

    # Read the content from the modified file
    with open(modified_file, 'r') as modified:
        modified_content = modified.readlines()

    # Compare the content and generate differences
    differ = difflib.Differ()
    diff = list(differ.compare(original_content, modified_content))

    highlighted_diff = []
    
    for line in diff:
        if line.startswith('  '):
            # Unchanged line
            highlighted_diff.append(f"  {line[2:]}")
        elif line.startswith('- '):
            # Line only in the original file (removed)
            highlighted_diff.append(f"- {line[2:]}")
        elif line.startswith('+ '):
            # Line only in the modified file (added)
            highlighted_diff.append(f"+ {line[2:]}")
    
    return highlighted_diff

def generate_modifications_dictionary(original_file, modified_file):
    # Read the content from the original file
    with open(original_file, 'r') as original:
        original_content = original.read()

    # Read the content from the modified file
    with open(modified_file, 'r') as modified:
        modified_content = modified.read()

    # Split the content into lines
    original_lines = original_content.splitlines()
    modified_lines = modified_content.splitlines()

    # Initialize a dictionary to store modifications
    modifications = {}

    # Compare original and modified lines
    for line_num, (original_line, modified_line) in enumerate(zip(original_lines, modified_lines)):
        if original_line != modified_line:
            modifications[f"Line {line_num + 1}"] = {
                "Original": original_line,
                "Modified": modified_line
            }

    return modifications


def flatten_list(mylist):
    flattened_list = []  # Create an empty list to store the flattened result

    for item in mylist:
        if isinstance(item, list):
            # If the current item is a list, recursively flatten it and extend the result
            flattened_list.extend(flatten_list(item))
        else:
            # If the current item is not a list, add it to the flattened list
            flattened_list.append(item)

    return flattened_list  # Return the flattened list as a new variable

#*****************************************************************************
# @brief GUI to choose file
# @param None
#
# @return number of words in file
#*****************************************************************************         
def file_chooser(text='Enter the file you wish to process'):
    psg.theme('Black')
    filename = psg.popup_get_file(text)
    #psg.popup('You entered', filename)
    return filename

#*****************************************************************************
# @brief Compare two text files
# @param paths to the two files
#
# @return comparison, file1, file2
#*****************************************************************************  
def compare(file1, file2):
    # Reading files
    f1_path = file1
    f2_path = file2
    
    with open(f1_path, "r") as f1, open(f2_path, "r") as f2:
        text1 = f1.read()
        text2 = f2.read()
    with open(f1_path, "r") as f1, open(f2_path, "r") as f2:
        f1_lines = f1.readlines()
        f2_lines = f2.readlines()
        
        f1_lines = [l for l in f1_lines if l != '\n']
        f2_lines = [l for l in f2_lines if l != '\n']
        
        #f1_lines = f1_lines[:5]
        #f2_lines = f2_lines[:5]

    result = []

    for line1, line2 in zip(f1_lines, f2_lines):
        # Tokenize the lines into words
        f1_words = line1.split()
        f2_words = line2.split()
    
        # Perform text comparison
        d = difflib.Differ()
        diff = list(d.compare(f1_words, f2_words))
        
        word_diff = []
        
        for word in diff:
            if word.startswith(' '):
                word_diff.append(word.strip())
            if word.startswith('- '):
                #word_diff.append(f'\033[1;31;40m{word[2:]}\033[0m')
                #result.append(f'\x1B[3m{word[2:]}\x1B[0m')
                word_diff.append(f"+{word[2:]}")
            elif word.startswith('+ '):
                #result.append(f'\x1B[4m{word[2:]}\x1B[0m')
                #word_diff.append(f'[{word[2:]}]')
                word_diff.append(f'-{word[2:]}')
            
                
        # Join the word differences and add the line ending
        result.append(' '.join(word_diff) + '\n')

    diff = ''.join(result)
    diff = diff.replace('  ', ' ')
    diff = diff.replace('  ', ' ')
    
    with open("compare.txt", 'w') as f:
        f.write(diff)
    
    return text1, text2, diff

#*****************************************************************************
# @brief Create text commands to colourise text based on the +/- values
# @param The text with the differences highlighted by +/-
#
# @return The text commands for the window
#***************************************************************************** 
def colourise_text(text):
    coloured_text = []
    cache = []
    words = text.split()
    
    last = -2

    for word in words:
        word = word.replace("'", "\\'") #to avoid errors
        
        plus =  "window['-ML3-'].print('"+' '.join(cache)+" ', text_color='green', background_color='black', end='')"
        minus = "window['-ML3-'].print('"+' '.join(cache)+" ', text_color='red', background_color='black', end='')"
        eq =    "window['-ML3-'].print('"+' '.join(cache)+" ', text_color='white', background_color='black', end='')"
        
        if word.startswith('+'):
            if last == 1:
                cache.append(word[1:])
            elif last == -1:   
                coloured_text.append(minus)
                cache = []
                cache.append(word[1:])
            elif last == 0:   
                coloured_text.append(eq)
                cache = []
                cache.append(word[1:])
            last = 1
        elif word.startswith('-'):
            if last == -1:
                cache.append(word[1:])
            elif last == 1:   
                coloured_text.append(plus)
                cache = []
                cache.append(word[1:])
            elif last == 0:   
                coloured_text.append(eq)
                cache = []
                cache.append(word[1:])
            last = -1
        else:
            if last == 0:
                cache.append(word)
            elif last == -1:   
                coloured_text.append(minus)
                cache = []
                cache.append(word)
            elif last == 1:   
                coloured_text.append(plus)
                cache = []
                cache.append(word)
            last = 0

    return coloured_text
    
def execute_command(text, window):
    #for command_string in text:
        text = '\n'.join(text)
        
        try:
            # Execute the command string as code
            exec(text)
        except Exception as e:
            print(f"Error executing the command: {e}")
        
#*****************************************************************************
# @brief Create a window showing the two files and the differences between them
# @param paths to the two files and their differences
#
# @return None
#*****************************************************************************      
def text_window(file1, file2, diff):
    layout = [  [psg.Text('Demonstration of Multiline Element Printing')],
                [psg.MLine(key='-ML1-'+psg.WRITE_ONLY_KEY, size=(50,10))],
                [psg.MLine(key='-ML2-'+psg.WRITE_ONLY_KEY,  size=(50,10))],
                [psg.MLine(key='-ML3-',  size=(50,20))],
                [psg.Button('Exit'), psg.Button('Save')]]

    window = psg.Window('Window Title', layout, finalize=True)

    # Note, need to finalize the window above if want to do these prior to calling window.read()
    window['-ML1-'+psg.WRITE_ONLY_KEY].print(file1, text_color='red', background_color='yellow')
    window['-ML2-'+psg.WRITE_ONLY_KEY].print(file2, text_color='green', background_color='yellow')
    #window['-ML3-'+psg.WRITE_ONLY_KEY].print(diff, text_color='white', background_color='black')
    #window['-ML3-'].print(diff, text_color='white', background_color='black')

    coloured_text = colourise_text(diff)
    execute_command(coloured_text[:20], window)

    while True:             # Event Loop
        event, values = window.read(timeout=100)
        if event in (psg.WIN_CLOSED, 'Exit'):
            break
        # To save the contents of the Multiline element to a file
        if event == 'Save':
            try:
                with open("test.txt", "w") as file:
                    file.write(values['-ML3-'])
            except Exception as e:
                print(f"Error: {e}")
            
    window.close()

    
def set_text(text):
    font = ('Courier New', 11)
    psg.theme('Black')
    psg.set_options(font=font)
    
    layout = [  [psg.Multiline(size=(50, 10), key='textbox', expand_x=True, expand_y=True)],
                [psg.Text(text)],
                [psg.Button('Exit')]]
    
    window = psg.Window('Window Title', layout, resizable=True, finalize=True)
    
    #set_text(text)
    
    while True:
        event, values = window.read()
        if event in (psg.WIN_CLOSED, 'Exit'):
            break

#f1_path = file_chooser()
#f2_path = file_chooser()

f1_path = 'C:/Users/vzocc/Documents/GitHub/The-Wall/Chapters/Chapter0/The Hole in the Wall - Ch0.orig.txt'
f2_path = 'C:/Users/vzocc/Documents/GitHub/The-Wall/Chapters/Chapter0/The Hole in the Wall - Ch0.txt'

file1, file2, diff = compare(f1_path, f2_path)    

    
text_window(file1, file2, diff)


