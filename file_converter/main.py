#Imports
import os
from File_Converter import convert


#Upgrades
# DOCX to PDF


def main():
    greet()
    while True:
        conv_type = conversion_type()
        inp, out = ask_files()
        print(convert(inp, out, conv_type))
        print('Do you want to convert another file? (y/n)')



#Supporting functions
def greet():
    print('Welcome to the File Converter!')
    print('> This program allows you to convert files between different formats.')
    print('> Since this a CLI program everything should be perfectly clear (I mean the input_file name should be exact.)')
    print('> Currently Supporting: TXT <--> DOCX')

def conversion_type():
    """Ask the user what type of conversion they want to do, and check if the input is valid."""
    print('What type of conversion do you want to do?')
    print('> 1. TXT to DOCX')
    print('> 2. DOCX to TXT')
    print('> 3. TXT to PDF')
    while True:
        conversion_typ = int(input())
        if conversion_typ == 1:
            break
        elif conversion_typ == 2:
            break
        elif conversion_typ == 3:
            break
        else:
            print('Invalid input. Please enter 1, 2, or 3:')
            continue
    return conversion_typ


def ask_files():
    """Ask the user the input and output files, and check if the input file exists."""
    print('Please enter the file you want to convert:')
    while True:
        input_file = input()
        if os.path.isfile(input_file):
            break
        else:
            print('File not found. Type the whole file path if the file is not in this directory. Please enter a valid file path:')
            continue
    print('Please enter the file you want to save as:')
    while True:
        print('The output file should have the extension of the format you want to convert.')
        output_file  = input()
        if '/' in output_file or '\\' in output_file or ':' in output_file:
            print('The output file name should not contain the characters / \\ or :. Please enter a valid file name:')
            continue
        if output_file.endswith('.txt') or output_file.endswith('.docx') or output_file.endswith('.pdf'):
            break
        else:
            print('The output file should have the extension of the format you want to convert. Please enter a valid file name:')
            continue

    return input_file, output_file





#if __name__ == "__main__": main()
