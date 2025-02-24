# Katip4
Add-in to support spellchecking in Microsoft Word desktop application using Hunspell library and dictionaries. 

Many languages used in the world do not having spellchecking functionality available to them. Hunspell is an open-source library used by thousands of people. The purpose of this add-in is to make Hunspell available to Microsoft Word users. 

# Installation

To install Katip you need to find out whether your Word (not Windows OS) installation is 32-bit or 64-bit. 

## Find out bitness of your Word installation

### Word 2010

- Go to File > Help and look on the right. It will show a text similar to this "Version: xx.x.xxx.xxxx (64-bit)"

### Word 2016

- Go to File > Account and look on the right. Click on About Word to find out bitness.

## Choose installation file

- For 32-bit Word use katip_setup_x86.exe
- For 64-bit Word use katip_setup_x64.exe

## Install your language's dictionary

Depending on your operating system Katip installs data folder in one of these locations:
- C:\ProgramData\Katip

Go to the program data folder and copy your dictionary files into Katip\dictionaries folder. For each language you need to install two files "language_name.aff" and "language_name.dic". For example, for Hungarian you need to install "hu-HU.aff" and "hu-HU.dic". The name of the files must match this format. If not Katip will not be able to recognize the language. 

To find out the language tag for your language you can look up "languages.txt" file in the data folder. If your language is not listed you can submit a request to the author.

# Supported systems

Currently tested Windows systems:
- Windows 11 64-bit
- Windows 7 64-bit

Currently tested Word versions:
- Word 2010 64-bit
- Word 2010 32-bit

More tests will be added in the future.

# Known issues

- Spellchecking doesn't integrate into Word's native proofing functionality. It runs in a separate window.
- Currently it works only for plain text content of the main body. It doesn't support headers, footers or other special types of content.
- Quality of spellchecking depends on the quality of Hunspell dictionaries for each language.
- Misspelling color is limited to only standard colors, despite being chosen from standard Font dialog window of Word.