# Katip

Katip is an add-in designed to enhance spellchecking capabilities in the Microsoft Word desktop application using the Hunspell library and dictionaries.

Many languages around the world lack built-in spellchecking functionality and are not supported by major software vendors. The goal of this add-in is to integrate Hunspell's robust spellchecking features into Microsoft Word, thereby enabling spellchecking for low-resource languages. 

[Hunspell](https://hunspell.github.io/) is a globally used open-source library.

# Features

- Access the UI via the Word Ribbon menu.
- Edit spelling using a dialog window.
- Edit spelling using right-click context menu.
- Change/Change All misspelled words with suggestions.
- Ignore Once/Ignore All options for the current spellchecking session.
- Add words to the permanent user dictionary.
- Change the UI display language of the add-in.
- Change the spelling language based on available languages.
- Toggle auto spellcheck and auto clear spellcheck functionalities on/off.
- Customize the color and underline style used for highlighting misspellings.
- Load additional dictionaries from a specified folder.
- Navigate through misspellings efficiently.
- Customize the set of characters considered as word separators.
- Save settings permanently.
- Log errors to a text file.

# Installation

To install Katip, you need to determine whether your Word installation (not Windows OS) is 32-bit or 64-bit.

## Determine the Bitness of Your Word Installation

### Word 2010

- Go to **File** > **Help**.
- Look on the right side. You will see text similar to "Version: xx.x.xxx.xxxx (64-bit)".

### Word 2016

- Go to **File** > **Account**.
- Click on **About Word** to find out the bitness.

## Choose the Installation File

- For 32-bit Word, use `katip_setup_x86.exe`.
- For 64-bit Word, use `katip_setup_x64.exe`.

## Install Your Language's Dictionary

Depending on your operating system, Katip installs the data folder in the following location:

- `C:\ProgramData\Katip`

To install your language's dictionary, follow these steps:
1. Navigate to the Program Data folder.
2. Copy your dictionary files into the `Katip\dictionaries` folder. For each language, you need to install two files: `language_name.aff` and `language_name.dic`. 

For example, for Hungarian, you need to install `hu-HU.aff` and `hu-HU.dic`. The file names must match this format; otherwise, Katip will not recognize the language. The dictionaries must be in UTF-8 encoding and comply with the Hunspell standard.

To find the language tag for your language, refer to the `languages.txt` file in the data folder. If your language is not listed, you can submit a request to the author.

# Supported Systems

### Currently Tested Windows Systems:
- Windows 11 64-bit
- Windows 7 64-bit

### Currently Tested Word Versions:
- Word 2010 64-bit
- Word 2010 32-bit
- Word 2016 64-bit

More tests will be added in the future.

# Creating Your Own Build

You can build your own release using the sources in the `x86` and `x64` folders. Here are some steps to consider:

- Browse through the code files to review the code.
- Modify `katip4.dotm` using the Word VBA IDE to change functions, forms, and modules.
- Add or remove dictionary files in the `/dictionaries` folder.
- Add or remove UI localization files in the `/locale` folder.
- Modify the default launch settings in the `settings.ini` file.
- Run the provided InnoSetup installer creation script. Adjust the paths and names in the file if you made any changes.

# Known Issues

- Spellchecking doesn't integrate into Word's native proofing functionality. It runs in a separate window.
- Currently, it works only for plain text content in the main body. It doesn't support headers, footers, or other special types of content.
- The quality of spellchecking depends on the quality of Hunspell dictionaries for each language.
- The misspelling color is limited to only standard colors, even if chosen from the standard Font dialog window of Word.
- Misspelling style is not fully visible in Settings window.
- Navigation through errors can be erratic if used too dynamically.

# Contribution

Contributions and feedback are welcome.

# License

Katip is available under the GPL License. Katip uses [Hunspell](https://github.com/hunspell/hunspell), governed by its respective licenses.