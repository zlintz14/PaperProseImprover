# About PaperProseImprover
This program is intended to improve your word choice in papers or any .docx file by highlighting words you specify and adding a link to thesarus.com. 

# How to Use
It will prompt you for a number, that number will represent the minimum number of occurences of a word in your .docx file in order for it to be highlighted. For example, if you enter 5, every word that appears 5 or more times in the paper, will be highlighted in a unique color and linked to thesarus.com. I have filtered out some common words such as "this, a, the, etc.", please contact me if you feel there's another common word I should filter out. The program can only read in .docx files. If you happen to try and open a non .docx file, simply close the program and reopen. 

# How it Works
A basic overview of the program, it is written in Java and I have some external libraries I imported called Apache POI in order to read in Microsoft Word .docx files and edit them. After reading in the Word docx, it removes all punctuation from the words and makes them all lower case for comparisions, then it puts all these edited words in a Map and associates them with the number of times each word occurs. Next it checks to see if the word is a common word, and if it's not, the program checks the number of occurences of every word against the number input by the user. If the word occurs a number of times that is greater than or equal to the user specified number for minimum occurences, then a random color is generated and the color has a number of conditionals that ensures it's a nice, vibrant color that isn't too light or too dark. The word is then put into a new Map just for highlighted words and their accompanying colors. Lastly, in the main class, the words are written back onto a new .docx file called Edited docx.docx, adding a URL linked to thesaurus.com for the specific word, for all highlighted words. 

# Other Important Info About the Program
YOU MUST HAVE MICROSOFT WORD INSTALLED TO RUN THIS PROGRAM. Also you must have the Apache POI downloaded and used as a reference library; to see the specific classes used from the Apache POI, check the import statements on the two classes. Unfourtunately, I was not able to get the Edited docx.docx to be double spaced, but you can easily do that yourself. I also couldn't get the links to show up on the actual highlighted word so the same word is printed next to the highlighted word and it is colored in blue and underlined like a normal link would be. My program will get the formatting close to the orignal formatting but not exactly alike (it won't underline, italicize, etc.). The program will create the Edited docx.docx file and save it in the same place the jar file for the program is saved, the program WILL NOT alter your original file. The program will automatically open the Edited docx.docx in Mircrosoft Word. If you run the program more than once, the exisitng Edited docx.docx file will simply be overwritten. The font size and style will be your defualt Microsoft Word size and style. Lastly please note that, for example if you say find every word that appear 15 times but there weren't any words that appear 15 times, nothing will happen so just realize if nothing happens it's probably because of this reason.

# The jar file is the executable jar for this program so download that if you just wish to use the program.

# Please contact me with any common words you would like to see filtered out (there are plenty I don't have in there yet, just added a few for starters) or any bugs and I will try and fix ASAP!

# Enjoy! 
![alt text](https://media.giphy.com/media/DfSLII45H40RW/giphy.gif)

