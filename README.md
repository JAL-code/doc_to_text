# doc_to_text  

A work in progress . . . Convert Word 97-2003 Documents with verse to text conversion via a Linux Terminal.

This repo is a practical application of the olefile module provided by Phillippe Lagadec @ https://www.decalage.info/olefile.
For the GitHub version, see https://github.com/decalage2/olefile.

The code, for the example, is listed at the Decalage website in the comment by "Thu, 09/22/2011 - 18:08 — v3ss (not verified)".

Renamed extract_embedded_ole.py and reformatted to work with Python 3.8.5, the process is defined as followed:

The code processes a word file with a .doc extension at the /home/<user> folder.   The word file for this example is hard coded as winter_extend.doc. (Add picture). When the document stream is extracted, the script exports the data from the word file  and saves it at a custom folder including the file_name. (Add picture). When the generated file is opened in Text, half of the characters are flagged invalid but text can be seen between the vertical lines. (Add picture).
