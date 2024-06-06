"""
This is an example of code that used to work, but does not work anymore.

General Setup:
Each doc.py has code inside that looks for its respective doc1_template.docx, fills it out, and saves a filled out version.
Everything lives in C:\\Users\\user\\Desktop\\MainFolder.


Issue:
When main directory had all the .py files and .docx templates, the following code worked.

New setup is:
C:\\Users\\user\\Desktop\\MainFolder --> has all .py files
C:\\Users\\user\\Desktop\\MainFolder\\ChildFolder --> has all the .docx files
Each .py file has a path to the .docx file when it looks for the template, and this works when run individually.
This does not work when running this run_all.py from "MainFolder" directory where it is currently saved.

"""

import subprocess


#print doc package
print("Printing doc package documents")

subprocess.run(["python", "doc1.py"])
print("Creating ouptput 1")
subprocess.run(["python", "doc2.py"])
print("Creating ouptput 2")
subprocess.run(["python", "doc3.py"])
print("Creating ouptput 3")
subprocess.run(["python", "doc4.py"])
print("Creating ouptput 4")
subprocess.run(["python", "doc5.py"])
print("Creating ouptput 5")