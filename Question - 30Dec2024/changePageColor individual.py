"""
Author:      Vera Neroni
Date Created:     12/27/2024

Purpose:
        - Change page color of a list of Word documents automatically
        - To be used to change spec color in digital bid packages
        --- helps avoid having to look up and type in the color codes all the time
        --- maybe make the templates blue?? no... white is better for editing...

Edited:
        - 12/27/2024 - trying to edit for color visibility. Applied StackOverflow solutions.


Sources:
        Google AI Overview search
        *query: "python change word page color to specific rgb value"
        *link: https://www.google.com/search?q=python+change+word+page+color+to+specific+rgb+value&sca_esv=03932fef72e18384&sxsrf=ADLYWIID9YTacHRS3n6rXTGp21GO3S_FsQ%3A1735346338579&ei=okhvZ5-EI82rptQPkKDh0Q8&ved=0ahUKEwif3crwnMmKAxXNlYkEHRBQOPoQ4dUDCBA&uact=5&oq=python+change+word+page+color+to+specific+rgb+value&gs_lp=Egxnd3Mtd2l6LXNlcnAiM3B5dGhvbiBjaGFuZ2Ugd29yZCBwYWdlIGNvbG9yIHRvIHNwZWNpZmljIHJnYiB2YWx1ZTIFECEYoAEyBRAhGKABSMogUNwEWIofcAF4AZABAJgBugKgAYUdqgEIMC4xOS4yLjG4AQPIAQD4AQGYAhegAo4ewgIKEAAYsAMY1gQYR8ICBRAhGJ8FwgIFECEYqwKYAwCIBgGQBgOSBwgxLjE3LjQuMaAH05IB&sclient=gws-wiz-serp
        *output:
        'Unfortunately, directly changing the page color of a Word document using Python is not straightforward.
        The python-docx library, which is commonly used for manipulating Word documents, primarily focuses on text, paragraphs,
        and styles, and it doesn't have direct support for modifying page-level properties like background color.'
        *It have a code example though:
        'import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open("your_document.docx")
        # Assuming you want an RGB value of (255, 128, 0)
        doc.Background.Fill.ForeColor.RGB = win32com.client.RGB(255, 128, 0)
        doc.Save()
        doc.Close()
        word.Quit()'
        *This didn't work - aaaaah (error says win32com.clinet does not have "RGB" method)

        - Results from StackOverflow
        https://stackoverflow.com/questions/56089015/need-to-change-color-of-rectangle-in-ms-word-using-win32com
        says to replace RGB with something else
        -more details on how to create replacement
        https://stackoverflow.com/questions/11444207/setting-a-cells-fill-rgb-color-with-pywin32-in-excel/21338319#21338319
        ---used both of these combined solutions, and it worked! mostly the second one
        --- needed adjustments from SO article below to work though

        - Maybe this will help with color visibility? (IT DID!)
        https://stackoverflow.com/questions/27761097/change-the-background-color-of-a-word-file-via-powershell
        See: answer by Matt, Jan 4 2015 at 1:09, and the original post by Shaya
        -used the suggsetions to
        --- set background visibility to "True" (Matt)
        --- set active window view to online (Shaya)

        For time elapsed:
        Google AI Overview
        *query: "python start time"
        *output:
        'import time
        start_time = time.time()
        # Your code here
        end_time = time.time()
        print(f"Elapsed time: {end_time - start_time} seconds")'



Results:
        - Ran without errors, but no change in color....

        Basically, this doesn't work.
        It seemed that the result is not a background color, even if you try a Macro
        or save it to a style.
        Also, after doing this, it does not even let you change the color manually.

        Have to make the page not have a color at all, save, then do the color manually again.

        Not doing this anymore...

        UPDATE: Sort of works????
        It changed the colour to blue when I opened the document, but it was in online view. When you switch to the page, you can't see the colour.

        UPDATE2: IT CONVERTS TO COLOUR PDF OK! Looks blue.
        Amazing.


"""

import win32com.client


#######################
# Testing
#######################


#initialize a Word app object (I think? similar to C#)
word = win32com.client.Dispatch("Word.Application")
print("Initialized Word application object")

###user input for doc location
##input_doc = input("Input path for Word doc location: ")
##clean_input_doc = input_doc.strip('"')

#initialize a Word doc object
#doc = word.Documents.Open(clean_input_doc)
doc = word.Documents.Open(r"C:\Users\vnner\Desktop\Test2\test1.docx")
print("Initialized doc file object")

#input is a tuple - a 3-digit RGB color value
def rgbToInt(rgb):
    """Function that converts three-value color code to integer color code.
    """
    print("Creating a unified integer color value")
    #formula: Red + (Green * 256) + (Blue * 256 * 256)
    color_int = rgb[0] + abs(rgb[1] * 256) + (rgb[2] * 256 * 256)
    print(f"The new color value is {color_int}")
    return color_int


#change the background
print("Starting to change the background color...")
doc.Background.Fill.ForeColor.RGB = rgbToInt((169, 233, 245))

#make colour visible
doc.Background.Fill.Visible = True

#switch to Online Layout vies to be able to see the colour
doc.ActiveWindow.View.Type = 6


print("If you see this, code completed")


doc.Save()
doc.Close()
word.Quit()
