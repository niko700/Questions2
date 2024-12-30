"""
This is an example of the individual .py code that runs a template.


"""

from pathlib import Path
import os
from docxtpl import DocxTemplate
from make_classes import *
from choose_current import current as cr


#choose set of data for filling out templates
for project in cr.current_projects:

    print("Working on template for " + project.shortname + "...")


    #where template is coming from
    document_path = Path(__file__).parent / 'InFiles' / 'Template Docs' / 'doc1.docx'


    #set save location
    #currently, the document parent path is the same folder as this Python file is in
    #this line assigns the template to "doc"
    doc = DocxTemplate(document_path)


    #----- FILL OUT TEMPLATE -------
    #make dictionary based on variables made in the template (all caps) and the corresponding values from the project class
    context = {
       "VAR1" : project.data1,
       "VAR2" : project.data2,
       ## etc
        }


    #create output file -------------
    doc.render(context)


    #ADD IF NAME EXISTS CHECK
    counter = 0

    #save output file ------------
    name = project.shortname

    thisFile = Path(__file__).parent/ "OutFiles" / "Filled Out Docs" / ("doc1_" + name + ".docx")

    ###check current path
    ##print("Path to save to: " + thisFile)

    if os.path.isfile(thisFile):
        print("File already exists")
        while os.path.isfile(thisFile):
                counter += 1
                thisFile = Path(__file__).parent/ "OutFiles" / "FIlled Out Docs" / ("doc1_" + name + str(counter) + ".docx")
        ###check new path
        ##print("Path after existing file check: " + thisFile)
        print("Saving to new path: ")
        print(thisFile)


    #save the doc to new path
    doc.save(thisFile)


    print("Done!")

