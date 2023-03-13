# Written by Luke Darling

# Imports
from tkinter import *
from tkinter import filedialog, messagebox
import os, shutil
import win32com.client

# Define constants
IMPORT_FOLDER_NAME = "Raw"
EXPORT_FOLDER_NAME = "Combo"

# Initialize tkinter
rootWindow = Tk()
rootWindow.attributes('-alpha', 0)

# Request original folder
originalFolder = filedialog.askdirectory(title="Select original document folder:")
if originalFolder == "":
    exit()
# Check whether selected folder is the root folder
# If not, move up a directory
(originalParentFolder, originalFolderName) = os.path.split(originalFolder)
if originalFolderName.lower() == IMPORT_FOLDER_NAME.lower() or originalFolderName.lower() == EXPORT_FOLDER_NAME.lower():
    (originalParentFolder, originalFolderName) = os.path.split(originalParentFolder)

# Request revised folder
revisedFolder = filedialog.askdirectory(title="Select revised document folder:")
if revisedFolder == "":
    exit()
# Check whether selected folder is the root folder
# If not, move up a directory
(revisedParentFolder, revisedFolderName) = os.path.split(revisedFolder)
if revisedFolderName.lower() == IMPORT_FOLDER_NAME.lower() or revisedFolderName.lower() == EXPORT_FOLDER_NAME.lower():
    (revisedParentFolder, revisedFolderName) = os.path.split(revisedParentFolder)

# Define project attributes
projectName = "combo_" + originalFolderName + "-" + revisedFolderName
originalPath = os.path.join(originalParentFolder, originalFolderName)
revisedPath = os.path.join(revisedParentFolder, revisedFolderName)
originalImportPath = os.path.join(originalPath, IMPORT_FOLDER_NAME)
revisedImportPath = os.path.join(revisedPath, IMPORT_FOLDER_NAME)
comboExportPath = os.path.join(revisedPath, EXPORT_FOLDER_NAME)

# Create combo folder if it doesn't exist
if not os.path.exists(comboExportPath):
    os.makedirs(comboExportPath)

# Get filenames from original directory
originalFiles = os.listdir(originalImportPath)

# Get filenames from revised directory
revisedFiles = os.listdir(revisedImportPath)

# Define file lists
originalFilenames = {}
revisedFilenames = {}
originalDuplicateNames = []
revisedDuplicateNames = []
skipNames = []
failureNames = []
successNames = []

# Check for duplicates
for filename in originalFiles:
    if filename.split("_")[0].lower() not in originalFilenames.keys():
        originalFilenames[filename.split("_")[0].lower()] = filename
    else:
        if filename.split("_")[0].lower() not in originalDuplicateNames:
            originalDuplicateNames.append(filename.split("_")[0].lower())
for filename in revisedFiles:
    if filename.split("_")[0].lower() not in revisedFilenames.keys():
        revisedFilenames[filename.split("_")[0].lower()] = filename
    else:
        if filename.split("_")[0].lower() not in revisedDuplicateNames:
            revisedDuplicateNames.append(filename.split("_")[0].lower())

# Remove duplicates from comparison
for name in originalDuplicateNames:
    if name in originalFilenames.keys():
        originalFilenames.pop(name)
    if name in revisedFilenames.keys():
        revisedFilenames.pop(name)
for name in revisedDuplicateNames:
    if name in originalFilenames.keys():
        originalFilenames.pop(name)
    if name in revisedFilenames.keys():
        revisedFilenames.pop(name)

# Sort by whether files exist in sets
readyNames = list(set(originalFilenames.keys()) & set(revisedFilenames.keys()))
mismatchedNames = list(set(originalFilenames.keys()) ^ set(revisedFilenames.keys()))

# Open an instance of Word
word = win32com.client.gencache.EnsureDispatch("Word.Application")

# Iterate through qualified documents
for name in readyNames:

    # Condense information into dictionary
    paper = {
        "author-name": name,
        "original-filename": originalFilenames[name],
        "original-path": os.path.join(originalImportPath, originalFilenames[name]),
        "revised-filename": revisedFilenames[name],
        "revised-path": os.path.join(revisedImportPath, revisedFilenames[name]),
        "combo-filename": name + "_" + projectName + os.path.splitext(originalFilenames[name])[1],
        "combo-path": os.path.join(comboExportPath, name + "_" + projectName + os.path.splitext(originalFilenames[name])[1])
    }

    # Combo already exists, skip it
    if os.path.exists(paper["combo-path"]):
        skipNames.append(name)
        continue

    # Attempt to compare documents in Word
    tries = 0
    success = False
    while tries < 3:
        try:
            word.CompareDocuments(
                word.Documents.Open(os.path.normpath(paper["original-path"])),
                word.Documents.Open(os.path.normpath(paper["revised-path"]))
            )
            word.ActiveDocument.ActiveWindow.View.Type = 3
            word.ActiveDocument.SaveAs(FileName=paper["combo-path"])
        except:
            tries += 1
            continue
        successNames.append(name)
        success = True
        break

# Close instance of Word
word.Quit()

# Record failure
if not success:
    failureNames.append(name)

# Generate message
message = ""
if len(successNames) > 0:
    message += str(len(successNames)) + " document" + ("" if len(successNames) == 1 else "s") + " successfully compared."
if len(skipNames) > 0:
    if message != "":
        message += "\n\n"
    message += str(len(skipNames)) + " existing comparison" + ("" if len(skipNames) == 1 else "s") + " skipped: " + ", ".join(skipNames)
if len(failureNames) > 0:
    if message != "":
        message += "\n\n"
    message += str(len(failureNames)) + " comparison" + ("" if len(failureNames) == 1 else "s") + " failed: " + ", ".join(failureNames)
if len(originalDuplicateNames) > 0:
    if message != "":
        message += "\n\n"
    message += str(len(originalDuplicateNames)) + " duplicate original submission" + ("" if len(originalDuplicateNames) == 1 else "s") + ": " + ", ".join(originalDuplicateNames)
if len(revisedDuplicateNames) > 0:
    if message != "":
        message += "\n\n"
    message += str(len(revisedDuplicateNames)) + " duplicate revised submission" + ("" if len(revisedDuplicateNames) == 1 else "s") + ": " + ", ".join(revisedDuplicateNames)
if len(mismatchedNames) > 0:
    if message != "":
        message += "\n\n"
    message += str(len(mismatchedNames)) + " incomplete set" + ("" if len(mismatchedNames) == 1 else "s") + ": " + ", ".join(mismatchedNames)
if message == "":
    message = "No documents found."

# Create message box
if len(failureNames) > 0:
    messagebox.showerror("Compare", message)
elif len(skipNames) > 0 or len(originalDuplicateNames) > 0 or len(revisedDuplicateNames) > 0 or len(mismatchedNames) > 0:
    messagebox.showwarning("Compare", message)
elif len(successNames) > 0:
    messagebox.showinfo("Compare", message)
else:
    messagebox.showwarning("Compare", message)
