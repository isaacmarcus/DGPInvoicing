import PyPDF2

# Function for copying file to an existing folder
def copy_file_to(invToCopy, folderToCopyTo, fileName):
    mergeObj = PyPDF2.PdfFileMerger()
    mergeObj.append(invToCopy)
    mergeObj.write(folderToCopyTo + fileName)
    mergeObj.close()