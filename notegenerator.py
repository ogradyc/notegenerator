#!/usr/bin/python -tt
# Copyright 2018 Cameron O'Grady.
#Licensed under the Apache License, Version 2.0 (the "License");
#you may not use this file except in compliance with the License.
#You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
#Unless required by applicable law or agreed to in writing, software
#distributed under the License is distributed on an "AS IS" BASIS,
#WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#See the License for the specific language governing permissions and
#limitations under the License.

# Using Python curiosity to improve everyday tasks
# http://c9n.org/noteGenerator ** need to create this webpage
# Using PyInstaller to build .exe files - http://stackoverflow.com/questions/32963057/is-there-a-py2exe-version-thats-compatible-with-python-3-5/36508064#36508064
#---2.type and enter: pip install pyinstaller
#---3.Again at Command Prompt type and enter: cd c:\....(the Folder where your file example.py is located)
#---4.Finally type and enter: pyinstaller --onefile noteGenSql.py



import re           # Regular expressions
import datetime     # Date and Time Module
import os
import sys
import io
import pyodbc
import shutil     # http://stackoverflow.com/questions/123198/how-do-i-copy-a-file-in-python
import traceback
import logging    # Found this here: http://stackoverflow.com/questions/4990718/python-about-catching-any-exception
#import encodings

# Should check this out for outlook http://stackoverflow.com/questions/20956424/how-do-i-generate-and-open-an-outlook-email-with-python-but-do-not-send



def actionFiles():
  # Performs the following actions using readFiles and copyFiles
  noteGenDirList = readFiles('noteGenerator')
  noteGenDir = readDir('noteGenerator')

  for filename in noteGenDirList:
    if filename == 'netbuildNumbers.txt':
      # Open and read the file.
      with open('noteGenerator/netbuildNumbers.txt') as f:
        netbuildNumbers = f.readlines()
        f.close()
      # you may also want to remove whitespace characters like `\n` at the end of each line
      # netbuildNumbers = [x.strip() for x in netbuildNumbers] 
      for line in netbuildNumbers:
        pr(line)
        netbuildProjectName = projectName(line)
        if netbuildProjectName is not None:
          pr(netbuildProjectName)

          workDirectory = '../'
          checkDirectory(workDirectory)
          workDirectoryArchive = workDirectory + netbuildProjectName + '/Archive/'
          checkDirectory(workDirectoryArchive)

          filesToCopyDirList = readFiles('templates/filesToCopy/')
          filesToCopyDir = readDir('templates/filesToCopy/')
          for copyFileName in filesToCopyDirList:
            pr('')
            pr(copyFileName)
            newFileName = replacePart(copyFileName, 'projectNumber', netbuildProjectName)
            pr(newFileName)
            pr('')
            newFilePath = workDirectoryArchive + newFileName
            fileToCopy = filesToCopyDir + '/' + copyFileName
            shutil.copy2(fileToCopy, newFilePath)

          noteGenRegEx(netbuildProjectName)

    else:
      pr(filename)
      netbuildProjectName = projectName(filename)
      if netbuildProjectName is not None:
        pr(netbuildProjectName)

        workDirectory = '../'
        checkDirectory(workDirectory)
        workDirectoryArchive = workDirectory + netbuildProjectName + '/Archive/'
        checkDirectory(workDirectoryArchive)

        initiatorFileToMove = noteGenDir + '/' + filename
        initiatorFile = workDirectoryArchive + '/' + filename


        filesToCopyDirList = readFiles('templates/filesToCopy/')
        filesToCopyDir = readDir('templates/filesToCopy/')
        for copyFileName in filesToCopyDirList:
          pr('')
          pr(copyFileName)
          newFileName = replacePart(copyFileName, 'projectNumber', netbuildProjectName)
          pr(newFileName)
          pr('')
          newFilePath = workDirectoryArchive + newFileName
          fileToCopy = filesToCopyDir + '/' + copyFileName
          shutil.copy2(fileToCopy, newFilePath)

        shutil.move(initiatorFileToMove, initiatorFile)
        noteGenRegEx(netbuildProjectName)



def readDir(actionDir):
  # Open and read the files in the notegen directory.
  # http://stackoverflow.com/questions/918154/relative-paths-in-python
  #currentDir = os.path.dirname(os.path.abspath(__file__))
  # Had an issue with the PyInstaller exe keeping the absolute path of the compiled spot, so found a fix below
  # https://github.com/pyinstaller/pyinstaller/issues/1726
  currentDir = os.path.dirname(sys.argv[0])
  noteGenDir = os.path.join(currentDir, actionDir)
  return noteGenDir



def readFiles(actionDir):
  # Open and read the file names in a directory.
  # http://stackoverflow.com/questions/918154/relative-paths-in-python
  noteGenDir = readDir(actionDir)
  noteGenDirList = os.listdir(noteGenDir)
  return noteGenDirList



def checkDirectory(directory):
  if not os.path.exists(directory):
      os.makedirs(directory)


def copyFiles(whatFileToMove, whereToMoveFileTo): #Need to fix this because its not copying, its moving
  shutil.move(whatFileToMove, whereToMoveFileTo)



def replacePart(text, searchTerm, replacementTerm):
  result = ( re.sub(searchTerm,replacementTerm,text)  )
  return result


def projectName(text):
  projectName = re.compile('(N\d\d\d\d\d\d)')
  theProjectName = None
  if projectName.search(text) is not None:
    theProjectName = printResult(projectName, text)
  return theProjectName


def pr(s):
  #I love this work around for the print exception error caused by windows
  #http://stackoverflow.com/questions/5419/python-unicode-and-the-windows-console/32176732#32176732
  try:
    print(s)
  except UnicodeEncodeError:
    for c in s:
      try:
        print( c, end='/')
      except UnicodeEncodeError:
        print( '?', end='')


def dumpclean(d):
  result = ''
  if type(d) == dict:
    od = collections.OrderedDict(sorted(d.items())) # https://stackoverflow.com/questions/9001509/how-can-i-sort-a-dictionary-by-key
    for k, v in od.items(): 
      pr(k, v)
      result = result + str(k, v)
  pr(result)
  return result



def noteGenRegEx(projectName):
# Open and read the file.
  omNotes = open('templates/Template.txt', 'rU')
  omNotes_text = omNotes.read()
  omNotes.close()
  # Could process the file line-by-line, but regex on the whole text at once is even easier.

  if projectName is not None:
    theProjectName = projectName
    data = sqlSearch(theProjectName)
    revisedOMNotes_text = ( re.sub("NetbuildIDHere",theProjectName,omNotes_text)  )
    timeStamp = (currentTime())
    revisedOMNotes_text = ( re.sub('currentTime',timeStamp,revisedOMNotes_text)  )
    
    if data is not None:
      projectTitle = data[0]['Name']
      theprojectTitle = projectTitle
      #print(repr(theprojectTitle))  #http://www.surlyjake.com/blog/2010/10/19/view-string-including-special-characters-like-newlines-in-python/
      revisedOMNotes_text = ( re.sub("TitleHere",theprojectTitle,revisedOMNotes_text)  )

      caseHistory = data[0]['ValueString']
      theCaseHistory = caseHistory.replace("\r\n", "\n")
      theCaseHistory = theCaseHistory.replace("\t", "  ")
      #print(repr(theCaseHistory))  #http://www.surlyjake.com/blog/2010/10/19/view-string-including-special-characters-like-newlines-in-python/
      revisedOMNotes_text = ( re.sub("SOWgoesHere",theCaseHistory,revisedOMNotes_text)  )
      commentData = sqlSearchComments(theProjectName)
      theCommentData = ''
      if commentData is not None:
        assignmentNoteTara = re.compile('Testing assigned to', re.IGNORECASE)
        assignmentNoteLisa = re.compile('will be the tester', re.IGNORECASE)
        array_length = len(commentData) # https://stackoverflow.com/questions/32554527/typeerror-list-indices-must-be-integers-or-slices-not-str
        pr('Number of Comments: ' + str(array_length))
        for i in range(array_length):
          #theCommentData = str(commentData)
          formattedCommentText = commentData[i]['Text']
          formattedCommentText = formattedCommentText.replace("\r\n", "\n")
          formattedCommentText = formattedCommentText.replace("\t", "  ")
          formattedCommentText = formattedCommentText.replace("<p>\n", "")
          formattedCommentText = formattedCommentText.replace("<br>", "")
          theCommentData = commentData[i]['Createdate'].strftime('%Y-%m-%d %H:%M') + ' - ' + commentData[i]['Name'] + ' - ' + formattedCommentText + "\n" + theCommentData
          #pr(theCommentData)
          if assignmentNoteTara.search(commentData[i]['Text']) or assignmentNoteLisa.search(commentData[i]['Text']) is not None:
            aformattedCommentText = commentData[i]['Text']
            aformattedCommentText = aformattedCommentText.replace("<strong>", "")
            aformattedCommentText = aformattedCommentText.replace("</strong>", "")
            aformattedCommentText = aformattedCommentText.replace("<p>", "\n")
            theAssigmentNote = commentData[i]['Createdate'].strftime('%Y-%m-%d %H:%M') + ' - ' + commentData[i]['Name'] + ' - ' + aformattedCommentText
            revisedOMNotes_text = ( re.sub("AssignmentNoteHere",theAssigmentNote,revisedOMNotes_text)  )
        revisedOMNotes_text = ( re.sub("NetbuildCommentsHere",theCommentData,revisedOMNotes_text)  )

    else:
      localTime = datetime.datetime.now()
      timeStamp = localTime.strftime('%Y-%m-%d-%H%M%S')
      newFilePath = '../' + projectName + '/' + timeStamp + ' - sqlErrors.txt'
      shutil.copy2('sqlErrors.txt', newFilePath)

    writeFile(revisedOMNotes_text, theProjectName)

  
  # Place where I found how to replace words in a file
  # http://www.daniweb.com/software-development/python/threads/70426/replace-words-in-a-file#
  
  

def sqlSearch(projectNumber):
    #This is for windows authentication
    #Found my answer here http://stackoverflow.com/questions/16515420/connecting-to-ms-sql-server-with-windows-authentication-using-python
    #Currently, this module only looks for the project description using the netbuild number, with the "N" stripped out
    data = None
    if projectNumber is not None:
      projectNumber = projectNumber.strip( 'N' )
      cnxn = pyodbc.connect(driver='{SQL Server}', server='usidcvsql0037', database='NOE', trusted_connection='yes')
      cursor = cnxn.cursor()
      cursor.execute("select @@VERSION")
      sqlQuery = """
        SELECT DISTINCT
        ---TableA.SequenceId,
        TableA.Name,
        TableC.ValueString

        FROM dbo.ThreeD_Shared_Project TableA 
          JOIN dbo.ThreeD_Shared_ManagedField TableC 
            ON TableA.Id = TableC.ProjectId

        WHERE 

        TableA.SequenceId = (""" + projectNumber + """) 
        AND TableC.DefinitionId = ('bd7b17dd-ebd6-44e3-9e37-49acd3e8210c')
              """
      try:
        cursor = cnxn.cursor()
        cursor.execute(sqlQuery)
        if cursor.rowcount is not 0: # http://stackoverflow.com/questions/16561362/python-how-to-check-if-a-result-set-is-empty
          desc = cursor.description
          #pr(desc)
          column_names = [col[0] for col in desc]
          data = [dict(zip(column_names, row))  
            for row in cursor.fetchall()]

        else:
          with io.open('sqlErrors.txt', "a", encoding="utf-8") as f:
            localTime = datetime.datetime.now()
            timeStamp = localTime.strftime('%Y-%m-%d-%H%M%S')
            output = timeStamp + ' - ' + 'N' + projectNumber + " - No Result for '" + projectNumber + "' Found on server=usidcvsql0037 database=NOE Table=dbo.ThreeD_Shared_Project\n"
            pr(output)
            f.write(output)
            f.close()

      except Exception as e:
        # Found this here: http://stackoverflow.com/questions/4990718/python-about-catching-any-exception
        logging.error(traceback.format_exc())
        # Logs the error appropriately. 


      finally:
        cnxn.close()

      #pr('***this is the dictoionary ***')
      #pr(data[0].keys())
      #pr(data[0]['Name'])
      #pr(data[0]['ValueString'])


    result = data
    return result



def sqlSearchComments(projectNumber):
    #This is for windows authentication
    #Found my answer here http://stackoverflow.com/questions/16515420/connecting-to-ms-sql-server-with-windows-authentication-using-python
    #Currently, this module only looks for the project description using the netbuild number, with the "N" stripped out
    data = None
    if projectNumber is not None:
      projectNumber = projectNumber.strip( 'N' )
      cnxn = pyodbc.connect(driver='{SQL Server}', server='usidcvsql0037', database='NOE', trusted_connection='yes')
      cursor = cnxn.cursor()
      cursor.execute("select @@VERSION")
      sqlQuery = """
        SELECT DISTINCT

        ---TableA.SequenceId,
        ---TableA.Name,

        TableD.Createdate,
        TableE.Name,
        TableD.Text



        FROM dbo.ThreeD_Shared_Project TableA 
          JOIN dbo.ThreeD_Shared_ManagedField TableC 
            ON TableA.Id = TableC.ProjectId

          JOIN dbo.ThreeD_Noe_NetworkOrderRevision TableB 
            ON TableA.Id = TableB.ProjectId

          JOIN dbo.ThreeD_Shared_Comment TableD
            ON TableB.CommentCollectionId = TableD.CollectionId

          JOIN dbo.ThreeD_Shared_User TableE
            ON TableD.SubmitterId = TableE.Id

        WHERE

        TableA.SequenceId = (""" + projectNumber + """) 
              """
      try:
        cursor = cnxn.cursor()
        cursor.execute(sqlQuery)
        if cursor.rowcount is not 0: # http://stackoverflow.com/questions/16561362/python-how-to-check-if-a-result-set-is-empty
          desc = cursor.description
          #pr(desc)
          #data = cursor.fetchall()
          column_names = [col[0] for col in desc]
          #data.update[dict(zip(column_names, row)) # https://stackoverflow.com/questions/6416131/python-add-new-item-to-dictionary
          data = [dict(zip(column_names, row))  
            for row in cursor.fetchall()]

        else:
          with io.open('sqlErrors.txt', "a", encoding="utf-8") as f:
            localTime = datetime.datetime.now()
            timeStamp = localTime.strftime('%Y-%m-%d-%H%M%S')
            output = timeStamp + ' - ' + 'N' + projectNumber + " - No Result for '" + projectNumber + "' Found on server=usidcvsql0037 database=NOE Table=dbo.ThreeD_Shared_Project\n"
            pr(output)
            f.write(output)
            f.close()

      except Exception as e:
        # Found this here: http://stackoverflow.com/questions/4990718/python-about-catching-any-exception
        logging.error(traceback.format_exc())
        # Logs the error appropriately. 


      finally:
        cnxn.close()

      #pr('***this is the dictoionary ***')
      #pr(data[0].keys())
      #pr(data[0]['Name'])
      #pr(data[0]['ValueString'])

    #pr(data)
    result = data
    return result




def currentTime():
 # Module that provides a timestamp for noteGen Creation
  utcTime = datetime.datetime.utcnow()
  localTime = datetime.datetime.now()
  result = utcTime.strftime('UTC %m/%d @ %H:%M - %I:%M%p') + ' - ' + localTime.strftime('%A - Local Time %m/%d @ %I:%M%p')
  return result
  

def printResult(var, text):
  # Module that helps me print my results
  searchResult = var.search(text)
  result = searchResult.group(1)
  #print(result.encode("utf-8"))
  return result


def writeFile(input, projectNumber):
  # Module that writes to a file, then closes it
  projectName = '../' + projectNumber + '/' + projectNumber + '.txt'
  # os.makedirs(os.path.dirname(projectName), exist_ok=True) 
  if os.path.isfile(projectName): # http://stackoverflow.com/questions/82831/how-do-i-check-whether-a-file-exists-using-python
    localTime = datetime.datetime.now()
    timeStamp = localTime.strftime('%Y-%m-%d-%H%M%S')
    projectName = '../' + projectNumber + '/' + timeStamp + ' - ' + projectNumber + '.txt'
    #copyFiles(projectName, oldProjectName) 
    #Here is where I can fix overwriting of the file by simply creating a blank backup.
  with io.open(projectName, "w", encoding="utf-8") as f:
    f.write(input)
    f.close()

def createDefaultFile():
  # Module that overwrites a file, then closes it
  netbuildNumbersFile =  'noteGenerator/netbuildNumbers.txt'
  with io.open(netbuildNumbersFile, "w", encoding="utf-8") as f:
    f.close()


# Define a main() function that calls noteGenRegEx.
def main():  
  actionFiles()
  createDefaultFile()

# This is the standard boilerplate that calls the main() function.
if __name__ == '__main__':
  main()
