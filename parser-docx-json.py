#!/opt/local/bin/python3.4
"""
Tachun Lin
Dept. of Computer Science & Information Systems
Bradley University
djlin@bradley.edu

To convert special DOCX to JSON format
"""
from docx import Document
import re
import sys, getopt

def usage():
   print("%s [-h] -i [input file (DOCX)] -o [output file (JSON)]"% __file__)
   print("Please read NOTES for input docx format")

class Record:
   def __init__(self):
      # personal name (Bold)
      self.pname = "" 
      # family name (bold + italic)
      self.fname = ""
      # English translation
      self.translate = ""
      # language
      self.lang = ""
      # gender
      self.gender = ""
      # written as (orthography)
      self.ortho = []
      # references
      self.ref = []
      # entries (first entry: category, followed by others)
      self.entries = []
      # c.f.
      self.cf = ""

   def printTag(self, tagName, tagValue):
      line = "\"" + tagName + "\": \"" + tagValue + "\"," 
      return line

   def printTagList(self, tagName, tagValueList):
      line = "\"" + tagName + "\": ["
      for i in tagValueList:
         line += "\"" + i + "\", "
      line = line.rstrip(", ")
      line += "],"
      return line

   def write(self, outputFile):
      target = open(outputFile, 'a')
      target.write("{")
      if self.pname != "":
         target.write( self.printTag("person-name", self.pname) )
      if self.fname != "":
         target.write( self.printTag("family-name", self.fname) )
      target.write( self.printTag("translate", self.translate) )
      target.write( self.printTag("language", self.lang) )
      target.write( self.printTag("gender", self.gender) )
      target.write( self.printTagList("orthography", self.ortho) )
      target.write( self.printTagList("refs", self.ref) )
      target.write( self.printTag("cf", self.cf) )
      line =  self.printTagList("entries", self.entries)
      line = line.rstrip(',') 
      target.write( line )
      target.write( "}" )
      target.close()

lineTypes = {0: 'name', 1: 'category', 2: 'entry'}
typeCat = r'^(\s+)(\d+)/(\d+)$'
typeEntry = r'^(\s+)(\d+)\.(\s+)(.+)'
typeFollowingEntry = r'^(\s+)(\w+).+'
typeName = r'^(\w+)'
typeEndEntry = r'^$'
typeRemoveLeadingId = r'(\s+)\d+\.'

def remove_pattern(text, pattern):
   for i in pattern:
      text = text.replace(i, '')
   return text

def parse_name(paragraph, rec):
   line = ""
   # first pass, just to add the subscript and superscript notes
   iter = 0
   for run in paragraph.runs:
      if run.bold and run.italic and iter == 0:
         rec.fname = run.text
         iter=1
      elif run.bold and iter == 0:
         rec.pname = run.text
         iter=1
      elif run.font.superscript:
         line += "<sup>%s</sup>"% run.text
      elif run.font.subscript:
         line += "<sub>%s</sub>"% run.text
      else:
         line += run.text

   # now we are good to parse the content 
   # the name is already stored in the Record object
   contents = re.split(';', line)

   # translate
   rec.translate = remove_pattern(contents[0], "()“”\"").strip()

   if len(contents) > 1:
      info = re.split('\s', contents[1])
      rec.lang = info[1]
      rec.gender = info[2]
      for i in range(4, len(info)):
         rec.ortho.append(info[i])
   # refs
   if len(contents) > 2:
      refs = re.split(',', contents[2])
      for i in refs:
         rec.ref.append(i.strip())
   if len(contents) == 4:
      line = contents[3]
      line = line.lstrip().lstrip("cf.").strip()
      rec.cf = line

def keep_scripts(paragraph):
   line = ""
   for run in paragraph.runs:
      if run.font.superscript:
         line += "<sup>%s</sup>"% run.text
      elif run.font.subscript:
         line += "<sub>%s</sub>"% run.text
      else:
         line += run.text
   return line

def parse_entry(paragraph, record):
   line = keep_scripts(paragraph)
   line = re.split(typeRemoveLeadingId, line)
   line = line[2].lstrip()
   record.entries.append(line)

def parse_following_entry(paragraph, record):
   line = keep_scripts(paragraph).strip()
   record.entries.append(line)

def DocxToJSON(inputFile, outputFile):
   # store all records
   records = []
   currentType = 0
   r = Record()
   records.append(r)
   document = Document(inputFile)
   for paragraph in document.paragraphs:
      rec = records[-1]
      if re.match(typeName, paragraph.text):
         currentType = 0
         parse_name(paragraph, rec)
      elif re.match(typeCat, paragraph.text):
         if currentType == 0:
            pass
         # if coming from antoher category - no entry inside of the category
         elif currentType == 1:
            # close the previous category
            pass
         elif currentType == 2:
            pass
         currentType = 1
         rec.entries.append(paragraph.text.strip())
      elif re.match(typeEntry, paragraph.text):
         # start here, match the new entry info here!
         if currentType == 0:
            # major/direct entry (not inside of a category)
            # stored differently? ask John
            parse_entry(paragraph, rec)
         elif currentType == 1 or currentType == 2:
            # right next to a category
            # new entries added here.
            parse_entry(paragraph, rec)
         currentType = 2
      elif re.match(typeEndEntry, paragraph.text):
         # end of a record
         currentType = 3 
         r = Record()
         records.append(r)
      elif re.match(typeFollowingEntry, paragraph.text):
         if currentType == 2 or currentType == 4:
            # 2: entry following "<space><number>. <content>"
            # 4: entry following the regular entry, in the form "<space><content>"
            parse_following_entry(paragraph, rec)
         currentType = 4
      else:
         currentType = -1

   target = open(outputFile, 'w')
   target.write("{ \"record\": [")
   target.close()
   for i in range(len(records)-1):
      records[i].write(outputFile)
      if i != len(records) - 2:
         target = open(outputFile, 'a')
         target.write(",")
         target.close()
   target = open(outputFile, 'a')
   target.write(" ] }")
   target.close()
   
def main(argv):
   inputFile = ''
   outputFile = 'default_output.json'
   try:
      opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
   except getopt.GetoptError:
      usage()
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         usage()
         sys.exit()
      elif opt in ("-i", "--ifile"):
         inputFile = arg
      elif opt in ("-o", "--ofile"):
         outputFile = arg
   if inputFile != "":
      DocxToJSON(inputFile, outputFile)
      
if __name__ == "__main__":
   main(sys.argv[1:])
