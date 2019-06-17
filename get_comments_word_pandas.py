from lxml import etree
import zipfile
import pandas as pd
import easygui

ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'w14':'http://schemas.microsoft.com/office/word/2010/wordml',
          'w15':'http://schemas.microsoft.com/office/word/2012/wordml'}

def get_comments(docxFileName):
  df = pd.DataFrame(columns=['id','author','date','comment'])
  docxZip = zipfile.ZipFile(docxFileName)
  commentsXML = docxZip.read('word/comments.xml')
  et = etree.XML(commentsXML)
  comments = et.xpath('//w:comment',namespaces=ooXMLns)
  for i, c in enumerate(comments):
    id_ = c.xpath('.//w:p',namespaces=ooXMLns)
    id_ = id_[len(id_)-1].xpath("@w14:paraId",namespaces=ooXMLns)[0]
    author = c.xpath('@w:author',namespaces=ooXMLns)[0]
    date = c.xpath('@w:date',namespaces=ooXMLns)[0]
    # comment:
    comment = c.xpath('string(.)',namespaces=ooXMLns)
    df.loc[len(df)] = {
        'id': id_,
        'author': author,
        'date': date,
        'comment': comment
    }
  return df

def get_comment_resolved(docxFileName):
  df = pd.DataFrame(columns=['id','done','parent','reply'])
  docxZip = zipfile.ZipFile(docxFileName)
  commentsXML = docxZip.read('word/commentsExtended.xml')
  et = etree.XML(commentsXML)
  comments = et.xpath('//w15:commentEx',namespaces=ooXMLns)
  for c in comments:
    id_ = c.xpath('@w15:paraId',namespaces=ooXMLns)[0]
    done = c.xpath('@w15:done',namespaces=ooXMLns)[0]
    parent = c.xpath('@w15:paraIdParent',namespaces=ooXMLns)
    parent = parent[0] if len(parent) > 0 else ''
    df.loc[len(df)] = {
        'id': id_,
        'done': done,
        'parent': parent,
        'reply': 'yes' if parent != '' else ''
    }
  return df

docfilepath = easygui.fileopenbox("Select docx file to extract")
outfilepath = easygui.filesavebox("Specify output excel location")

# docfilepath = input("Please specify path to docx file \n(e.g.: draft_report.docx or C:\\results\\report.docx):\n")
# outfilepath = input("\n\nPlease specify path to output excel \n(e.g.: comments.xlsx or C:\\comments\\c.xlsx)\n")

comments = get_comments(docfilepath)
solved = get_comment_resolved(docfilepath)
comments.merge(solved, how='left')[['reply','author','date','comment','done']].to_excel(outfilepath)
