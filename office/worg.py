#encoding:gbk
from win32api import RGB

import win32com.client as handler
from  time import  sleep
from Tkinter import Tk

TEXT='''<p> 
      <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/> 
      <xsl:text>, </xsl:text> 
      <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:First"/> 
      <xsl:text>. (</xsl:text> 
      <xsl:value-of select="b:Year"/> 
      <xsl:text>). </xsl:text> 
      <i> 
         <xsl:value-of select="b:Title"/> 
         <xsl:text>. </xsl:text> 
      </i> 
      <xsl:value-of select="b:City"/> 
      <xsl:text>: </xsl:text> 
      <xsl:value-of select="b:Publisher"/> 
      <xsl:text>.</xsl:text> 
   </p> 
</xsl:template>
 

When you reference a book source in your Word document, Word needs to access this HTML so that it can use the custom style to display the source, so you'll have to add code to your custom style sheet to enable Word to do this.

XML  Copy codeCopy code  
<!--Defines the output of the entire Bibliography-->
 
<xsl:template match="b:Bibliography"> 

   <html xmlns="http://www.w3.org/TR/REC-html40"> 
   
      <body> 

         <xsl:apply-templates select ="b:Source[b:SourceType = 'Book']"> 

         </xsl:apply-templates> 

      </body> 
   
   </html> 
</xsl:template>
 

In a similar fashion, you'll need to do the same thing for the citation output. Follow the pattern (Author, Year) for a single citation in the document.

XML  Copy codeCopy code  
<!--Defines the output of the Citation-->
<xsl:template match="b:Citation/b:Source[b:SourceType = 'Book']"> 
   <html xmlns="http://www.w3.org/TR/REC-html40"> 
      <body> 
         <!-- Defines the output format as (Author, Year)--> 
         <xsl:text>(</xsl:text> 
            <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/> 
         <xsl:text>, </xsl:text> 
         <xsl:value-of select="b:Year"/> 
         <xsl:text>)</xsl:text> 
      </body> 
   </html> 
</xsl:template>
 
'''

"""
http://www.cnblogs.com/xh6300/p/5915717.html
"""
def word():
    word = handler.DispatchEx('Word.Application')
    doc = word.Documents.Add()
    word.Visible=True
    sleep(0.5)
    doc.PageSetup.TopMargin = 570.0
    doc_range = doc.Range(0, 0)
    doc_range.Font.Size = 18
    doc_range.Font.Italic   =1
    doc_range.Font.Color  ='255,88,77'
    doc_range.InsertAfter('Py word generated \r\n\r\n')
    doc_range.InsertAfter('Py word generated \r\n\r\n'+TEXT)
    sleep(1)
    for i in range(1,20):

        doc_range.InsertAfter('this is'+str(i)+'\r\n')
        print(doc.Range(i,100))
        sleep(0.4)
    file = r'C:\Users\ck\Desktop\123.doc'
    doc.SaveAs(file)
    word.Application.Quit()
if __name__=='__main__':
    Tk().withdraw()
    word()
