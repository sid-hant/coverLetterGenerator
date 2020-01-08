#importing libraries
from docx import Document
from docx.shared import Pt
import convertapi

#convert api key
convertapi.api_secret = 'your secret key here'

#Cover letter creator function. first param = template name in docx type
def createCoverLetter(filename, companyName, position):
	doc = Document(filename) #creates a document instance

	#for loop replaces the {COMPANY} with company name and {POSITION} with position name
	for i,p in enumerate(doc.paragraphs):
		if '{COMPANY}' in p.text:
			p.text = p.text.replace('{COMPANY}',companyName)

		if '{POSITION}' in p.text:
			p.text = p.text.replace('{POSITION}',position)

		#styling
		style = doc.styles['Normal']
		font = style.font
		font.name = 'Helvetica Neue'
		font.size = Pt(12)
		p.style = doc.styles['Normal']

	#document name and save
	doc.save('SidhantCoverLetter_'+str(companyName)+'.docx')

	#converts into pdf
	result = convertapi.convert('pdf', { 'File': 'SidhantCoverLetter_'+str(companyName)+'.docx' })
	result.file.save('SidhantCoverLetter_'+str(companyName)+'.pdf')


if __name__ == '__main__':
	info = input('Company Name;Position Name: ') #input your company name and position name seperated by ';' i.e Tesla;Automation Engineer
	info = info.split(';') #splits the input
	createCoverLetter('SidhantCover.docx',info[0],info[1]) #executes the function