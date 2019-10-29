#! python3
#Will make a set of folders for all the books and chapters of the Old and New Testament

#You will have to install openpyxl to run this
import os, openpyxl

#This function creates the .xlsx file for regular reading and summary
def chapter_summary(chapters):
    x = int(chapters)
    wb = openpyxl.Workbook()
    sheet = wb.get_active_sheet()
    
    sheet.title = 'Chapter Summaries'
    sheet['A1'] = "Chapter"
    sheet['B1'] = "Summary"
    
    #had to bump the range by 1 to get all chapters
    for y in range(1,x+1):
        sheet['A' + str(y+1)] = y
    	
    wb.save('chapter_summary.xlsx')


#This is the function that creates the book and chapter folders	
def build_book(title, chapters):
    starting_point = os.getcwd()
    os.makedirs(title)
    os.chdir(title)
    
    x = int(chapters)
    chapter_summary(x)
    
    #had to bump the range by 1 to get all chapters
    for y in range(1,x+1):
        chapter = "Chapter_" + str(y)
        os.makedirs(chapter)
     
    os.chdir(starting_point)
	
#Dictionaries of the Old and New Testament the books are the keys 
#and the chapters are the values
old_test = {
    '01-Genesis': 50,
	'02-Exodus': 40,
	'03-Leviticus': 27,
	'04-Numbers': 36,
	'05-Deuteronomy': 34,
	'06-Joshua': 24,
	'07-Judges': 21,
	'08-Ruth': 4,
	'09-1 Samuel': 31,
	'10-2 Samuel': 24,
	'11-1 Kings': 22,
	'12-2 Kings': 25,
	'13-1 Chronicles': 29,
	'14-2 Chronicles': 36,
	'15-Ezra': 10,
	'16-Nehemiah': 13,
	'17-Esther' : 10,
	'18-Job' : 42,
	'19-Psalms' : 150,
	'20-Proverbs' : 31,
	'21-Ecclesiastes' : 12,
	'22-Song of Solomon' : 8,
	'23-Isaiah' : 66,
	'24-Jeremiah' : 52,
	'25-Lamentations' : 5,
	'26-Ezekiel' : 48,
	'27-Daniel' : 12,
	'28-Hosea' : 14,
	'29-Joel' : 3,
	'30-Amos' : 9,
	'31-Obadiah' : 1,
	'32-Jonah' : 4,
	'33-Micah' : 7,
	'34-Nahum' : 3,
	'35-Habakkuk' : 3,
	'36-Zephanianh' : 3,
	'37-Haggai' : 2,
	'38-Zechariah' : 14,
	'39-Malachi' : 4
	}

new_test = {
    '01-Matthew' : 28,
	'02-Mark' : 16,
	'03-Luke' : 24,
	'04-John' : 21,
	'05-Acts' : 28,
	'06-Romans' : 16,
	'07-1 Corinthians' : 16,
	'08-2 Corinthians' : 13,
	'09-Galatians' : 6,
	'10-Ephesians' : 6,
	'11-Philippians' : 4,
	'12-Colossians' : 4,
	'13-1 Thessalonians' : 5,
	'14-2 Thessalonians' :3,
	'15-1 Timothy' : 6,
	'16-2 Timothy' :4,
	'17-Titus' : 3,
	'18-Philemon' : 1,
	'19-Hebrews' : 13,
	'20-James' : 5,
	'21-1 Peter' : 5,
	'22-2 Peter' : 3,
	'23-1 John' : 5,
	'24-2 John' : 1,
	'25-3 John' : 1,
	'26-Jude' : 1,
	'27-Revelation' : 22
	}

#Calling the Main Directory to be able to return home	
main_dir = os.getcwd()

#Making the Old testament
os.makedirs('01-Old Testament')
os.chdir('01-Old Testament')
for k in old_test.keys():
	v = old_test[k]
	build_book(k,v)

#Returning to the main directory
os.chdir(main_dir)
#making the new testament
os.makedirs('02-New Testament')
os.chdir('02-New Testament')

for k in new_test.keys():
	v = new_test[k]
	build_book(k,v)
