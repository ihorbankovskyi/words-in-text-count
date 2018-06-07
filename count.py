import textract, glob, re, sys
import docx2txt

result = {}
present_keywords = set()

COUNT_LIMIT = 2
try:
    COUNT_LIMIT = int(raw_input('Hello\nWanna change word count limit ? (default is 2):'))
except:
    pass
print 'Word count limit = {}'.format(COUNT_LIMIT)

#read keywords
with open('words.csv') as f:
    keywords = [w.lower().strip() for w in f.read().split('\n') if w]

print 'Total {} keywords in words.csv'.format(len(keywords))
 
for i,fname in enumerate(glob.glob('text/*')):
    #try:
    sys.stdout.write('{}/{}\r'.format(i,len(glob.glob('text/*'))))
    text = textract.process(fname).lower()
    wordcount = {}
    for keyword in keywords:
        count = sum(1 for _ in re.finditer(r'\b%s\b' % re.escape(keyword), text))
        if count > COUNT_LIMIT:
            wordcount[keyword] = count
            if keyword not in present_keywords:
                present_keywords.add(keyword)
    result[fname] = wordcount
    #except Exception as e:
    #    print e, fname
        
#write result
import xlsxwriter

wb = xlsxwriter.Workbook('result_occurences.xlsx')
ws = wb.add_worksheet()

#write header
ws.write(0,0, "File Name")
max_tags = len(max(result.values(), key=lambda wc:len(wc)))
for i in range(max_tags):
    ws.write(0,1 + 1*i, "Tag %s" % (i+1))

row = 1
for fpath, wordcount in result.iteritems():
    ws.write(row,0,fpath[5:])
    i = 0
    for word, count in wordcount.iteritems(): 
        ws.write(row,1 + 1*i,word)
        i += 1
    row += 1
wb.close()

wb = xlsxwriter.Workbook('result_count.xlsx')
ws = wb.add_worksheet()

#write header
ws.write(0,0, "File Name")
max_tags = len(max(result.values(), key=lambda wc:len(wc)))
for i in range(max_tags):
    ws.write(0,1 + 2*i + 1, "Num %s" % (i+1))
    ws.write(0,1 + 2*i, "Tag %s" % (i+1))

row = 1
for fpath, wordcount in result.iteritems():
    ws.write(row,0,fpath[5:])
    i = 0
    for word, count in wordcount.iteritems(): 
        ws.write(row,1 + 2*i + 1,count)
        ws.write(row,1 + 2*i,word)
        i += 1
    row += 1
wb.close()

with open('present_keywords.csv', 'w') as f:
    for keyword in present_keywords:
        f.write(keyword)
        f.write('\n')

