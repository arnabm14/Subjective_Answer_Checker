import tkinter as tk
from tkinter import *
from tkinter import messagebox
#import language_check
import xlrd
from fuzzywuzzy import fuzz
#from nltk.corpus import words

ans=0
text=''
kt=0
cm=0
gm=0
fr=0
g=0
file=""
fileQ=""
Qtext=""
frf=1
ktf=0
cmf=0
gmf=0
def openmyfile(x):
	# print(x)
	global fileQ
	global Qtext
	global frf
	global ktf
	global cmf
	global gmf
	fileQ=x 								#The Scoring factor excel file was proposed to have one values for all questions
	loc = ("Questions/"+fileQ+".xlsx") 		#but it was found that different questions having different parameters provide better accuracy.
	wb = xlrd.open_workbook(loc) 			#the other answers were added on the last day and need refinement. The parameters were determined via a separate code.
	sheet = wb.sheet_by_index(0)
	#print(Qtext)
	Qtext=sheet.cell_value(3, 0)
	# print(Qtext)
	frf=sheet.cell_value(1, 1)
	ktf=sheet.cell_value(1, 2)
	cmf=sheet.cell_value(1, 3)
	gmf=sheet.cell_value(1, 4)
	ans_key(sheet)

strans=""
ansl=[]


keyword=[]
com =["and","that","the","for","it","it's","was","his","who","work","used","way","also","by","can","which","as","known","then","if","between","through","another","","or","my","in","from","a","any","on","combination","to","into","is","of","It","A","each","both"]

def load_words():
	with open('words_alpha.json') as word_file:
		valid_words = set(word_file.read().split())
		#print(valid_words)
	return valid_words
	
def ans_key(s):
	global strans
	strans=""
	global ansl
	ansl.clear()
	global keyword
	keyword.clear()
	for i in range(6,s.nrows): 
		t=s.cell_value(i, 0)
		ansl.append(t)
		strans= strans + " \n\n" +(str)(i-5)+")"+ t
	# print(strans)
	for a in ansl:
		ass= a.split()
		for sas in ass:
			sas=sas.lower()
			for check in ansl:
				if a==check:
					continue
				else:
					assc= check.split()
				for x in assc:
					x=x.lower()
					if x==sas and x not in com and x not in keyword:
						keyword.append(x)
	# keyword=key
	# print (keyword)








	


#for a in ans:
	#a=a.strip()
	#a=a.replace(u'\xa0',u'').encode('utf-8')
#print(ansl)

class Test(tk.Frame):
	def __init__(self):
		new =tk.Frame.__init__(self)
		new = Toplevel(self)
		#new.pack()
		#new.geometry("700x600+361+100")
		new.title("Question Paper")
		#self.resizable(False,False)
		self.tk_setPalette(background='#ececec')
		#sw=self.master.winfo_screenwidth()
		#sh=self.master.winfo_screenheight()
		#w= 361
		#h=223
		#self.master.geometry('%d*%d+%d+%d' % (sw,sh,w,h))
		new.geometry("610x377+361+223")
		
		
		#Dframe= tk.Frame(self)
		#Dframe.pack(padx=5,pady=5)
		#self.title("Automatic Answer checker")
		global Qtext
		tk.Label(new, text=Qtext).grid(row= 0, column=0, padx=13, pady=21,sticky=W)
		
		
		self.entryA = Text(new, height=14,width=46, padx=5, pady=5, wrap=WORD, background='white')
		self.entryA.grid(row= 0, column=1, padx=13, pady=21,sticky=E)
		
		#Bframe= tk.Frame(self)
		#Bframe.pack(padx=20,pady=20, anchor='e')
		tk.Button(new,text='  Submit  ',default='active',command=self.click_ok).grid(row= 1, column=1, padx=13, pady=21)
		
	def click_ok(self):
		#print("Working Successfully")
		global text
		text= self.entryA.get("1.0",END)
		#print(text)
		#texts=text.split()
		if text=="" or text==" " :
			messagebox.showinfo("Blank Input Error","Please enter a Statement")
		elif len(text.split())<5:
			messagebox.showinfo("Too Short Input Error","Please enter a proper Statement")
		#messagebox.showinfo("Evaluation report","Working Successfully")
		else :
			self.newWindow = Report()
			#self.root.withdraw()
			#self.hide()
		#self.withdraw()
		

class App(tk.Frame):
	def __init__(self,master):
		tk.Frame.__init__(self,master)
		self.pack()
		self.master.resizable(False,False)
		self.master.tk_setPalette(background='#ececec')
		#sw=self.master.winfo_screenwidth()
		#sh=self.master.winfo_screenheight()
		#w= 361
		#h=223
		#self.master.geometry('%d*%d+%d+%d' % (sw,sh,w,h))
		self.master.geometry("500x190+500+323")
		
		
		#Dframe= tk.Frame(self)
		#Dframe.pack(padx=5,pady=5)
		self.master.title("Automatic Answer checker")
		
		
		
		# Create a Tkinter variable
		tkt = StringVar(root)

		# Dictionary with options
		choices = { 'E-Commerce','NLP','Cryptography','Cyber-Security','Philosophy'}
		tkt.set('E-Commerce') # set the default option
		openmyfile('E-Commerce')
		popupMenu = OptionMenu(self, tkt, *choices)
		Label(self, text="Subject     : ").grid(row = 2, column = 0, padx=5, pady=10,sticky=W)
		popupMenu.grid(row = 2, column =1)
		global file
		global loc
		global Qtext
		#global sheet
		#global wb
		global frf
		global ktf
		global cmf
		global gmf
		def change_dropdown(*args):
		# on change dropdown value
			#print( tkt.get() )
			file=tkt.get()
			#print (file)
			openmyfile(file)
			# loc = (file+".xlsx") 
			# wb = xlrd.open_workbook(loc) 
			# sheet = wb.sheet_by_index(0) 
			# Qtext=sheet.cell_value(3, 0)
			# frf=sheet.cell_value(1, 1)
			# ktf=sheet.cell_value(1, 2)
			# cmf=sheet.cell_value(1, 3)
			# gmf=sheet.cell_value(1, 4)
			# ans_key(sheet)

		# link function to change dropdown
		tkt.trace('w', change_dropdown)

		
		
		
		
		tk.Label(self, text="Username : ").grid(row= 0, column=0, padx=5, pady=10,sticky=W)
		
		self.entryA = tk.Entry(self,width=26, background='white')
		self.entryA.grid(row= 0, column=1, padx=5, pady=10, sticky=W)
		self.entryA.focus_set()
		tk.Label(self, text="Password : ").grid(row= 1, column=0, padx=5, pady=10,sticky=W)
		
		self.entryB = tk.Entry(self,width=26, background='white', show="*")
		self.entryB.grid(row= 1, column=1, padx=5, pady=15, sticky=W)
		
		
		#Bframe= tk.Frame(self)
		#Bframe.pack(padx=20,pady=20, anchor='e')
		tk.Button(self,text='  Submit  ',default='active',command=self.click_ok).grid(row= 3, column=1, padx=13, pady=21, sticky=S)
	
			
	def click_ok(self):
		#print("Working Successfully")
		#print(text)
		#messagebox.showinfo("Evaluation report","Working Successfully")
		user= self.entryA.get()
		password= self.entryB.get()
		
		op = ("login.xlsx") 
		wbr = xlrd.open_workbook(op) 
		sh = wbr.sheet_by_index(0)
		cuser= sh.row_values(1)
		cpass= sh.row_values(3)
		
		#print(cuser)
		#print(cpass)
		#print(user)
		#print(password)
		
		if user in cuser and password in cpass:
			self.newWindow = Test()
			root.withdraw()
		else:
			messagebox.showinfo("Login error","Please enter the correct credentials")
		#self.withdraw()


		
		
		
		
	
	
class Report(tk.Frame):

		#new.geometry("100x50+665+410")
	def __init__(self):
		new =tk.Frame.__init__(self)
		new = Toplevel(self)
		#new.pack()
		new.geometry("350x180+500+300")
		new.title("Evaluation Report")
		global ans
		ans=0
		global ansl
		global text
		global keyword
		# for ev in ansl:
			# ans = ans + fuzz.token_set_ratio(ev,text)
			# ans = ans + fuzz.ratio(ev,text)
			
		#matches = lang_tool.check(text)
		#lt= len(text)
		#lm=len(matches)
		#lm=lt-lm
		
		
		
		global gm
		global g
		g=0
		text=text.strip()
		english_words = load_words()
		if text==" " or text=="":
			gm=0
		else:
			text= text.split()
			#print(text)
			for t in text:
				t= t.lower()
				if t[-1]=="." : 
					t= t[:-1]
				if t in english_words:
					#print(t)
					g=g+1
				# if g==10:
					# break
		if g>7:
			for ev in ansl:
				ans = ans + fuzz.token_set_ratio(ev,text)
				ans = ans + fuzz.ratio(ev,text)
			# lang_tool = language_check.LanguageTool()
			# matches = lang_tool.check(text)
			# lt= len(text)
			# lm=len(matches)
			# lm=lt-lm
			# gm=lm/lt
			
		
		
		global kt
		global cm
		repeat=[]
		for t in text:
			#if t=="NLP":
			#	lm=lm+1
			t=t.lower()
			if t in keyword and t not in repeat:
				value= keyword.index(t)
				if value>=2:
					kt=kt+0.05
				elif value==0:
					kt=kt+0.1
				elif value==1:
					kt=kt+0.08
				#kt=kt+keywords.get(t)
				repeat.append(t)
		check=[]
		c=0
		#print(repeat)
		#print(keyword)
		for i in range(0,len(repeat)-1):
			if keyword.index(repeat[i])<keyword.index(repeat[i+1]) and repeat[i] not in check:
				check.append(repeat[i])
				check.append(repeat[i+1])
			# a= keyword.index(repeat[i])
			# for t in range(a+1,len(repeat)-1):
				# if a<keyword.index(repeat[t]) and repeat[t] not in check and repeat[a] not in check :
					# check.append(repeat[t])
					# check.append(repeat[a])
		#check.append(len(repeat)-1)
		#print(check)
		#gm=lm/lt
		#km=kt/len(keywords)
		c-len(check)
		if kt > 1:
			kt=1
		if len(check)==0:
			cm=0
		else:
			cm=len(check)/len(repeat)
		#r = (0.2)*gm + (0.6)*kt + (0.2)*cm	
				
			
			
			
	
		global fr
		global x
		# if g>10:
			# gm=1
		# else:
		gm=g/len(text)
		fr= ans/(len(ansl))
		global frf,ktf,cmf,gmf
		# frf=3
		#print(fr)
		#print(len(ansl))
		#x= fr/frf
		#print(x)
		# ktf=0.15
		# cmf=0.1
		# gmf=0.1
		r = fr/(frf*100) + ktf*kt +cmf*cm +gmf*gm
		#print(r)
		
		if r>0.95:
			r=10
		elif r>0.9:
			r=9.5
		elif r>0.85:
			r=9
		elif r>0.8:
			r=8.5
		elif r>0.75:
			r=8
		elif r>0.7:
			r=7.5
		elif r>0.65:
			r=7
		elif r>0.6:
			r=6.5
		elif r>0.55:
			r=6
		elif r>0.5:
			r=5.5
		elif r>0.45:
			r=5
		elif r>0.4:
			r=4.5
		elif r>0.35:
			r=4
		elif r>0.3:
			r=3.5
		elif r>0.25:
			r=3
		elif r>0.2:
			r=2.5
		elif r>0.15:
			r=2
		elif r>0.1:
			r=1.5
		elif r>0.05:
			r=1
		else:
			r=0		
		
		
		ans = (str)(r) 
		#QT= " Your Answer is = " + ans
		new.label = tk.Label(new, text=" Your Marks is = " + ans)
		#new.label.pack(side='left')
		new.label.grid(row= 2, column=2, padx=10, pady=15,sticky=N)
		new.button = tk.Button( new, text = "Detailed Report", width = 15,command = self.Det_report )
		#new.button.pack(side='right')
		new.button.grid(row= 3, column=3, padx=10, pady=15,sticky=N)
		new.button2 = tk.Button( new, text = "Close", width = 15,command = self.close_window )
		#new.button.pack(side='right')
		new.button2.grid(row= 4, column=3, padx=10, pady=15,sticky=N)
	def Det_report(self):
		self.newWindow = Det_Report()
		# self.hide()
	
	def close_window(self):
		self.master.destroy()

class Det_Report(tk.Frame):

		#new.geometry("100x50+665+410")
	def __init__(self):
		new =tk.Frame.__init__(self)
		new = Toplevel(self)
		#new.pack()
		new.geometry("700x650+361+100")
		new.title("Evaluation Full Report")
		global ans
		ans = (str)(ans)
		global x
		# x= (str)(x)
		#QT= " Your Answer is = " + ans
		#new.label = tk.Label(new, text=" Your Total Marks is = " + ans)
		#new.label.pack(side='left')
		#new.label.grid(row= 2, column=2, padx=5, pady=10,sticky=W)
		tk.Message(new, text=" Your Total Marks is = " + ans, font='Arial 10 underline',justify='left',aspect=1500).grid(row= 1, column=2, padx=2, pady=2,sticky=W)
		#new.label2 = tk.Label(new, text=" The Keyword accuracy of the sentence is :        {:.2%}".format(kt))
		#new.label.pack(side='left')
		#new.label2.grid(row= 3, column=2, padx=5, pady=10,sticky=W)
		tk.Message(new, text=" The Similarity factor of the sentence is :            {:.2%}".format(fr/100), font='Arial 10 underline',justify='left',aspect=1500).grid(row= 2, column=2, padx=2, pady=2,sticky=W)
		tk.Message(new, text=" The Grammar accuracy of the sentence is :              {:.2%}".format(gm), font='Arial 10 underline',justify='left',aspect=1500).grid(row= 3, column=2, padx=2, pady=2,sticky=W)
		tk.Message(new, text=" The Total Keywords found in the sentence is :          {:.2%}".format(kt), font='Arial 10 underline',justify='left',aspect=10000).grid(row= 4, column=2, padx=2, pady=2,sticky=W)
		
		
		tk.Message(new, text=" The Keyword order accuracy of the sentence is :        {:.2%}".format(cm), font='Arial 10 underline',justify='left',aspect=10000).grid(row= 5, column=2, padx=2, pady=2,sticky=W)
		#new.label.pack(side='left')
		#new.label3.grid(row= 4, column=2, padx=5, pady=10,sticky=W)
		key="> "
		for k in keyword:
			if k==keyword[len(keyword)-1]:
				key= key+ " , " + k + " . "
			elif k==keyword[0]:
				key= key + k
			else:
				key= key+ " , " + k
		#key=key.reverse()
		tk.Message(new, text=" Some of the Sample Answers are: "+strans, font='System 14',justify='left',aspect=200).grid(row= 6, column=2, padx=2, pady=2,sticky=W)
		tk.Message(new, text=" The Keywords Extracted are :- \n"+key, font='System 12',justify='left',aspect=900).grid(row= 7, column=2, padx=2, pady=2,sticky=W)
		
		new.button2 = tk.Button( new, text = "Close", width = 15,command = self.close_window )
		#new.button.pack(side='right')
		new.button2.grid(row= 8, column=2, padx=100, pady=15,sticky=S)
		
	def close_window(self):
		self.master.destroy()
	

	
if __name__=='__main__':
	root= Tk()
	app = App(root)
	app.mainloop()