from tkinter import *
from tkinter import messagebox
from tkinter import simpledialog
import dropbox
import os
import shutil
import PyPDF2
import getpass
import xlrd

#Event Functions
def getSearchResults(data, start, resultList):
	if data.searchEntry == "":
		if start == 0:
			results = data.dbx.files_list_folder("/Doc_Library")
			for file in results.entries:
				if file.name.endswith(".pdf"):
					resultList.append(file)
			if results.has_more:
				return resultList + getSearchResults(data, results.cursor, resultList)
		else:
			results = data.dbx.files_list_folder_continue(start)
			for file in results.entries:
				if file.name.endswith(".pdf"):
					resultList.append(file)
			if results.has_more:
				return resultList + getSearchResults(data, results.cursor, resultList)
	else:
		results = data.dbx.files_search("/Doc_Library", data.searchEntry, start)
		for file in results.matches:
			if file.metadata.name.endswith(".pdf"):
				resultList.append(file.metadata)
		if results.more:
			return resultList + getSearchResults(data, results.start, resultList)
	return resultList

def setSearchResults(data):
	data.searchresults.delete(0, END)
	x = 0
	for result in data.searchResultList:
		data.searchresults.insert(x, result.name)
		x += 1

def updateSearch(data):
	data.searchEntry = data.searchBar.get()
	data.searchBar.delete(0, END)
	data.searchResultList = getSearchResults(data, 0, [])
	setSearchResults(data)

def setPDFs(data):
	data.pdfs.delete(0, END)
	x = 0
	for result in data.pdfList:
		data.pdfs.insert(x, result.name)
		x += 1

def updatePDFs(data):
	data.selectedResults = []
	selectedIndexes = data.searchresults.curselection()
	for index in selectedIndexes:
		data.selectedResults.append(data.searchResultList[index])
	data.pdfList.extend(data.selectedResults)
	setPDFs(data)

def clearPDFs(data):
	if data.pdfs.curselection() == ():
		data.pdfList = []
	else:
		selectedIndexes = data.pdfs.curselection()
		selectedIndexes = selectedIndexes[::-1]
		for index in selectedIndexes:
			del data.pdfList[index]
	setPDFs(data)

def reset(data):
	data.searchResultList = []
	data.pdfList = []
	data.selectedResults = []
	data.selectedPDFs = []

	setSearchResults(data)
	setPDFs(data)

def createPDF(data):
	filePaths = []

	for file in data.pdfList:
		filePath = file.path_lower
		filePaths.append(filePath)

	if ("Selected Files" in os.listdir()):
		shutil.rmtree("Selected Files")
	os.makedirs("Selected Files")

	filePathsInComp = []

	for filePath in filePaths:
		data.dbx.files_download_to_file("Selected Files/" + filePath.split("/")[-1], filePath)
		filePathsInComp.append("Selected Files/" + filePath.split("/")[-1])
		
	pdfMerger = PyPDF2.PdfFileMerger()

	for filePath in filePathsInComp:
		pdfMerger.append(filePath)

	name = simpledialog.askstring("New PDF", "What is the filename?", parent = data.root)

	if (name + ".pdf" in os.listdir(data.destinationPath)):
		os.remove(data.destinationPath + "\\" + name + ".pdf")
		
	pdfMerger.write(name + ".pdf")

	pdfMerger.close()

	shutil.move(name + ".pdf", data.destinationPath)

	messagebox.showinfo("Process", "Done!")

	reset(data)

def importExcel(data):
	excelPath = simpledialog.askstring("Excel File", "What's the path?\n(To get the path, right click the excel file and click on properties,\nthen highlight the text after 'Location:' and copy paste that into here and then add a '\\' and the file name afterward (with the .xls file type)", parent = data.root)
	if excelPath != None:
		wb = xlrd.open_workbook(filename = excelPath)
		sheetName = simpledialog.askstring("Excel Sheet", "What's the sheet name?\n(The sheets are the tabs near the bottom of an excel workbook)", parent = data.root)
		if sheetName != None:
			ws = wb.sheet_by_name(sheetName)
			cellRange = simpledialog.askstring("Cell Range", "What is the range of cells to select?\n(In the format: Column:First Row:Last Row, i.e. A:1:10)", parent = data.root)
			cellRanges = cellRange.split(":")
			nameList = ws.col_values(ord(cellRanges[0])-65, int(cellRanges[1])-1, int(cellRanges[2]))
			print(nameList)
			for fileName in nameList:
				file = data.dbx.files_search("/Doc_Library", fileName + ".pdf").matches[0].metadata
				data.pdfList.append(file)

			setPDFs(data)

def changeDestination(data):
	newPath = simpledialog.askstring("File Destination", "The current destination is " + data.destinationPath + ".\n What do you want the new file destination to be?\n(To get the path, right click the desired folder and click on properties,\nthen highlight the text after 'Location:' and copy past that into here and then add a '\\' and the folder name afterward", parent = data.root)
	data.destinationPath = newPath

def getHelp(data):
	messagebox.showinfo("Help Instructions", """Near the top are general instructions to help guide your way through the process.

Below the instruction box is the search bar. Here you can search for the pdfs that you want to merge (press enter to search).

You can either search for the full part name or a section of the name, and in the search results box below, all files in the Document Library with your search query somewhere in the name will appear.

In the search results box, you can click to select and deselect the files you want, and then click the '+' button near the top of the box to add the selected files to the merging pdfs box.

In the merging pdfs box, you can also select and deselect files and then press the red 'clear pdfs' button near the bottom to remove from your merging pdfs, or you can click the red 'clear pdfs' button to completely clear the merging pdfs.

When you're done, click the 'merge pdfs' button to merge all of the pdfs in the merging pdfs box.

When merging, the program will ask you what to name the file, and once finished merging, will place the file in the desktop by default.

If you want to change where the program places the resulting merged pdf, you can click the 'Change Destination' button on the top, and enter the file path of the new destination.

If you want to link an excel, click the 'Import Excel' button on the top, and enter the file path of the excel file.

From there, enter the excel cells which contain the part names that you want (follow listed instructions), and then those parts will be put into the merging pdfs box.

For a video walkthrough: https://youtu.be/snCPksdYlpI""")

#Init Functions
def init_all(data):
	def init_data(data):
		data.dbx = dropbox.Dropbox("GZoeG7azR-AAAAAAAAAAFrxyIJdb2l0BcBT-2u4LLDBv8llCkZPbnzhj6J3wBOZV")
		data.searchEntry = ""
		data.searchResultList = []
		data.pdfList = []
		data.selectedResults = []
		data.selectedPDFs = []
		data.destinationPath = "C:\\Users\\" + getpass.getuser() + "\\Desktop"

	def init_frames(data):
		data.bgFrame = Frame(data.root, bg = "#073416")
		data.bgFrame.place(width = data.width, height = data.height, x = 0, y = 0)

		data.instructionFrame = Frame(data.root)
		data.instructionFrame.place(width = 0.95*data.width, height = 0.12*data.height, x = 0.025*data.width, y = 0.025*data.height)

		data.searchLabelFrame = Frame(data.root)
		data.searchLabelFrame.place(width = 0.125*data.width, height = 0.05*data.height, x = 0.025*data.width, y = 0.16*data.height)
		data.searchBarFrame = Frame(data.root)
		data.searchBarFrame.place(width = 0.825*data.width, height = 0.05*data.height, x = 0.15*data.width, y = 0.16*data.height)

		data.searchResultLabelFrame = Frame(data.root)
		data.searchResultLabelFrame.place(width = 0.375*data.width, height = 0.05*data.height, x = 0.025*data.width, y = 0.225*data.height)
		data.addButtonFrame = Frame(data.root)
		data.addButtonFrame.place(width = 0.075*data.width, height = 0.05*data.height, x = 0.4*data.width, y = 0.225*data.height)
		data.searchResultFrame = Frame(data.root)
		data.searchResultFrame.place(width = 0.45*data.width, height = 0.675*data.height, x = 0.025*data.width, y = 0.3*data.height)

		data.pdfLabelFrame = Frame(data.root)
		data.pdfLabelFrame.place(width = 0.45*data.width, height = 0.05*data.height, x = 0.525*data.width, y = 0.225*data.height)
		data.pdfFrame = Frame(data.root)
		data.pdfFrame.place(width = 0.45*data.width, height = 0.575*data.height, x = 0.525*data.width, y = 0.3*data.height)

		data.trashcanFrame = Frame(data.root)
		data.trashcanFrame.place(width = 0.175*data.width, height = 0.075*data.height, x = 0.525*data.width, y = 0.9*data.height)
		data.createPDFFrame = Frame(data.root)
		data.createPDFFrame.place(width = 0.175*data.width, height = 0.075*data.height, x = 0.8*data.width, y = 0.9*data.height)

	def init_labels(data):
		def getInstructionText():
			return "To merge multiple PDFs, type in the part number into the ‘Search Bar,’"\
			" then select the part from the ‘Search Result’ box. After selecting one or multiple "\
			"parts click the ‘Add +’ button to the right of the ‘Search Result’ title box. Repeat this "\
			"process till all desired PDF names are displayed in ‘Merging PDFs’ box. When the new PDF is ready"\
			" to be created click the green ‘Create PDF’ button. When ready to create a different PDF click the "\
			"‘Clear PDFs’ button and start the process over. 'The Merge PDFs' button will merge your pdfs in a file on the desktop."

		data.instructions = Label(data.instructionFrame, text = getInstructionText(), font = ("Comic Sans MS", 8, "bold"), relief = RAISED, bd = 5, wraplength = 750, justify = LEFT, bg = "#a4b2a8")
		data.instructions.pack(expand = True, fill = BOTH)

		data.searchLabel = Label(data.searchLabelFrame, text = "⌕", font = ("Comic Sans MS", 20,), relief = RAISED, bd = 5, pady = 10, bg = "#466d52")
		data.searchLabel.pack(expand = True, fill = BOTH)

		data.searchResultLabel = Label(data.searchResultLabelFrame, text = "Search Results:", font = ("Comic Sans MS", 16,), relief = RAISED, bd = 5, bg = "#a4b2a8")
		data.searchResultLabel.pack(expand = True, fill = BOTH)

		data.pdfLabel = Label(data.pdfLabelFrame, text = "Merging PDFs:", font = ("Comic Sans MS", 16,), relief = RAISED, bd = 5, bg = "#a4b2a8")
		data.pdfLabel.pack(expand = True, fill = BOTH)

	def init_buttons(data):
		data.addButton = Button(data.addButtonFrame, text = "+", font = ("Comic Sans MS", 16,), bg = "#466d52", activebackground = "#355940", bd = 5, command = lambda: updatePDFs(data))
		data.addButton.pack(expand = True, fill = BOTH)

		data.trashcanButton = Button(data.trashcanFrame, text = "Clear PDFs", font = ("Comic Sans MS", 16,), bg = "#f23932", activebackground = "#c13530", bd = 5, command = lambda: clearPDFs(data))
		data.trashcanButton.pack(expand = True, fill = BOTH)
		data.createPDFButton = Button(data.createPDFFrame, text = "Merge PDFs", font = ("Comic Sans MS", 16,), bg = "#33f439", activebackground = "#2abf2f", bd = 5, command = lambda: createPDF(data))
		data.createPDFButton.pack(expand = True, fill = BOTH)

	def init_entries(data):
		data.searchBar = Entry(data.searchBarFrame, bd = 5, relief = RAISED, textvariable = data.searchEntry, font = ("Comic Sans MS", 16,), bg = "#bac6be")
		data.searchBar.pack(expand = True, fill = BOTH)

	def init_scrollbars(data):
		data.searchScroll = Scrollbar(data.searchResultFrame, bg = "#a4b2a8", activebackground = "#7caa7e", bd = 3)
		data.searchScroll.pack(side = RIGHT, fill = Y)

		data.pdfScroll = Scrollbar(data.pdfFrame, bg = "#a4b2a8", activebackground = "#7caa7e", bd = 3)
		data.pdfScroll.pack(side = RIGHT, fill = Y)

	def init_listboxes(data):
		data.searchresults = Listbox(data.searchResultFrame, cursor = "target", font = ("Comic Sans MS", 16,), selectmode = MULTIPLE, bd = 5, yscrollcommand = data.searchScroll.set, bg = "#bac6be")
		data.searchScroll.config(command = data.searchresults.yview)
		data.searchresults.pack(expand = True, fill = BOTH)

		data.pdfs = Listbox(data.pdfFrame, cursor = "target", font = ("Comic Sans MS", 16,), selectmode = MULTIPLE, bd = 5, yscrollcommand = data.pdfScroll.set, bg = "#bac6be")
		data.pdfScroll.config(command = data.pdfs.yview)
		data.pdfs.pack(expand = True, fill = BOTH)

	def init_menus(data):
		data.menuBar = Menu(data.root)
		data.menuBar.add_command(label = "Change Destination", command = lambda: changeDestination(data))
		data.menuBar.add_command(label = "Import Excel", command = lambda: importExcel(data))
		data.menuBar.add_command(label = "Help", command = lambda: getHelp(data))
		data.root.config(menu = data.menuBar)

	init_data(data)
	init_frames(data)
	init_labels(data)
	init_buttons(data)
	init_entries(data)
	init_scrollbars(data)
	init_listboxes(data)
	init_menus(data)

def main(width = 800, height = 800):
	class Struct(object): pass
	data = Struct()
	data.width = width
	data.height = height

	data.root = Tk()
	data.root.title("OZ Document Merger")
	data.root.geometry(str(data.width) + "x" + str(data.height))

	data.root.bind("<Return>", lambda event:
							updateSearch(data))

	init_all(data)

	data.root.mainloop()

main()