## Dependencies
# From the pywin32 library (http://sourceforge.net/projects/pywin32/)
import win32process, win32security, win32api, win32gui, win32con

## Standard stuff
import Tkinter
import re
import time
import datetime
from ctypes import Structure, windll, c_uint, sizeof, byref

## Globals
IDLE_TIMEOUT = 10 # Idle duration in seconds before 'Idle' is logged instead of the active window
VERSION = "1.03"

class soglog:
	def __init__(self, root):
		self.recording = False
		self.root = root
		self.log = {}
		self.totalticks = 0
		self.totalTime = 0
		self.startTime = None
		self.timer = RepeatTimer(self)
		self.started = False
		
		# initialize buttons
		buttonWidth = 16

		self.frame = Tkinter.Frame(self.root, padx = 24, pady = 12)
		self.frame.grid()
		
		self.startButton = Tkinter.Button(self.frame, width=buttonWidth,\
				text = "Start", command = self.start, padx=8,pady=6)
		self.startButton.grid(row=1,column=1)
		
		self.resetButton = Tkinter.Button(self.frame, width=buttonWidth,\
				text = "Reset", command = self.reset, padx=8,pady=6)
		self.resetButton.grid(row=1,column=2)
				
		self.stopButton = Tkinter.Button(self.frame, width=buttonWidth,\
				text = "Stop", command = self.stop, padx=8,pady=6, state=Tkinter.DISABLED)
		self.stopButton.grid(row=2,column=1)
		
		self.exitButton = Tkinter.Button(self.frame, width=buttonWidth,\
				text = "Exit", command = self.exit, padx=8,pady=6)
		self.exitButton.grid(row=2,column=2)
		
		yr = datetime.date.today().year
		Tkinter.Label(self.frame, height=1, text ='By Jeroen Kools, 2012-%s. GPL v3' % yr).grid(row=5,column=1,padx=10)
		
		# initialize canvas
		self.pieCanvas = Tkinter.Canvas(self.frame, width=280, height=520)
		self.pieCanvas.grid(row=4,column=1, columnspan=3, padx=4, pady=2)
				
		xy = 20, 20, 270, 270
		self.pieCanvas.create_oval(xy, tags="foobar")
		self.colors = ['slateblue4', 'mediumblue', 'deepskyblue3', 'aquamarine4', 'darkolivegreen3',\
		 'greenyellow', 'gold3', 'chocolate3','red4', 'violetred4', 'darkorchid4']
		self.idleColor = 'grey30'
		self.otherColor =  'grey70'
		
		# Set up escape = quit shortkey
		self.root.bind("<Escape>", self.exit)
				
		
	def start(self):
		self.startTime = time.time()
		if not self.started:			
			self.timer.start()
			self.started = True
		self.stopButton.configure(state=Tkinter.ACTIVE)
		self.startButton.configure(text="Minimize", command = self.minimize)
		
	def reset(self):
		self.log = {}
		if not self.started:
			self.pieCanvas.delete(Tkinter.ALL)
		else:
			self.startTime = time.time()
		self.totalticks = 0
		self.totalTime = 0
	
	def stop(self):
		self.timer.stop()
		self.totalTime += time.time() - self.startTime
		self.started = False
		self.stopButton.configure(state=Tkinter.DISABLED)
		self.startButton.configure(text="Start", command = self.start)
			
	def exit(self, event = None):
		self.stop()
		self.root.after_idle(self.root.quit)
		
	# Update the pie chart and statistics
	def updatePie(self):
		self.pieCanvas.delete(Tkinter.ALL)
		sortedprograms = sorted(self.log, key = self.log.get, reverse = True )
		
		xy = 20, 20, 270, 270 # Pie bounding box
		b = 277				  # starting height of text
		self.pieCanvas.create_oval(xy, fill=self.otherColor)
		at = 0				  # current pie position in degrees
		toptensize = 0		  # pie angle (in degrees) covered by top 10 programs
		
		# Draw pie slices and legend for top ten programs
		for i in range(min(10,len(self.log))):
			program = sortedprograms[i]
			programshort = program
			if len(program)>=30:
				programshort = program[:27]+'...'
				
			size = float(self.log[program]) / self.totalticks * 360.0
			if size < 0.3: # very small pie slices are sometimes drawn as full circles (Tkinter bug?), avoid this
				size = 0 
			elif size == 360: # full circles are drawn as 0 degree slices
				size = 359.99999
				
			color = self.colors[i]
			if program == "Idle":
				color = self.idleColor

			self.pieCanvas.create_arc(xy, start=at, extent=size, fill=color)
			at += size
			row = 20*i			
			text = '%i. %s' % (i+1, programshort)
			percent = '%.2f%%' % (size/3.6)
			toptensize += size/360
			
			self.pieCanvas.create_text(25, b+row, text= text, anchor=Tkinter.NW, tag='legend')
			self.pieCanvas.create_text(220, b+row, text= percent, anchor=Tkinter.NW, tag='legend')
			self.pieCanvas.create_rectangle(10, b+2+row, 20, b+12+row, fill=color, tag='legend')
		
		# Draw 'Other' pie slice if needed
		other = 1 - toptensize
		if other > .01:
			i += 1
			row = 20*i
			percent = '%.2f%%' % (100.0*other)
			self.pieCanvas.create_text(25, b+row, text= "Other/None", anchor=Tkinter.NW, tag = 'legend')					
			self.pieCanvas.create_text(220, b+row, text= percent, anchor=Tkinter.NW, tag='legend')
			self.pieCanvas.create_rectangle(10, b+2+row, 20, b+12+row, fill=self.otherColor, tag = 'legend')
		
		# Draw total time logged line
		timeLogged = self.totalTime + time.time() - self.startTime
		timetext = 'Total time logged: %ih%02im%02is\n' % (timeLogged/3600, (timeLogged/60)%60, timeLogged%60)
		self.pieCanvas.create_text(10,510, text = timetext, anchor=Tkinter.W)

	def minimize(self):
		self.root.iconify()

class RepeatTimer(): 
	def __init__(self, app): 
		self.running = False
		self.app = app
		
	def start(self): 		
		self.running = True
		self.update()
		
	def update(self):
		if self.running:
			try:
				# find active window and increment its count
				self.app.totalticks += 1
				
				windowname = win32gui.GetWindowText(win32gui.GetForegroundWindow())
				windowname = self.filter(windowname)
				if windowname == '':
					pass
				elif windowname in self.app.log:
					self.app.log[windowname] += 1
				else:
					self.app.log[windowname] = 1 				
				
				# update GUI
				self.app.updatePie()
			
			except Exception as e:
				#print 'Exception : %s' % e
				self.stop()
				return
	
			self.app.root.after(500, self.update)
			
	# Some manipulations of the actual window names in order to get more meaningful entries
	def filter(self,windowname):
		windowname = windowname.replace('*', '')
		
		# Common browsers
		if re.search(r'chrome|firefox|safari|internet explorer|opera', windowname, re.I): 
			if 'Facebook' in windowname:
				return 'Facebook'
			if 'Stack Overflow' in windowname:
				return "StackOverflow"
			if re.search(r'forum', windowname, re.I):
				return 'Forums'
			if '- MATLAB' in windowname:
				return 'Matlab documentation'
			if re.search(r'Python.*documentation', windowname):
				return 'Python documentation'
			if 'Wikipedia,' in windowname:
				return 'Wikipedia'
			if '@gmail.com - Gmail' in windowname:
				return 'Gmail'
			return 'Misc. Internet'
			
		# Matlab subwindows
		if re.search(r'Editor.*\.m$', windowname):
			return 'Matlab Editor'
		if re.search(r'Figure \d', windowname):
			return 'Matlab Figures'
		
		# PuTTy variations
		if 'PuTTy' in windowname:
			return 'PuTTy'
		if re.search(r'\[screen \d', windowname):
			return 'PuTTy'
			
		# 'Filename - Program' --> Program
		if '-' in windowname:
			windowname = windowname.split('-')[-1]
			
		# 'C:\stuff\location\folder' -> folder
		if ':\\' in windowname:
			windowname = windowname.split('\\')[-1]
		
		# What was this for again?
		if ': ' in windowname:
			windowname = windowname.split(':')[0]
			
		if getIdleDuration() > IDLE_TIMEOUT:
			windowname = "Idle"
		
		return windowname.strip()
	
	def stop(self):
		if self.running: 
			self.running = False

class LASTINPUTINFO(Structure):
	"""ctypes structure that receives the output of the win32 function GetLastInputInfo"""
	_fields_ = [
        ('cbSize', c_uint),
        ('dwTime', c_uint),
    ]

def getIdleDuration():
	"""Get the user's idle duration in seconds"""
	lastInputInfo = LASTINPUTINFO()
	lastInputInfo.cbSize = sizeof(lastInputInfo)
	windll.user32.GetLastInputInfo(byref(lastInputInfo))
	millis = windll.kernel32.GetTickCount() - lastInputInfo.dwTime
	return millis / 1000.0

if __name__ == "__main__":
	root = Tkinter.Tk()
	app = soglog(root)
	root.title("Soglog "+VERSION)
	root.iconbitmap("soglog.ico")
	root.mainloop()	
	