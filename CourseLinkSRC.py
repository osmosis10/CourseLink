'''
Author: Moses Lemma


Description: courseDataToolSRC.py is a simple tkinter application designed to give a basic UI to the 
             functions in courseUpdaterLib.py that automates certain steps of the PCL Learn course
             data updating process.

             
'''
# IMPORTS
import CourseUpdaterLib
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox
from tkinter import scrolledtext
import os

# Main class for the tkinter window
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("CourseLink")
        self.geometry('800x600')
        self.configure(bg='#00422c')
        self.resizable(False, False)


    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


# Class for frames, currently only one frame is ever used.
class MainFrame(tk.Frame):
    def __init__(self, container):
        super().__init__(container)   
        
        self.configure(background='#083f12')
        #-----------------------------------------------------------------------------------------------------------------------------------------------#        
        # BUTTONS
        
        # Update Providers
        self.updateButton = tk.Button(self, width=15, height=3,
                                      text='Update Providers', background="#fbbf3a", activebackground='#d4a43a',fg='black', command=self.updateProviders)
        self.updateButton.place(x=150, y=150)

        # Update Approval
        self.approvals = tk.Button(self, width=15, height=3, 
        text='Update Approval', background="#fbbf3a", activebackground='#d4a43a',
        fg='black', command=self.changeApprovalToCompleted)
        self.approvals.place(x=150, y=250)

        # Check for dupes
        self.dupeButton = tk.Button(self, width=15, height=3, 
        text='Check For Duplicates', background="#fbbf3a", activebackground='#d4a43a',
        fg='black', command=self.checkDuplicates)
        self.dupeButton.place(x=150, y=350)

        # Display 'Archive'
        self.dupeButton = tk.Button(self, width=15, height=3, 
        text='Display "Archive"', background="#fbbf3a", activebackground='#d4a43a',
        fg='black', command=self.ShowArchive) 
        self.dupeButton.place(x=150, y=450)
        #-----------------------------------------------------------------------------------------------------------------------------------------------#        
        # LABELS

        self.outputFont = tkfont.Font(family='Barlow', size=16, weight='bold')
        self.output_label = tk.Label(self, text="CONFIRM OUTPUT BELOW", foreground='#fbe279', bg='#083f12', font=self.outputFont)
        self.output_label.place(x=400, y=35)
        #-----------------------------------------------------------------------------------------------------------------------------------------------#        
        #OUTPUT BOX
        
        self.text_area = scrolledtext.ScrolledText(self, wrap=tk.WORD, width=50, height=30)
        self.text_area.tag_configure("bolded", font=("Barlow", 14, "bold"))
        self.text_area.configure(bg='#00422c', foreground='#fbbf3a')
        self.text_area.config(state=tk.DISABLED)
        self.text_area.place(x=350, y=75)
        #-----------------------------------------------------------------------------------------------------------------------------------------------#        
        # TOGGLE WARNINGS

        self.toggleFont = tkfont.Font(family='Barlow', size=12, weight='bold')
        self.warning_toggle = tk.Checkbutton(self, command=self.setWarningFlag, bg="#083f12", fg='#fbbf3a', relief='flat', 
                                             text="Toggle warnings", onvalue=1, offvalue=0, font=self.toggleFont)
        self.warning_toggle.place(x=50, y=50)

        self.warning_flag = True
        self.fileMsg = "One or more necessary files missing !"

        # show the frame on the container
        self.pack(fill=tk.BOTH, expand=True)


        ### OUTPUT TEXT FILE ###
        self.outputFname = self.newOutputFileName() # accounts for existing output files


    # Method that updates all the provider names in report1 from report2
    def updateProviders(self):

        # If files are missing, output 'one or more files missing'
        if (self.filesMissing()):
            return
        
        providerDict = CourseUpdaterLib.getAllProviders() # Get providers from report2
        flag, outputList = CourseUpdaterLib.updateReport2(providerDict)
        
        # Make output box active and delete contents
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete('1.0', tk.END)

        # updates performed
        if (flag == 1):
            count = 1
            with open(self.outputFname, 'a+') as f: # Save the output to text file
                f.write("\n\nCOURSE PROVIDER UPDATES\n\n")
                for value in outputList[:-1]:
                    self.text_area.insert(tk.INSERT, f"{count}. {value}", 'bolded')
                    self.text_area.insert(tk.INSERT, "\n\n")
                    f.write(f"{count}. {value}\n\n")
                    count += 1
                f.write(outputList[-1])
                f.close()
            self.text_area.insert(tk.INSERT, outputList[-1], 'bolded')
            self.text_area.config(state=tk.DISABLED)  
                
        # No updates were performed
        elif (flag == 0):
            self.text_area.insert(tk.INSERT, "No updates performed", 'bolded')

            if (self.warning_flag == True): # If warnings enabled
                    messagebox.showinfo(title="No updates performed", message=f"No updates performed or updates already performed.\nCheck current folder for latest updateRecords.txt output file.")


    # Method that changes linkedin learning courses status to 'Completed'
    def changeApprovalToCompleted(self):

        # If files are missing, output 'one or more files missing'
        if (self.filesMissing()):
            return
        
        # Change all linkedin learning course status to 'Completed' and return output
        flag, outputList = CourseUpdaterLib.updateReport1TrainingStatus()
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete('1.0', tk.END)

        # Updates performed, display output
        if (flag == 1):
            count = 0
            with open(self.outputFname, 'a+') as f: # Save the output to text file
                f.write("\n\nAPPROVAL UPDATES\n")
                for value in outputList[:-1]:
                    self.text_area.insert(tk.INSERT, value, 'bolded')
                    f.write(f"{value}\n")
                    count += 1

                f.write(outputList[-1])
            self.text_area.insert(tk.INSERT, outputList[-1], 'bolded')

        # No updates performed
        elif (flag == 0):
                self.text_area.insert(tk.INSERT, "No updates performed", 'bolded')
                if (self.warning_flag == True):
                    messagebox.showinfo(title="No updates performed", message="No updates performed or updates already performed.\nCheck current folder for updatedRecords.txt output file.")

        self.text_area.config(state=tk.DISABLED)


    def checkDuplicates(self):
        # If files are missing, write 'one or more files missing'
        if (self.filesMissing()):
            return

        # Check for duplicates and return any hits
        flag, outputList = CourseUpdaterLib.checkForDupesinReport2()
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete('1.0', tk.END)

        # Duplicates found and returned 
        if (flag == 1):
            for value in outputList[:-1]:
                self.text_area.insert(tk.INSERT, value, 'bolded')

            self.text_area.insert(tk.INSERT, outputList[-1], 'bolded')
            

        # No duplicates found
        elif (flag == 0):
            self.text_area.insert(tk.INSERT, 'No duplicates found', 'bolded')
        
        self.text_area.config(state=tk.DISABLED)


    def ShowArchive(self):
        # If files are missing, output 'one or more files missing'
        if (self.filesMissing()):
            return
        
        flag, outputList = CourseUpdaterLib.displayArchive()
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete('1.0', tk.END)

        # If any instance of 'Archive' found, return/output list
        if (flag == 1):
            count = 1
            with open(self.outputFname, 'a+') as f: # Save the output to text file
                f.write("\n\nARCHIVE\n\n")    
                for value in outputList[:-1]:
                    f.write(f"{count}. {value}\n\n")
                    self.text_area.insert(tk.INSERT, value, 'bolded')
                    count += 1
            
                f.write(outputList[-1])
            self.text_area.insert(tk.INSERT, outputList[-1], 'bolded')
            

        # If no instance of 'Archive'found
        elif (flag == 0):
            self.text_area.insert(tk.INSERT, 'No instance of "Archive" found', 'bolded')
            
            if (self.warning_flag == True):
                messagebox.showinfo(title="No updates performed", message="No updates performed or updates already performed.\nCheck current folder for updatedRecords.txt output file.")

        self.text_area.config(state=tk.DISABLED)



    ### HELPER METHODS ###

    # If toggle on, warning flag is False
    def setWarningFlag(self):
        if (self.warning_flag == True):
            self.warning_flag = False

        else:
            self.warning_flag = True


    def onClosing(self):
            
        # If warning flag is off then just close app
        if (not self.warning_flag):
            app.destroy()   # Destroy the main window
            return
    
        # Show warning to remind user to complete the manual steps 3 and 6
        messagebox.showwarning(title='Warning', message="NOTE: Complete steps 3 and 6 and 7\nfrom the instructions sheet before uploading")
        
        app.destroy()   # Destroy the main window

    # Check if files are in the folder, if flag is off, don't show warning
    def filesMissing(self):
        if (CourseUpdaterLib.checkIfAllFilesPresent() == False):
                self.text_area.config(state=tk.NORMAL)
                self.text_area.delete('1.0', tk.END)
                self.text_area.insert(tk.INSERT, "One or more files missing", 'bolded')
                self.text_area.config(state=tk.DISABLED)
                
                # If the warning toggle is off, don't show message box
                if (self.warning_flag == False):
                    return True
                
                messagebox.showwarning(title="Files missing !", message=self.fileMsg)
                return True
        return False
    

    def newOutputFileName(self):
        numTxtFiles = 0
        for file in os.listdir(os.curdir):
            if file.endswith('.txt') and file.startswith("updatedRecords"):
                numTxtFiles += 1
        
        if (numTxtFiles == 0):
            return 'updatedRecords.txt'
        
        else:
            return f'updatedRecords({numTxtFiles}).txt'


# Try running app
try:
    app = App()

    def main():
        frame1 = MainFrame(app)
        app.protocol("WM_DELETE_WINDOW", frame1.onClosing)
        app.mainloop()
    main()

except Exception as e:
    messagebox.showerror(title='Exception encountered', message=e)

