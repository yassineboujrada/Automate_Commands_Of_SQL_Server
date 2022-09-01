import pyodbc
import tkinter as tk
from tkinter import filedialog as fd

class window_tk(tk.Frame):
    
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        tk.Frame.grid(self,row=0,column=0,sticky="NW")
        tk.Frame.update(self)
        self.root = parent
        
        
        ###################################  designe my GUI  #################################

        self.root.geometry("600x350")
        self.root.title("Automate SQL Server")
        self.root.resizable(0,0)
        self.root.configure(background="grey28")

        self.path = ""

        self.warning=tk.Label(self.root, text="DataStory ",bg="grey28",font=("Lucida Sans", 12))

        self.server=tk.Label(self.root, text="Server Name :")
        self.db=tk.Label(self.root, text="DataBase Name :")
        self.serve_label=tk.Entry(self.root)
        self.db_label=tk.Entry(self.root)

        self.serve_label.place(x=235,y=82)
        self.db_label.place(x=235,y=132)

        self.server.configure(foreground="goldenrod2",bg="grey28",font=("Century Gothic", 11))
        self.db.configure(foreground="goldenrod2",bg="grey28",font=("Century Gothic", 11))

        self.server.place(x=75,y=80)
        self.db.place(x=75,y=130)

        self.cmd_file=tk.Label(self.root, text="Select File Of Commands :")
        self.cmd_file.configure(foreground="goldenrod2",bg="grey28",font=("Century Gothic", 11))
        self.cmd_file.place(x=75,y=180)

        self.open_button = tk.Button(
            self.root,
            text='Open a File',
            width=13, bg='SteelBlue3', fg='White',
            command=self.select_file
        )

        self.execute_queries = tk.Button(
            self.root,
            text='Execute',
            height=2, width=13, bg='DarkOrchid4', fg='White',
            command=self.command_
        )

        self.open_button.place(x=295,y=182)
        self.execute_queries.place(x=245,y=272)
        self.warning.configure(foreground="brown2",font=("Century Gothic", 17))
        self.warning.place(x=230,y=20)
    
    def select_file(self):
        filetypes = (
            ('text files', '*.txt'),
            ('All files', '*.*')
        )

        filename = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)

        if filename != "":
            self.path=filename
        else:
            pass


    def check_Sql_server(self):
        for d in pyodbc.drivers():
            if d in ["SQL Server","SQL Server Native Client 11.0"]:
                return True
        return False
    
    def command_(self):
        if self.check_Sql_server():
            Driver_NAME='SQL SERVER'
            Server_NAME = self.serve_label.get()
            DataBase_name = self.db_label.get()
            if self.path != "" and Server_NAME != "" and DataBase_name != "":
                with open(self.path) as f:
                    lines = f.readlines()
                    ls="".join(lines).split(";")
                    for i in ls:
                        querry="".join(i).strip()+";"
                        try:
                            cnxn_str = ("Driver={"+Driver_NAME+"};"
                                "Server="+Server_NAME+";"
                                "Database="+DataBase_name+";"
                                "Trusted_Connection=yes;")
                            cnxn = pyodbc.connect(cnxn_str)
                            cursor = cnxn.cursor()

                            c=cursor.execute(querry)

                            cnxn.commit()
                            self.warning.config(text = "Everything was going successfully",foreground="green4")
                            self.warning.place(x=180,y=20)
                        except:
                            self.warning.config(text = "There is an error in querry file or in input information")
                            self.warning.place(x=100,y=20)
            else:
                self.warning.config(text = "Fill Labels")
        else:
            self.warning.config(text = "there's no Sql Server in youre computer")
            self.warning.place(x=165,y=20)
        

if __name__ == "__main__":
    root = tk.Tk()
    window_tk(root)
    root.mainloop()