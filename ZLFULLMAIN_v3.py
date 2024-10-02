
import customtkinter
from hPyT import*

#AutoZLabel
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options

#
# import threading
import xlwings as xw

#
import win32com.client as win32
import shutil



#
from win32com import client
#
import PIL
import time

#
import os
import json





config_file = "C:\\Users\\KN12QFB\\OneDrive-Deere&Co\\OneDrive - Deere & Co\\Desktop\\Importantes\\path_config.json"

def load_paths():
    default_paths = {"path_word":"", "path_excel": ""}
    if not os.path.exists(config_file):
        with open(config_file, 'w') as f:
            json.dump(default_paths, f)
        return default_paths["path_word"], default_paths["path_excel"]

    with open(config_file, 'r') as f:
        paths=json.load(f)
    return paths.get("path_word", default_paths["path_word"]), paths.get("path_excel", default_paths["path_excel"])

def save_paths(path_word, path_excel):
    paths = {
        "path_word": path_word,
        "path_excel": path_excel
    }
    with open(config_file, 'w') as f:
        json.dump(paths,f)





def printWordDocument(filename, printer_name):
    word = client.Dispatch("Word.Application")
    # Open the document
    doc = word.Documents.Open(filename)
    
    # Set the printer if specified
    
    if printer_name:
        word.ActivePrinter = printer_name
    # Print the document
    doc.PrintOut()
    time.sleep(2)  # Wait for the print job to start
    
    # Close the document
    doc.Close()
    # Example usage with a specific printer
    # Quit the Word application
    word.Quit()
    


#ULTIMAS FILAS PARA ESCRIBIR
def write_to_excel(valores, excel):
        stx = time.process_time()
        workbooklocation = excel
        sheetname = 'Sheet1'
        columnletter = 'G'
        xw.App(visible=False)




        wb = xw.Book(workbooklocation)
        X = wb.sheets[sheetname].range(columnletter + str(wb.sheets[sheetname].cells.last_cell.row)).end('up').row + 1
        cell_SERIAL = columnletter + str(X)
        cell_USER = "I" + str(X)
        sht = wb.sheets['Sheet1']

        sht.range(cell_SERIAL).value = valores[2]
        sht.range(cell_USER).value = valores[3]

        wb.save()
        wb.close()
        etx = time.process_time()
        elapsed_timex = etx - stx
        print("-------------------------------------")
        print('Execution time EXCEL:', elapsed_timex, 'seconds')


        


#ESCRIBIR EXCEL




#ETIQUETA EN WORD
def Valores_to_Word(valores, word_path):
    print("VALORES WORD PATH: ", word_path)
    st2 = time.process_time()
    # Your variables for each run
    new_values = {
        'PCModel': valores[1],
        'Asset': valores[0],
        'Asignado': valores[3],
        'Serial': valores[2]
    }

    # Paths
    original_template_path = word_path#'C:\\Users\\KN12QFB\\OneDrive-Deere&Co\\OneDrive - Deere & Co\\Desktop\\Importantes\\IJD_Etiquetas_Laptop_Mini.docx'
    print("PATH WORD HERE: ", original_template_path)
    temp_template_path = original_template_path.split("\\")
    temp_template_path.pop()
    temp_template_path.append("temp_template.docx")
    temp_template_path = "\\".join(temp_template_path)
    #temp_template_path = 'C:\\Users\\KN12QFB\\OneDrive-Deere&Co\\OneDrive - Deere & Co\\Desktop\\Importantes\\temp_template.docx'

    # Copy the original template to a temporary file
    shutil.copyfile(original_template_path, temp_template_path)

    # Open Word application
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    # Open the temporary template document
    doc = word.Documents.Open(temp_template_path)

    # Replace placeholders with new values
    for shape in doc.Shapes:
        if shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            for placeholder, new_value in new_values.items():
                if placeholder in text:
                    shape.TextFrame.TextRange.Text = text.replace(placeholder, new_value)

    # Save the modified document
    modified_doc_path = 'C:\\Users\\KN12QFB\\OneDrive-Deere&Co\\OneDrive - Deere & Co\\Desktop\\Importantes\\temp_template.docx'
    doc.SaveAs(modified_doc_path)
    doc.Close()

    # Optionally, make Word visible and open the modified document
    #word.Visible = True
    #word.Documents.Open(modified_doc_path)
    et2 = time.process_time()
    elapsed_time2 = et2 - st2
    print("-------------------------------------")
    print('Execution time WORD:', elapsed_time2, 'seconds')

#########################




#########################################################################

  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("marsh.json")


#FRAME PRINCIPAL

        
class MyTabView(customtkinter.CTkTabview):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        #create tabs
        self.add("Main")
        self.add("Setup")
        self.path_word, self.path_excel = load_paths()

        


        #---------------------------------------------------------------
        def optionmenu_callback(choice):
                    print("option menu dropdown clicked: ", choice)


        def button_press():
            if validate_input(self.entry.get()):
                st3 = time.process_time()
                #print(self.entry.get())
                print("BOTON PRESIONADO")
                # Scraper(self.entry.get())
                ex_word(self.path_word, self.path_excel)
                et3 = time.process_time()
                elapsed_time3 = et3 - st3
                print('Execution time ALL FUNCTIONS:', elapsed_time3, 'seconds')
                printWordDocument("C:\\Users\\KN12QFB\\OneDrive-Deere&Co\\OneDrive - Deere & Co\\Desktop\\Importantes\\temp_template.docx", "ZDesigner ZT410-300dpi ZPL")
            else: 
                return
                               
        



        def start_new_thread():
            x = threading.Thread(target=button_press)
            if not x.is_alive():
                x.start()

        def start_new_thread2():
            y = threading.Thread(target=self.save_paths)
            if not y.is_alive():
                y.start()



        def ex_word(word, excel):
            pass
            # y = threading.Thread(target=write_to_excel(Scraper.array_datos, excel))
            # z = threading.Thread(target=Valores_to_Word(Scraper.array_datos, word))
            # y.start()
            # z.start()
                
            # y.join()
            # z.join()
            


        #USER LABEL
        self.label = customtkinter.CTkLabel(master=self.tab("Main"), text="USER",font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"))
        self.label.grid(row=0, column=0, padx=(5,5), pady=(5,0))
        

        #USER ENTRY
        self.entry = customtkinter.CTkEntry(master=self.tab("Main"), placeholder_text="Ingrese user", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=13))
        #self.entry = "0000"+self.entry
        self.entry.grid(row=1, column=0, padx=(5,5), pady=(0,15))
        self.entry.configure(justify="center")



        
        #BARCODE LABEL
        self.label = customtkinter.CTkLabel(master=self.tab("Main"), text="BARCODE",font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"))
        self.label.grid(row=0, column=1, padx=(0, 2), pady=(5,0))
        
        # self.labelx = customtkinter.CTkLabel(master=self.tab("Main"), text="0000", font=customtkinter.CTkFont(family="Arial", size=16))
        # self.labelx.grid(row=3, column=0, padx=20,  pady=(0,5))

        #BARCODE ENTRY
        self.entry = customtkinter.CTkEntry(master=self.tab("Main"), placeholder_text="Ingrese barcode", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=13))
        #self.entry = "0000"+self.entry
        self.entry.grid(row=1, column=1, padx=(0, 2), pady=(0,15))
        self.entry.configure(justify="center")


        #SERIAL LABEL
        self.label = customtkinter.CTkLabel(master=self.tab("Main"), text="SERIAL",font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"))
        self.label.grid(row=2, column=0, padx=(5,5), pady=(5,0))
        

        #SERIAL ENTRY
        self.entry = customtkinter.CTkEntry(master=self.tab("Main"), placeholder_text="Ingrese serial", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=13))
        #self.entry = "0000"+self.entry
        self.entry.grid(row=3, column=0, padx=(5,5), pady=(0,15))
        self.entry.configure(justify="center")
        


        #MODELO LABEL
        self.label2 = customtkinter.CTkLabel(master=self.tab("Main"), text="MODELO", font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"))
        self.label2.grid(row=2, column=1, padx=(0, 2), pady=(5,0))

        #MENU DE MODELO
        optionmenu=customtkinter.CTkOptionMenu(master=self.tab("Main"), values=["Z G10","840 G10", "Z G8", "840 G8"], command=optionmenu_callback)################## 
        optionmenu.set("Modelo")
        optionmenu.grid(row=3, column=1, padx=(0, 2), pady=(0,15))

        
        
        
        #TIPO DE USO LABEL
        self.label2 = customtkinter.CTkLabel(master=self.tab("Main"), text="TIPO", font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"))
        self.label2.grid(row=6, column=0, padx=(5,0), pady=(0,0))

        #MENU DE USO
        optionmenu=customtkinter.CTkOptionMenu(master=self.tab("Main"), values=["Shared Services", "FINANCIAL", "Préstamo", "Viajera"], command=optionmenu_callback)################## 
        optionmenu.set("Shared Services")
        optionmenu.grid(row=7, column=0, padx=(5,0), pady=(0,35))


        #CHECKBOX
        def checkbox_event():
            print("checkbox toggled, current value:", check_var.get())

        check_var = customtkinter.StringVar(value="on")
        self.checkbox = customtkinter.CTkCheckBox(master=self.tab("Main"), text="¿Certificada?", command=checkbox_event,
                                     variable=check_var, onvalue="on", offvalue="off")
        self.checkbox.grid(row=7, column=1, padx=(0, 2), pady=(0,35))






        #BOTÓN IMPRIMIR
        self.button = customtkinter.CTkButton(master=self.tab("Main"), text="Imprimir", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"),  command=start_new_thread)#self.button = customtkinter.CTkButton(self, text="Imprimir", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"),  command=threading.Thread(target=button_press).start)
        self.button.grid(row=10, column=0, padx=0, pady=(5, 0), columnspan=2)

        #---------------------------------------------------------------
        #********************************SETUP TAB**************************************************
        
        #ENTRY CARPETA WORD
        self.label_word = customtkinter.CTkLabel(master=self.tab("Setup"), text="PATH WORD", font=customtkinter.CTkFont(family="Arial", size=16,weight="bold"))
        self.label_word.grid(row=0, column=1, padx=75, pady=(25,0), columnspan=2)
        self.label_word.configure(justify="center")

        
        self.entry_word = customtkinter.CTkEntry(master=self.tab("Setup"), placeholder_text="Path carpeta word", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=13))
        self.entry_word.grid(row=1, column=1, padx=75, pady=(5,0), columnspan=2)
        self.entry_word.configure(justify="center")
        self.entry_word.insert(0,self.path_word)

        #ENTRY EXCEL
        self.label_excel = customtkinter.CTkLabel(master=self.tab("Setup"), text="PATH EXCEL", font=customtkinter.CTkFont(family="Arial", size=16, weight="bold"))
        self.label_excel.grid(row=2, column=1, padx=75,  pady=(60,0), columnspan=2)
        self.label_excel.configure(justify="center")

        self.entry_excel = customtkinter.CTkEntry(master=self.tab("Setup"), placeholder_text="Path carpeta excel", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=13))
        self.entry_excel.grid(row=3, column=1, padx=75, pady=(5,0), columnspan=2)
        self.entry_excel.configure(justify="center")
        self.entry_excel.insert(0,self.path_excel)

        #BOTON ACEPTAR
        self.button = customtkinter.CTkButton(master=self.tab("Setup"), text="Aceptar", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"),  command=start_new_thread2)#self.button = customtkinter.CTkButton(self, text="Imprimir", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"),  command=threading.Thread(target=button_press).start)
        self.button.grid(row=5, column=1, padx=65, pady=(85, 0), sticky="nsew",columnspan = 2)

        #********************************SETUP TAB**************************************************
    def save_paths(self):
            path_word = self.entry_word.get()
            #path_word_template = word +"\\temp_template.docx"
            path_excel = self.entry_excel.get()

            save_paths(path_word, path_excel)
            print(f"Saved path_word: {path_word}")
            print(f"Saved path_excel: {path_excel}")
        
        




#INIT APP
class App(customtkinter.CTk):
    
    def __init__(self):
        super().__init__()
        

        self.iconbitmap("label (1).ico")
        image = PIL.Image.open("JDW.jfif")
        background_image = customtkinter.CTkImage(image, size=(500, 500))
        b_lbl=customtkinter.CTkLabel(self, text="", image=background_image)
        b_lbl.place(x=0, y=0)
        self.title("ZLabel")
        self.geometry("375x450")
        self.resizable(0,0)
        self.grid_rowconfigure(0, weight=1)  # configure grid system
        self.grid_columnconfigure(0, weight=1)
        
        




        self.my_frame = MyTabView(master=self)
        #self.my_frame = MyFrame(master=self)
        self.my_frame.grid(row=0, column=0, padx=25, pady=25, sticky="nsew")
        #self.wm_attributes("-transparentcolor", "white")
        #self.my_frame.configure(fg_color="#367C2B")
        window_frame.center(self)



#CORRER APP

app = App()

app.mainloop()


