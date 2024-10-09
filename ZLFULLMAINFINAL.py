
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
import threading
import xlwings as xw
import openpyxl
#
import win32com.client as win32
import shutil

#
import time

#
from win32com import client
import time

#
import PIL

#
import os
import json
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

#set chromedriver.exe path
options = Options()
#options.add_argument('--remote-debugging-pipe')
options.add_argument("--headless")
#options.add_argument("--disable-gpu")
options.add_argument("--disable-images")
#options.add_argument("--disable-javascript")
#options.add_argument("--disable-css")
#options.add_argument(r"--user-data-dir=C:\Users\KN12QFB\AppData\Local\Google\Chrome\User Data") #e.g. C:\Users\You\AppData\Local\Google\Chrome\User Data

#options.add_argument(r'--profile-directory=Az Sz') #e.g. Profile 3

# driver = webdriver.Chrome(executable_path=r'C:\Users\KN12QFB\Downloads\chromedriver.exe',chrome_options=options)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)



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
    print(" printing ----------- BZ")
    # word = client.Dispatch("Word.Application")
    # # Open the document
    # doc = word.Documents.Open(filename)
    
    # # Set the printer if specified
    
    # if printer_name:
    #     word.ActivePrinter = printer_name
    # # Print the document
    # doc.PrintOut()
    # time.sleep(2)  # Wait for the print job to start
    
    # # Close the document
    # doc.Close()
    # # Example usage with a specific printer
    # # Quit the Word application
    # word.Quit()
    





#url launch
def Scraper(barcode):
    
    url_parts=["https://johndeere.service-now.com/nav_to.do?uri=%2Falm_asset_list.do%3Fsysparm_query%3Dasset_tagSTARTSWITH",barcode, "%26sysparm_first_row%3D1%26sysparm_view%3D%26sysparm_choice_query_raw%3D%26sysparm_list_header_search%3Dtrue" ]
    full_url = "".join(url_parts)
    #print(full_url)

    driver.get(full_url)
    array_datos=[]
    sty = time.process_time()
    
 

    

    delay=5
    try:
        WebDriverWait(driver, delay).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@id='gsft_main']"))) 

        #driver.switch_to.frame(iframe)
        asset_details = driver.find_elements(By.CLASS_NAME, "vt")
        bcode = asset_details[0].text
        array_datos.append(bcode)
        model = asset_details[1].text
        model = model[13:]
        array_datos.append(model)
        serial = asset_details[2].text
        array_datos.append(serial)
        user = (asset_details[7].text).split("(")[0]
        array_datos.append(user)
        '''
        Verificar qué es más rápido: Hacer el scraping de los datos del equipo desde SNow ó Leerlos del EXCEL 
        '''
        #print(f'Barcode | {bcode} \n Modelo | {model} \n Serial | {serial} \n Usuario | {user}')
        
        #driver.switch_to.default_content()

        driver.implicitly_wait(5)

    except TimeoutException:
        print(f'ELEMENTO NO ENCONTRADO')


    
    driver.quit()
    ety = time.process_time()
    elapsed_timey = ety - sty
    print("-------------------------------------")
    print('Execution time SCRAPER:', elapsed_timey, 'seconds')

    ''' 
    Escribir en la siguiente fila vacía de excel los valores de SERIAL y USER
    '''

    

    
    Scraper.array_datos = array_datos
    print("ARRAY DATOS", Scraper.array_datos)


#ULTIMAS FILAS PARA ESCRIBIR
def write_to_excel(valores, excel):
        stx = time.process_time()
        workbooklocation = excel
        sheetname = 'Sheet1'
        columnletter = 'G'

        with xw.App(visible=False) as app:
            with xw.Book(workbooklocation) as wb:
                sht = wb.sheets[sheetname]
                X = sht.range(columnletter + str(sht.cells.last_cell.row)).end('up').row + 1
                cell_SERIAL = columnletter + str(X)
                cell_USER = "I" + str(X)

                # Assuming valores[2] and valores[3] are single values
                sht.range(cell_SERIAL).value = [valores[2]]
                sht.range(cell_USER).value = [valores[3]]
        # stx = time.process_time()
        # workbooklocation = excel
        # sheetname = 'Sheet1'
        # columnletter = 'G'
        # xw.App(visible=False, update_links=False)




        # wb = xw.Book(workbooklocation, update_links=False)
        # X = wb.sheets[sheetname].range(columnletter + str(wb.sheets[sheetname].cells.last_cell.row)).end('up').row + 1
        # cell_SERIAL = columnletter + str(X)
        # cell_USER = "I" + str(X)
        # sht = wb.sheets['Sheet1']

        # sht.range(cell_SERIAL).value = valores[2]
        # sht.range(cell_USER).value = valores[3]

        # wb.save()
        # wb.close()
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
    # temp_template_path = original_template_path.split("\\")
    # temp_template_path.pop()
    original_template_path=original_template_path + "temp_template.docx"
    # temp_template_path = "\\".join(temp_template_path)
    #temp_template_path = 'C:\\Users\\KN12QFB\\OneDrive-Deere&Co\\OneDrive - Deere & Co\\Desktop\\Importantes\\temp_template.docx'

    # # Copy the original template to a temporary file
    # shutil.copyfile(original_template_path, temp_template_path)

    # Open Word application
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    # Open the temporary template document
    doc = word.Documents.Open(original_template_path)

    # Replace placeholders with new values
    for shape in doc.Shapes:
        if shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            for placeholder, new_value in new_values.items():
                if placeholder in text:
                    shape.TextFrame.TextRange.Text = text.replace(placeholder, new_value)

    # # Save the modified document
    # modified_doc_path = 'C:\\Users\\KN12QFB\\OneDrive-Deere&Co\\OneDrive - Deere & Co\\Desktop\\Importantes\\temp_template.docx'
    # doc.SaveAs(modified_doc_path)
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
            Tiempo_desde_boton_press = time.process_time()
            print("TIEMPO INICIO", Tiempo_desde_boton_press)
            if validate_input(self.entry.get()):
                st3 = time.process_time()
                #print(self.entry.get())
                print("BOTON PRESIONADO")
                Scraper(self.entry.get())
                ex_word(self.path_word, self.path_excel)
                et3 = time.process_time()
                elapsed_time3 = et3 - st3
                print('Execution time ALL FUNCTIONS:', elapsed_time3, 'seconds')
                printWordDocument("C:\\Users\\KN12QFB\\OneDrive-Deere&Co\\OneDrive - Deere & Co\\Desktop\\Importantes\\temp_template.docx", "ZDesigner ZT410-300dpi ZPL")
                Tiempo_desde_boton_press2 = time.process_time()
                tiempo_total = Tiempo_desde_boton_press2 - Tiempo_desde_boton_press
                print("EL TIEMPO TOTAL ES: ", tiempo_total)
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
            y = threading.Thread(target=write_to_excel(Scraper.array_datos, excel))
            z = threading.Thread(target=Valores_to_Word(Scraper.array_datos, word))
            y.start()
            z.start()
                
            y.join()
            z.join()
            

        
        #BARCODE LABEL
        self.label = customtkinter.CTkLabel(master=self.tab("Main"), text="BARCODE",font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"))
        self.label.grid(row=0, column=0, padx=20, pady=(25,0),columnspan = 2)
        
        self.labelx = customtkinter.CTkLabel(master=self.tab("Main"), text="0000", font=customtkinter.CTkFont(family="Arial", size=16))
        self.labelx.grid(row=1, column=0, padx=20,  pady=(0,75), sticky="e")

        #BARCODE ENTRY
        self.entry = customtkinter.CTkEntry(master=self.tab("Main"), placeholder_text="Ingrese barcode", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=13))
        #self.entry = "0000"+self.entry
        self.entry.grid(row=1, column=1, padx=20, pady=(0,75),sticky="w")
        self.entry.configure(justify="center")
        
        
        
        #TIPO DE USO LABEL
        self.label2 = customtkinter.CTkLabel(master=self.tab("Main"), text="USO", font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"))
        self.label2.grid(row=2, column=0, padx=20, pady=(0,0), columnspan = 2)

        #MENU DE USO
        optionmenu=customtkinter.CTkOptionMenu(master=self.tab("Main"), values=["JD Shared Services", "JD FINANCIAL", "Préstamos Sistemas", "Viajera"], command=optionmenu_callback)################## 
        optionmenu.set("JD Shared Services")
        optionmenu.grid(row=3, column=0, padx=40, pady=(0,5), sticky="nsew", columnspan = 2)

        #BOTÓN IMPRIMIR
        self.button = customtkinter.CTkButton(master=self.tab("Main"), text="Imprimir", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"),  command=start_new_thread)#self.button = customtkinter.CTkButton(self, text="Imprimir", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"),  command=threading.Thread(target=button_press).start)
        self.button.grid(row=5, column=0, padx=65, pady=(75, 0), sticky="nsew",columnspan = 2)

        #VALIDATE BARCODE
        self.labelbc = customtkinter.CTkLabel(master=self.tab("Main"), text="", font=customtkinter.CTkFont(family="Arial", size=12, weight="bold"))
        self.labelbc.grid(row=6, column=0, padx=0, pady=(5,0), columnspan=2)

        def validate_input(bc):
            print(type(bc))
            input_data = bc
            if len(input_data)==10:
                try:
                    int(input_data)
                    #self.labelbc.configure(text=f"Valid numeric value: {input_data}",fg_color="green",)
                    return True
                except ValueError:
                    self.labelbc.configure(text=f"Barcode debe ser 6 caracteres numéricos",text_color="#ff566d")
            else:
                self.labelbc.configure(text=f"Barcode debe ser 6 caracteres numéricos",text_color="#ff566d")
        #---------------------------------------------------------------
        #********************************SETUP TAB**************************************************
        
        #ENTRY CARPETA WORD
        self.label_word = customtkinter.CTkLabel(master=self.tab("Setup"), text="PATH WORD", font=customtkinter.CTkFont(family="Arial", size=16,weight="bold"))
        self.label_word.grid(row=0, column=1, padx=65, pady=(25,0), columnspan=2)
        self.label_word.configure(justify="center")

        
        self.entry_word = customtkinter.CTkEntry(master=self.tab("Setup"), placeholder_text="Path carpeta word", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=13))
        self.entry_word.grid(row=1, column=1, padx=65, pady=(5,0), columnspan=2)
        self.entry_word.configure(justify="center")
        self.entry_word.insert(0,self.path_word)

        #ENTRY EXCEL
        self.label_excel = customtkinter.CTkLabel(master=self.tab("Setup"), text="PATH EXCEL", font=customtkinter.CTkFont(family="Arial", size=16, weight="bold"))
        self.label_excel.grid(row=2, column=1, padx=65,  pady=(60,0), columnspan=2)
        self.label_excel.configure(justify="center")

        self.entry_excel = customtkinter.CTkEntry(master=self.tab("Setup"), placeholder_text="Path carpeta excel", corner_radius=35, font=customtkinter.CTkFont(family="Arial", size=13))
        self.entry_excel.grid(row=3, column=1, padx=65, pady=(5,0), columnspan=2)
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
        self.geometry("330x450")
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

