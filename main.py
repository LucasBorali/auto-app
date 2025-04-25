from routine import Routine 
from tkinter import *
from tkinter import messagebox, ttk, filedialog
import pyautogui
import time
from functools import partial
import pickle
import pandas as pd
import os




TIME_UNTIL_CLICK = 5

def set_pause_time(*args):
        pyautogui.PAUSE = float(speed_entry.get())


# ORGANIZAR O CÓDIGO
# Caminho para as planilhas dentro do próprio aplicativo

# poder renomear os objetos


def load_data(file_name="mapa_comandos.pkl"):
    try:
        with open(file_name, "rb") as file:
            return pickle.load(file)
    except FileNotFoundError:
        return {}

loaded_routes = load_data()


def update_routine_list():
    combo_routine['values'] = list(loaded_routes.keys())
    
def update_command_list():
    
    command_listbox.delete(0, END)
    
    selected_routine = combo_routine.get()
    
    if selected_routine in loaded_routes:
        
        
        for func in loaded_routes[selected_routine].functions:
        
            command_listbox.insert(END, f" {func.func.__name__}{func.args}")
            
        show_loop_times()
   
            
            
def show_loop_times():
    if loaded_routes[combo_routine.get()].table_sheet:
        df = pd.read_excel(f'./DATA_BASES/{loaded_routes[combo_routine.get()].table_sheet}')
        loop_label.config(text=f'O mapa será executado {len(df)} vezes')
    else:
        loop_label.config(text=f'O mapa será executado em loop')
        

            
            
            
            
            
# Funções para adicionar nas rotinas
def write_txt(e_text):
    # Escreve um texto
    if e_text == "":
        messagebox.showerror("Erro", "Digite algum texto")
    else:
        append_func(partial(pyautogui.write, e_text))

def add_pause_time(pause):
    # Adiciona uma pausa em segundos 
    try:
        if pause == '':
             messagebox.showerror("Erro", "Digite algum valor")
        else:
            value = float(pause)   
            append_func(partial(time.sleep,value))
    except ValueError:
        messagebox.showerror("Erro", "Precisa ser um valor numérico")
    
        

def click_point(double_click = False, right_click = False):
    if double_click and right_click:
        messagebox.showerror('Erro', "Selecionar apenas UM tipo de clique")
    else:
        # Adiciona um ponto de click normal, duplo ou botão direito 
        show_coords = ttk.Label(enrty_command_div, text='Leve o mouse para a posição do click')
        show_coords.grid(column=0,row=3)
        
        def countdown(i):
            if i > 0:
                show_coords.config(text=f'{i} segundos para definir a posição')
                enrty_command_div.after(1000, countdown, i - 1)
            else:
                coorX, coorY = pyautogui.position()
                show_coords.config(text=f'X={coorX}, Y={coorY}')
                
                if double_click:
                    append_func(partial(pyautogui.doubleClick,x=coorX,y=coorY))
                elif right_click:
                    append_func(partial(pyautogui.rightClick,x=coorX,y=coorY))
                else:
                    append_func(partial(pyautogui.click,x=coorX,y=coorY))
                enrty_command_div.after(5000, show_coords.destroy)
                    
        countdown(TIME_UNTIL_CLICK)
    
def hotkey(hotkey):
    # Adiciona comandos de teclas múltiplas
    if hotkey == '':
        messagebox.showerror('Error', 'Digite um comando')
    else:
        hotkey_list = hotkey.split(',')
        
    if len(hotkey_list) <= 1:
        messagebox.showerror('Error', 'É necessário digitar MAIS que UMA tecla')
    else:
        append_func(partial(pyautogui.hotkey,*hotkey_list))
    
    
def press_key(key):
    # Adiciona comando de tecla única
    if key == '':
        messagebox.showerror('Error', 'Digite uma tecla')
    else:
        append_func(partial(pyautogui.press, key))
    
# COMANDOS DO USUÁRIO


# Adicionar funções à rotina
def append_func(function):
    
    command_index = command_listbox.curselection()
   
    partial_func = partial(function)
    
    if command_index == ():
        loaded_routes[combo_routine.get()].add_function(partial_func)
    else:    
        loaded_routes[combo_routine.get()].add_function(partial_func, command_index[0])
    
    save_route()
    update_command_list()
    
    entry_txt_line.delete(0, END)
    entry_pause_time.delete(0, END)
    entry_hotkey.delete(0, END)
    entry_command_key.delete(0, END)

# Excluir comandos dentro das rotinas
def exclude_func():
    command_index = command_listbox.curselection()
    
    
    if combo_routine.get() == '':
        messagebox.showerror('Error', 'Por favor, selecione um mapa')
    elif command_index == ():
        loaded_routes[combo_routine.get()].exclude_function()
    else:
        loaded_routes[combo_routine.get()].exclude_function(command_index[0])
        
    save_route()
    update_command_list()

   
# Executar rota de comandos    
def execute_route():
    
    if combo_routine.get() == '':
        messagebox.showerror('Error', "Por favor, selecione um mapa")
    else:
        if loaded_routes[combo_routine.get()].table_sheet:
            df = pd.read_excel(f'./DATA_BASES/{loaded_routes[combo_routine.get()].table_sheet}')
            i = 1
            time.sleep(2)
            while i <= len(df):
                for function in loaded_routes[combo_routine.get()].functions:
                    function()
                i += 1
                time.sleep(1)
        else:
            while True:
                for function in loaded_routes[combo_routine.get()].functions:
                    function()
                
                
                    

# Atualizar mapa com as novas funções
def save_route(file_name="mapa_comandos.pkl"):
    with open(file_name, "wb") as file:
        pickle.dump(loaded_routes, file)
        
# Criar novos mapas
def create_objects(key, file_name='mapa_comandos.pkl', ):
    if key == '':
        messagebox.showerror('Error', 'Digite um nome, por favor')
    else:
        result = messagebox.askquestion('Configurar planilha', 'Deseja configurar uma planilha para o mapa de comandos?')
        
        if result == 'yes':
            file_path = filedialog.askopenfilename(title="Selecione um arquivo", 
                                                 initialdir='./DATA_BASES',
                                       filetypes=[("Arquivos excel", "*.xlsx"), 
                                                  ("Todos os arquivos", "*.*")])
            file_name = os.path.basename(file_path)
            if file_name == "":
                return
            else: 
                loaded_routes[key] = Routine(file_name)
           
        elif result == 'no':
            loaded_routes[key] = Routine()
            
        else:
            return
        
    save_route()
                
    entry_create_route.delete(0, END)
            
    update_routine_list()
            

# Excluir mapas
def exclude_routines():
    if combo_routine.get() == '':
        messagebox.showerror('Error', 'Por favor, selecione um mapa primeiro')
    else:
        result = messagebox.askokcancel('Excluir mapa de comandos', 'Tem certeza?')

        if result:
            loaded_routes.pop(combo_routine.get())
            
            save_route()
            
            update_routine_list()
            combo_routine.delete(0, END)
        else:
            return
      

# Interface
root = Tk()
root.title('TaskFy')
root.iconbitmap('./config/tech.ico')

speed_div = ttk.Frame()
speed_div.pack(pady=15)

label_status = ttk.Label(speed_div, text='Definir velocidade:').grid(column=0,row=0)


pause_var = StringVar()
pause_var.trace_add("write", set_pause_time)

speed_entry = ttk.Entry(speed_div, textvariable=pause_var, width=10)
speed_entry.grid(column=1, row=0)

combo_routine_div = ttk.Frame()
combo_routine_div.pack(pady=15)
combo_routine = ttk.Combobox(combo_routine_div, values=list(loaded_routes.keys()), width= 50)
combo_routine.pack()

combo_routine.bind("<<ComboboxSelected>>", lambda event: update_command_list())

enrty_command_div = ttk.Frame(root, padding=10)
enrty_command_div.pack()

entry_txt_line = ttk.Entry(enrty_command_div, width=30)
entry_txt_line.grid(column=0, row=0)
ttk.Button(enrty_command_div, width=25, cursor='hand2', text='Escrever texto', command=lambda: write_txt(entry_txt_line.get())).grid(column=1, row=0, padx=10)


entry_pause_time = ttk.Entry(enrty_command_div, width=30)
entry_pause_time.grid(column=0, row=1)
ttk.Button(enrty_command_div, width=25, cursor='hand2', text='Definir segundos de pausa', command=lambda: add_pause_time(entry_pause_time.get())).grid(column=1, row=1)

check_buttons = ttk.Frame(enrty_command_div)
check_buttons.grid(column=0, row=2)
double_click = BooleanVar(value=False)
right_click = BooleanVar(value=False)
ttk.Checkbutton(check_buttons, text='Clique duplo', variable=double_click, cursor='hand2').grid(column=0,row=0)
ttk.Checkbutton(check_buttons, text='Clique direito', variable=right_click, cursor='hand2').grid(column=1,row=0)
ttk.Button(enrty_command_div, width=25, cursor='hand2',  text='Definir local de click', command=lambda: click_point(double_click.get(), right_click.get())).grid(column=1, row=2)

entry_hotkey = ttk.Entry(enrty_command_div, width=30)
entry_hotkey.grid(column=0, row=4)
ttk.Button(enrty_command_div, width=25, cursor='hand2',  text='Combinação de comandos',command=lambda: hotkey(entry_hotkey.get())).grid(column=1, row=4)

entry_command_key = ttk.Entry(enrty_command_div, width=30)
entry_command_key.grid(column=0, row=5)
ttk.Button(enrty_command_div, width=25,  cursor='hand2', text='Comando de tecla', command= lambda: press_key(entry_command_key.get())).grid(column=1, row=5)

entry_create_route = ttk.Entry(enrty_command_div, width=30)
entry_create_route.grid(column=0, row=6)
ttk.Button(enrty_command_div, width=25,  cursor='hand2', text='Criar novo mapa', command= lambda: create_objects(key=entry_create_route.get())).grid(column=1, row=6)


command_listbox = Listbox(root, width=50 , height=7)
command_listbox.pack()


loop_label = ttk.Label(root)
loop_label.pack()



config_buttons_div = ttk.Frame()
config_buttons_div.pack(pady=15)
ttk.Button(config_buttons_div, padding=5, cursor='hand2', text='Executar mapa', command=execute_route).grid(column=0, row=0, padx=5.5)
ttk.Button(config_buttons_div, padding=5, cursor='hand2', text='Excluir mapa', command=exclude_routines).grid(column=1, row=0, padx=5.5)
ttk.Button(config_buttons_div, padding=5, cursor='hand2', text='Excluir comando', command=exclude_func).grid(column=2, row=0, padx=5.5)


root.mainloop()

