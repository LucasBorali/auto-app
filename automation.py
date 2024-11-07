import pyautogui
import time
import pandas as pd

pyautogui.PAUSE = 0.7

table_sheet = pd.read_excel('codigos.xlsx')


place = {
    'dimensional' : {'x': 241, 'y': 566},
    'material' : {'x': 375, 'y': 566},
    'desempenho' : {'x': 519, 'y': 566},
    'processo' : {'x': 709, 'y': 566},
}

type = {
    'dimensional' : {'x': 385, 'y': 393},
    'material' : {'x': 375, 'y': 566},
    'desempenho' : {'x': 366, 'y': 444},
    'processo' : {'x': 376, 'y': 464},
}


exclude_chars = False
change_type = False
change_number = False
change_codes_table = False
place_holder = 'choose place'
type_holder = 'choose type'

action = int(input('Digite o número indicado para tomar uma ação: Excluir características (1), Mudar tipo das características(2), Mudar código das características(3)==> '))

if action == 1:
    exclude_chars = True
elif action == 2:
    action = int(input('Tem uma tabela com os códigos? Sim(1) Não(2)'))
    if action == 1:
        change_codes_table = True
    elif action == 2:
        change_type = True
        question = int(input('escolha entre os tipos que deseja mudar: dimensional(1), material(2), desempenho(3), processo(4) ==> '))
        if question == 1:
            type_holder = 'dimensional'
        elif question == 2:
            type_holder = 'material'
        elif question == 3:
            type_holder = 'desempenho'
        elif question == 4:
            type_holder = 'processo'
elif action == 3:
    change_number = True
    question = int(input('escolha entre os locais que deseja voltar: dimensional(1), material(2), desempenho(3), processo(4) ==> '))
    if question == 1:
        place_holder = 'dimensional'
    elif question == 2:
         place_holder = 'material'
    elif question == 3:
         place_holder = 'desempenho'
    elif question == 4:
         place_holder = 'processo'
    



# Excluir características
while exclude_chars:
    for line in table_sheet.index:
        # Coordenada para selecionar primeira característica
        # x=375, y=661
        # pyautogui.click(x=375, y=661)
        
         # Editar número da característica
        pyautogui.click(x=228, y=368)
        
        # Apacar característica existente
        pyautogui.hotkey('ctrl','a')
        
        # Inserir código desejado
        pyautogui.write('0' + str(table_sheet.loc[line, 'Código']))

        # Coordenada para acessar a lixeira
        # x=384, y=954
        pyautogui.click(x=384, y=954)

        # Excluir característica
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.click(x=384, y=954)
        pyautogui.press('enter')
        pyautogui.press('left')
        pyautogui.press('enter')
        time.sleep(3)



# Mudar o tipo da característica
while change_type:
    pyautogui.click(x=375, y=661)
      
    #   edit
    pyautogui.click(x=228, y=955)
    
    # chage type 
    pyautogui.click(x=383, y=377)
    
    # select type = PROCESS
    pyautogui.click(x=type[type_holder]['x'], y=type[type_holder]['y'])
    
    # save
    pyautogui.click(x=277, y=951)
    pyautogui.press('enter')
    
    pyautogui.click(x=352, y=565)
    time.sleep(3)

# Mudar as características com tabela
if change_codes_table:
    for line in table_sheet.index:
        # Editar número da característica
        pyautogui.click(x=228, y=368)
        
        # Apacar característica existente
        pyautogui.hotkey('ctrl','a')
        
        # Inserir código desejado
        pyautogui.write(str(table_sheet.loc[line, 'Código']))
        
        # Editar
        pyautogui.click(x=225, y=955)
         
        # Selecionar caixa de tipos
        pyautogui.click(x=416, y=370)
        
        # Mandar para 'material' 
        pyautogui.click(x=353, y=419)
        
        # Salvar 
        pyautogui.click(x=277, y=958)
        
        # Apertar enter
        pyautogui.press('enter') 
        time.sleep(3)

# Mudar número das características
if change_number:
    for line in table_sheet.index:
        pyautogui.click(x=375, y=661)
            
            #   edit
        pyautogui.click(x=228, y=955)

            # change number
        pyautogui.click(x=244, y=376)

            # erase number
        pyautogui.hotkey('ctrl','a') 

            # write new number
        pyautogui.write(str(table_sheet.loc[line, 'Código']))

            # chage type 
        pyautogui.click(x=383, y=377)

            # select type = MATERIAL
        pyautogui.click(x=352, y=420)
            
            # save
        pyautogui.click(x=277, y=951)
        pyautogui.press('enter')

            # go back to 
        pyautogui.click(x=place[place_holder]['x'], y=place[place_holder]['y'])
        time.sleep(3)