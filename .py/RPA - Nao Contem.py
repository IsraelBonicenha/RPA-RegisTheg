# Importações --------------------------------------------------------------------------------------------------------------
import pandas as pd
import pyautogui
import time



# Variáveis Principais -----------------------------------------------------------------------------------------------------
url_excel = "https://docs.google.com/spreadsheets/etc..."

dt_hashtags = pd.read_excel(url_excel,  sheet_name='nome-da-aba')
dt_hashtags = dt_hashtags.iloc[:, [0,1]].fillna("")
dt_hashtags.drop_duplicates()



# Código do Sistema --------------------------------------------------------------------------------------------------------
pyautogui.PAUSE = 0.15                                     # Define o tempo de execução para cada comando.
pyautogui.hotkey('alt', 'tab')                             # Altera para a janela do sistema.
pyautogui.click(1155, 315)                                 # Clica na janela do sistema. (!)

esperar_salvar = False                                     # Declara variáveis úteis para o funcionamento do loop.
index_marcação = 0                                         # ---------------------------------------------------------------

continue_loop = True                                       
while continue_loop:                                      

    if esperar_salvar:                                    # Espera o 5 minutos e volta para a célula inicial.
        time.sleep(300)                                   # ----------------------------------------------------------------                                                      
        pyautogui.press('enter', presses=2)               # ----------------------------------------------------------------
        pyautogui.press('tab', presses=4, interval=0.1)   # ----------------------------------------------------------------
        pyautogui.hotkey('ctrl', 'down')                  # ----------------------------------------------------------------


    for i in range(index_marcação, len(dt_hashtags)):      

        hashtag = dt_hashtags.iloc[i, 0]                   # Armazenas os valores a serem inseridos.
        valores = dt_hashtags.iloc[i, 1]                   # ----------------------------------------------------------------

        pyautogui.write(str(hashtag))                      # Preenche as células do DataGrid.
        pyautogui.press('right')                           # ----------------------------------------------------------------
        pyautogui.write(str(valores))                      # ----------------------------------------------------------------

        pyautogui.press('right')                           # Seleciona a opção ' Não Contém' e volta para a célula inicial.
        pyautogui.press('space')                           # ----------------------------------------------------------------
        pyautogui.press('Down', presses=2)                 # ----------------------------------------------------------------
        pyautogui.press('enter')                           # ----------------------------------------------------------------
        pyautogui.press('left', presses=2)                 # ----------------------------------------------------------------
        index_marcação += 1                                # ----------------------------------------------------------------
        
        if (index_marcação % 1000  == 0):                  # Pausa o loop for atual para o robô salvar 1000 hashtags.
            break                                          # ----------------------------------------------------------------
        
    
    pyautogui.press('tab', presses=4)                      # Navega até o botão 'Processar'.
    pyautogui.press('enter')                               # Pressiona 'Enter'
    esperar_salvar = True                                  # Var para ativar a opção de esperar salvar antes de iniciar loop.

    if index_marcação >= len(dt_hashtags):                 # Encerra o loop while se todas as linhas do dt forem processadas.
        continue_loop = False                              # ----------------------------------------------------------------



# ---------------------------------------------------------------------------------------------------------------------------