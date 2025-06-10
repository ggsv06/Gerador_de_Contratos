import generator as fgen
import FreeSimpleGUI as sg
import os
import sys

# CAMINHO DO ICO
if getattr(sys, 'frozen', False):  # If running as a compiled .exe
    icon_path = os.path.join(sys._MEIPASS, 'img.ico')
else:  # If running as a .py file
    icon_path = 'img.ico'

def item_row(item_num):
    """
    A "Row" in this case is a Button with an "X", an Input element and a Text element showing the current counter
    :param item_num: The number to use in the tuple for each element
    :type:           int
    :return:         List
    """
    sg.theme("LightBrown1")
    row =  [sg.pin(sg.Col([[sg.B("X", border_width=0, button_color=(sg.theme_text_color(), sg.theme_background_color()), k=('-DEL-', item_num), tooltip='Remover item'),
                            sg.In(size=(20,1), k=('-DESC-', item_num, 0)),
                            sg.In(size=(20,1), k=('-DESC-', item_num, 1)),
                            # sg.T(f'Key number {item_num}', k=('-STATUS-', item_num))
                            ]], k=('-ROW-', item_num)))]
    return row


def make_window(initial_rows):

    sg.theme("LightBrown1")
    layout = [  [sg.Text('\tChaves\t            Conteúdo', font='_ 15')],
                [sg.Col(initial_rows, k='-TRACKING SECTION-')],
                [sg.pin(sg.Text(size=(35,1), font='_ 8', k='-REFRESHED-',))],
                [sg.B('Nova linha', k='Add Item', tooltip='Nova linha'), sg.B('Cancelar', k='Cancel', tooltip='Cancela as ultimas alterações', button_color='red'),sg.Push(), sg.B('Gerar pdf', k='Gen pdf', tooltip='Gerar arquivo pdf', button_color=('black', '#3de226'))]]

    window = sg.Window('Editar chaves', layout, use_default_focus=False, font='_ 11', metadata=len(initial_rows)-1, finalize=True, icon=icon_path)

    return window

def get_current_dict(values, window):
    # Criar dic_temp
    dic_temp = {
        k: v for k, v in values.items()
        if ('-ROW-', k[1]) in window.AllKeysDict and window[('-ROW-', k[1])].visible
    }
    list_values_temp = list(dic_temp.values()) 
    # Criar dic oficial
    dic = {}
    for n in range(0, len(list_values_temp), 2):
        key = list_values_temp[n].strip()
        val = list_values_temp[n + 1].strip()
        if key:  # Only include non-empty keys
            dic[key] = val
    return dic

def main():
    # Carregar com informações anteriores
    try:
        pre_dic = fgen.read_json()
    except:
        pre_dic = {}
        if not os.path.exists('config'):
            os.makedirs('config')

    l_keys = list(pre_dic.keys())
    # Layout inicial
    initial_rows = []
    for i, key in enumerate(l_keys):
        initial_rows.append(item_row(i))
    
    if not initial_rows:
        initial_rows = [item_row(0)]
        
    window = make_window(initial_rows)
    for i, key in enumerate(l_keys):
        window[('-DESC-', i, 0)].update(key)

    while True:
        event, values = window.read()     # wake every hour
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Add Item':
            window.metadata += 1
            window.extend_layout(window['-TRACKING SECTION-'], [item_row(window.metadata)])
        elif event[0] == '-DEL-':
            window[('-ROW-', event[1])].update(visible=False)

        # Criar dic_temp
        dic = get_current_dict(values, window)
        # Gerar pdf
        if event == 'Gen pdf':
            try:
                fgen.write_json(dic)
                fgen.create_pdf(dic)
                sg.popup('\nArquivo pdf criado com sucesso!\n')
                break
            except:
                sg.popup('\nOcorreu um erro!\n\nVerifique se o arquivo "Contrato_template.docx" existe na pasta config.\nVerifique se as chaves estão corretas {{CHAVE}}.\nExemplo:\n\n\tMeu nome é {{NOME}}, nasci no dia {{DATA}}, ...\n\nVerifique se os campos deste programa estão de acordo.\nPrimeira coluna: variáveis contidas nas chaves\nSegunda coluna: valores para as variáveis\nExemplo:\n\n\tNOME\tGian\n\tDATA\t03/02/06\n', icon=icon_path)

        if event == 'Cancel':
            for i in range(window.metadata + 1):
                if ('-ROW-', i) in window.AllKeysDict:
                    window[('-ROW-', i)].update(visible=False)

            # Add fresh rows from JSON keys
            window.metadata = len(l_keys) - 1
            for i, key in enumerate(l_keys):
                if ('-ROW-', i) not in window.AllKeysDict:
                    window.extend_layout(window['-TRACKING SECTION-'], [item_row(i)])
                window[('-ROW-', i)].update(visible=True)
                window[('-DESC-', i, 0)].update(key)
                window[('-DESC-', i, 1)].update('')
    window.close()
main()