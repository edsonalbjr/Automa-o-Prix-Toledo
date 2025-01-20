import pyautogui
import time
import pandas as pd
import threading  # Para monitorar a tecla Esc em paralelo
import keyboard  # Biblioteca para capturar eventos de teclado

# Variável global para sinalizar parada
stop_execution = False

# Função para monitorar a tecla Esc
def monitor_esc():
    """
    Função que fica em execução contínua em uma thread separada para monitorar a tecla Esc.
    Quando a tecla Esc é pressionada, a execução é interrompida definindo stop_execution como True.
    """
    global stop_execution
    while True:
        if keyboard.is_pressed("esc"):
            stop_execution = True
            print("Execução interrompida pelo usuário.")
            break

# Função para mover e clicar em uma coordenada com pausa entre os movimentos
def move_and_click(x, y, delay=0.2):
    """
    Mova o mouse para as coordenadas (x, y) e clica no local.
    Uma pausa de 'delay' segundos é adicionada entre os movimentos para evitar que os cliques sejam rápidos demais.
    """
    if stop_execution:
        exit_execution()
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.click()

# Função para mover, clicar duas vezes e limpar o campo
def move_and_double_click(x, y, delay=0.2):
    """
    Mova o mouse para as coordenadas (x, y) e realiza um duplo clique no local.
    Isso é útil, por exemplo, para limpar um campo de texto onde o valor por padrão já está definido como 0.
    
    Caso não seja realizado o duplo clique, o valor será inserido antes do 0 existente, o que pode resultar em valores incorretos.
    Exemplo: se a validade for 15 dias e o campo já contiver o valor 0, ao digitar '15', o valor final será '150'.
    Com o duplo clique, o valor '0' é selecionado e pode ser apagado, permitindo que o novo valor '15' seja inserido corretamente.
    """
    if stop_execution:
        exit_execution()
    pyautogui.moveTo(x, y, duration=delay)
    pyautogui.doubleClick()  # Duplo clique para selecionar o valor existente

# Função para digitar um texto
def type_text(text, delay=0.1):
    """
    Digita um texto no campo atual com intervalo entre as teclas, definido pelo parâmetro 'delay'.
    """
    if stop_execution:
        exit_execution()
    pyautogui.typewrite(str(text), interval=delay)

# Função para finalizar a execução
def exit_execution():
    """
    Função chamada para finalizar a execução do script imediatamente.
    """
    print("Finalizando a execução imediatamente.")
    exit()

# Lê os dados da planilha
excel_path = "dados.xlsx"  # Caminho para sua planilha Excel
try:
    df = pd.read_excel(excel_path)
except Exception as e:
    print(f"Erro ao carregar a planilha: {e}")
    exit()

# Verifica se a planilha contém dados
if df.empty:
    print("A planilha está vazia. Verifique os dados e tente novamente.")
    exit()

# Percurso com preenchimento dos dados
if __name__ == "__main__":
    """
    Função principal que executa o preenchimento de dados nos campos da interface gráfica,
    interagindo com os campos e inserindo os dados da planilha Excel.
    """
    # Inicia o monitoramento da tecla Esc em uma thread separada
    esc_thread = threading.Thread(target=monitor_esc, daemon=True)
    esc_thread.start()

    # Pausa inicial para preparar o ambiente
    time.sleep(1)

    for index, row in df.iterrows():
        # Verifica se a execução foi interrompida
        if stop_execution:
            break

        try:
            codigo_item = row['Codigo']  # Coluna da planilha que contém o código
            preco = row['Preco']  # Coluna da planilha que contém o Preco
            linha_descritivo = row['Descritivo1aLinha']  # Linha descritiva
            validade = row['DiasValidade']  # DiasValidade
            inf_extra = row['CodInfoExtra']  # Informação extra

            # Valida os campos obrigatórios
            if pd.isna(codigo_item) or pd.isna(linha_descritivo):
                print(f"Erro: Código ou 1ª linha descritivo ausente na linha {index + 1}. Item não cadastrado.")
                continue

            # 1. Botão Incluir
            move_and_click(1419, 357)

            # 2. Campo Codigo
            # move_and_click(614, 243) # Não é necessário clicar, já vai direto para o campo após "Incluir"
            type_text(codigo_item)

            # 3. Departamento - Menu Suspenso
            move_and_click(666, 273)

            # 4. Escolher Departamento
            move_and_click(678, 296)

            # 5. Campo Preco
            move_and_click(972, 308)
            type_text(preco)

            # 6. Descritivo1aLinha
            move_and_click(674, 342)
            type_text(linha_descritivo)

            # 7. DiasValidade
            move_and_double_click(1261, 346)  # Duplo clique apenas neste campo
            type_text(validade)

            # 8. CodInfoExtra - Menu Suspenso
            move_and_click(673, 547)

            # Seleciona a opção certa em CodInfoExtra com base no valor
            if inf_extra == 1:  # Ambiente
                move_and_click(679, 585)
            elif inf_extra == 2:  # Congelado
                move_and_click(672, 596)
            elif inf_extra == 3:  # Refrigerado
                move_and_click(674, 617)

            # 9. Botão Salvar
            move_and_click(1435, 322)

            # Dicionário para mapear os valores de inf_extra
            conserva_map = {
                1: "Ambiente",
                2: "Congelado",
                3: "Refrigerado"
            }

            # Linha de saída formatada
            conserva = conserva_map.get(inf_extra, "Desconhecido")  # Pega o valor mapeado, ou "Desconhecido" se não existir
            print(f"Código do Item: {codigo_item}, Produto: {linha_descritivo}, Preço: {preco}, Validade: {validade} dias, Conserva: {conserva}")

        except Exception as e:
            print(f"Erro ao processar o item na linha {index + 1}: {e}")
            continue  # Pula para o próximo registro

    print("Execução finalizada!")
