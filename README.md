# Cadastro de Produtos no MGV 7

Este repositório contém dois scripts para automatizar o cadastro de produtos no sistema MGV 7. Eles foram projetados para facilitar a inclusão de dados repetitivos e minimizar erros humanos durante o processo.

## Arquivos do Repositório

1. `cadastro_produtos_mgv7.py`: Script principal responsável por realizar o cadastro automático de produtos no sistema MGV 7.
2. `coordenadas_do_mouse.py`: Script auxiliar usado para mapear as coordenadas dos botões e campos necessários para o cadastro, conforme a resolução do monitor.

## Equipamentos Compatíveis com o MGV 7

O script foi desenvolvido para ser usado com o sistema MGV 7, que é compatível com os seguintes equipamentos:

- Balança Prix 4 Uno
- Impressoras de etiquetas integradas ao MGV 7
- Outros dispositivos configuráveis pelo sistema MGV 7

**Importante:** Certifique-se de que o sistema MGV 7 está instalado e configurado corretamente no seu equipamento.

## Configuração de Resolução

As coordenadas dos botões e campos utilizados no script foram mapeadas em um monitor com resolução **1920x1080**. Caso utilize um monitor com resolução diferente, será necessário executar o script `coordenadas_do_mouse.py` para remapear os pontos de clique.

## Dependências

Para executar os scripts, é necessário instalar as seguintes bibliotecas Python:

- `pyautogui`
- `pandas`
- `keyboard`

Instale-as com o comando:

```bash
pip install pyautogui pandas keyboard
```

## Uso

### Executando os Scripts

1. Certifique-se de que o sistema MGV 7 esteja aberto e pronto para receber os dados.
2. Coloque a planilha `dados.xlsx` com as informações dos produtos na mesma pasta do script.
3. Execute o script `cadastro_produtos_mgv7.py` como administrador. Para isso:
   - Se estiver usando o terminal: clique com o botão direito e escolha **Executar como administrador**.
   - Se estiver usando o Visual Studio Code: execute o programa como administrador (clique com o botão direito no ícone do VS Code e selecione **Executar como administrador**).

**Observação:** O script requer permissões de administrador para interagir corretamente com a interface do sistema MGV 7.

### Mapeando Novas Coordenadas

Se estiver usando um monitor com resolução diferente de **1920x1080**, siga os passos abaixo para remapear as coordenadas:

1. Execute o script `coordenadas_do_mouse.py`.
2. Mova o cursor do mouse até o local desejado e anote as coordenadas exibidas no terminal.
3. Substitua as coordenadas no script `cadastro_produtos_mgv7.py` pelas novas.

### Formato da Planilha

A planilha `dados.xlsx` deve conter as seguintes colunas:

- `Codigo`: Código do produto
- `Preco`: Preço do produto
- `Descritivo1aLinha`: Nome do produto
- `DiasValidade`: Dias de validade do produto
- `CodInfoExtra`: Código de informação extra referente ao tipo de conservação:
  - 1: Ambiente
  - 2: Congelado
  - 3: Refrigerado

Exemplo de planilha:
| Codigo | Preco | Descritivo1aLinha | DiasValidade | CodInfoExtra |
|--------|-------|----------------------------|--------------|--------------|
| 001 | 10 | Produto 1 | 10 | 1 |
| 002 | 20 | Produto 2 | 15 | 2 |
| 003 | 30 | Produto 3 | 20 | 3 |

## Fluxo do Script

1. O script carrega os dados da planilha `dados.xlsx`.
2. Para cada linha, realiza as seguintes ações:
   - Preenche o código do produto.
   - Seleciona o departamento.
   - Preenche o preço, nome, validade e código de conservação.
   - Salva o registro.
3. Durante a execução, é possível interromper o processo pressionando a tecla **Esc**.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias.

## Licença

Este projeto está licenciado sob a MIT License. Consulte o arquivo LICENSE para mais informações.
