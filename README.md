
App de Poupança Pessoal
Este é um aplicativo simples de poupança pessoal desenvolvido em Python usando a biblioteca Tkinter para a interface gráfica, Openpyxl para manipulação de planilhas Excel e Matplotlib para criação de gráficos.

Funcionalidades
Salvar Valor Diário:

Insira um valor diário no campo designado.
Clique no botão "Salvar" para adicionar o valor a uma lista.
O valor é salvo em uma planilha Excel ("valores_diarios.xlsx").
A soma total dos valores é exibida na interface.
Visualizar Economia ao Longo do Tempo:

O botão "Gráficos" permite visualizar a economia ao longo do tempo.
Um gráfico de barras mostra a economia total por mês.
Um gráfico de pizza destaca o maior e o menor valor economizado.
Pré-requisitos
Certifique-se de ter as seguintes bibliotecas instaladas:

tkinter
openpyxl
matplotlib
bash
Copy code
pip install tkinter openpyxl matplotlib
Uso
Execute o script e utilize a interface para inserir e salvar valores diários. Clique no botão "Gráficos" para visualizar as estatísticas de economia.

bash
Copy code
python seu_script.py
Estrutura do Código
Funções:

salvar_valor(): Salva o valor diário na lista e na planilha.
plotar_grafico(): Gera gráficos de barras e pizza para visualização de dados.
Interface Gráfica:

Utiliza o Tkinter para criar uma interface simples e interativa.
Planilha Excel:

Os valores diários são armazenados na planilha "valores_diarios.xlsx".
Contribuição
Sinta-se à vontade para contribuir, relatar problemas ou sugerir melhorias. Abra uma issue ou envie um pull request!

Licença
Este projeto está sob a licença MIT.
