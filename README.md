# Macro VBA no excel
Este script VBA tem como objetivo automatizar a transferência de dados entre duas planilhas do Microsoft Excel: uma contendo informações de uma carteira de investimentos e outra para controle de despesas.

# Uso
Antes de executar o script, certifique-se de seguir estas instruções:

# Abra as planilhas envolvidas: "Planilha_Carteira.xlsx" e "Planilha_Despesas.xlsb".
Certifique-se de que as planilhas e intervalos mencionados no script existem nas respectivas pastas de trabalho.
Execute o script na planilha "Planilha_Despesas.xlsb".

# O script realiza as seguintes etapas:
Abre a planilha "Planilha_Carteira.xlsx".
Copia dados das planilhas "InfoGerais", "Disponibilidade", "Pagar", "Receber", "CotasAplicadas", "RendaFixa", "Imoveis" e "RendaVariavel".
Abre a planilha "Planilha_Despesas.xlsb".
Cole os dados copiados em intervalos específicos na planilha "BD".
Fecha a planilha "Planilha_Carteira.xlsx".

# Notas
Certifique-se de substituir os caminhos de arquivo pelos caminhos reais onde as planilhas estão localizadas.
Este script foi projetado para uso específico com as planilhas mencionadas e pode precisar de ajustes para funcionar com outras planilhas.
Aviso: Execute este script com cuidado, uma vez que manipula dados automaticamente e pode causar perda de informações se não for usado corretamente.

# Autora
Rafaela Bernardes Rabelo
