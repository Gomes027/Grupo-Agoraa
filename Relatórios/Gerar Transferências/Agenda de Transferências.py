from datetime import datetime
from src.Transferencias.Balanceamento import gerar_transferencia

# Obtendo a data de hoje
hoje = datetime.now()

# Descobrindo o dia da semana
dia_da_semana = hoje.strftime("%A")

# Criando uma estrutura if-else com base no dia da semana
if dia_da_semana == "Monday":
    variavel = "Hoje é segunda-feira"
    
elif dia_da_semana == "Tuesday":
    variavel = "Hoje é terça-feira"

elif dia_da_semana == "Wednesday":
    variavel = "Hoje é quarta-feira"

elif dia_da_semana == "Thursday":
    variavel = "Hoje é quinta-feira"

elif dia_da_semana == "Friday":
    variavel = "Hoje é sexta-feira"

else:
    variavel = "É fim de semana"

gerar_transferencia("SMJ", "STT", 1, 1.5, 2, 1.2, 0.2, 2, 0.5)
gerar_transferencia("VIX", "STT", 1, 2, 2, 1.2, 0.2, 2, 0.5)
gerar_transferencia("VIX", "SMJ", 1, 1.2, 2, 1.2, 0.2, 2, 0.5)