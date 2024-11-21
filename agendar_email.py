import schedule
import time
import subprocess
def executar_script():
    try:
        # Substitua pelo caminho completo para o script principal
        subprocess.run(["python", "relatorio.py"], check=True)
        print("Script executado com sucesso!")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao executar o script: {e}")


schedule.every().day.at("17:30").do(executar_script)

print("Agendador iniciado. Aguardando horário definido...")

while True:
    schedule.run_pending()  
    time.sleep(60)      
