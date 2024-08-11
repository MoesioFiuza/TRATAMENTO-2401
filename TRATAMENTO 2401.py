import subprocess
import os
import sys
import time 

start_time = time.time()

def run_script(file_name):
    script_path = os.path.join(r'/home/moesiosf/Área de trabalho/TRATAMENTO', file_name)
    subprocess.run([sys.executable, script_path])

if __name__ == "__main__":
    scripts = [
        "Download de dados.py",
        "validações.py",
        "domicilio.py",
        "morador.py",
        "REORDENAR COLUNAS.py",
        "endereços deslocamento.py",
        "segundo tratamento deslocamento.py",
        "organizaçãofinal.py",
        "remover testes.py"
    ]
    
    for script in scripts:
        run_script(script)
    
    print("Scripts executados com sucesso.")

end_time = time.time()
execution_time = end_time - start_time    

print(execution_time)    