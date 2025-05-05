import subprocess
import os
import sys


def run_server():
    # Chemin du script python du serveur
    server_script = os.path.join(os.path.dirname(__file__), "excel_server.py")

    # Lancer le serveur Python
    print("Démarrage du serveur MCP Excel...")
    subprocess.run([sys.executable, server_script])


if __name__ == "__main__":
    run_server()