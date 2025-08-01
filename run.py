#!/usr/bin/env python
"""
Arquivo principal para iniciar a aplicação Word AI Styler
"""
import os
import sys
from api.routes import app
from backend.config import Config

def main():
    """Inicia a aplicação"""
    print("=== Word AI Styler ===")
    print("Iniciando servidor...")
    
    # Verifica configurações
    if not Config.OPENAI_API_KEY:
        print("ERRO: OPENAI_API_KEY não configurada no arquivo .env")
        sys.exit(1)
    
    # Cria diretórios necessários
    Config.create_directories()
    
    # Inicia servidor
    port = int(os.environ.get('FLASK_PORT', 5000))
    app.run(
        host='0.0.0.0',
        port=port,
        debug=os.environ.get('FLASK_ENV') == 'development'
    )

if __name__ == '__main__':
    main()