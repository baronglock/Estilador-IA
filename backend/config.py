import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # API Keys
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
    
    # Directories
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
    OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
    TEMP_DIR = os.path.join(BASE_DIR, 'temp')
    
    # File settings
    MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
    ALLOWED_EXTENSIONS = {'docx'}
    
    # OpenAI settings
    GPT_MODEL = "gpt-4.1"
    MAX_TOKENS_PER_REQUEST = 4000
    TEMPERATURE = 0.3
    
    @staticmethod
    def create_directories():
        """Cria diretórios necessários se não existirem"""
        for directory in [Config.UPLOAD_DIR, Config.OUTPUT_DIR, Config.TEMP_DIR]:
            os.makedirs(directory, exist_ok=True)