import os
import shutil
from datetime import datetime
from typing import Dict, List
from docx import Document
from backend.config import Config

class FileManager:
    def __init__(self, book_name: str):
        self.book_name = self._sanitize_filename(book_name)
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = None
        
    def _sanitize_filename(self, filename: str) -> str:
        """Remove caracteres inválidos do nome do arquivo"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename.strip()
    
    def create_output_structure(self) -> str:
        """Cria estrutura de pastas para os arquivos de saída"""
        # Cria pasta principal com nome do livro e timestamp
        folder_name = f"{self.book_name}_{self.timestamp}"
        self.output_dir = os.path.join(Config.OUTPUT_DIR, folder_name)
        
        # Cria diretórios
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Cria subpastas
        subdirs = ['completo', 'questoes', 'gabaritos']
        for subdir in subdirs:
            os.makedirs(os.path.join(self.output_dir, subdir), exist_ok=True)
        
        return self.output_dir
    
    def save_documents(self, documents: Dict[str, Document]) -> List[Dict]:
        """Salva todos os documentos gerados"""
        saved_files = []
        
        for doc_name, document in documents.items():
            # Determina o subdiretório apropriado
            if 'completo' in doc_name:
                subdir = 'completo'
            elif 'questoes' in doc_name:
                subdir = 'questoes'
            elif 'gabarito' in doc_name:
                subdir = 'gabaritos'
            else:
                subdir = ''
            
            # Cria caminho completo
            filename = f"{self.book_name}_{doc_name}.docx"
            file_path = os.path.join(self.output_dir, subdir, filename)
            
            # Salva documento
            document.save(file_path)
            
            # Adiciona informações do arquivo salvo
            saved_files.append({
                'name': filename,
                'path': file_path,
                'size': self._get_file_size(file_path),
                'type': subdir or 'other'
            })
        
        return saved_files
    
    def _get_file_size(self, file_path: str) -> str:
        """Retorna o tamanho do arquivo formatado"""
        size = os.path.getsize(file_path)
        
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        
        return f"{size:.1f} TB"
    
    def create_zip_archive(self) -> str:
        """Cria arquivo ZIP com todos os documentos"""
        zip_path = os.path.join(Config.OUTPUT_DIR, f"{self.book_name}_{self.timestamp}")
        shutil.make_archive(zip_path, 'zip', self.output_dir)
        
        return f"{zip_path}.zip"
    
    def cleanup_temp_files(self):
        """Remove arquivos temporários"""
        if os.path.exists(Config.TEMP_DIR):
            shutil.rmtree(Config.TEMP_DIR)
            os.makedirs(Config.TEMP_DIR, exist_ok=True)
    
    def get_output_summary(self, saved_files: List[Dict]) -> Dict:
        """Retorna resumo dos arquivos gerados"""
        summary = {
            'total_files': len(saved_files),
            'output_directory': self.output_dir,
            'files_by_type': {
                'completo': [],
                'questoes': [],
                'gabaritos': []
            }
        }
        
        for file_info in saved_files:
            file_type = file_info['type']
            if file_type in summary['files_by_type']:
                summary['files_by_type'][file_type].append({
                    'name': file_info['name'],
                    'size': file_info['size']
                })
        
        return summary