import os
import time
from typing import Dict, List
from backend.config import Config
from backend.document_reader import DocumentReader
from backend.ai_processor import AIProcessor
from backend.style_applier import StyleApplier
from backend.document_splitter import DocumentSplitter
from backend.file_manager import FileManager

class WordStylerProcessor:
    def __init__(self):
        Config.create_directories()
        
    def process_document(self, file_path: str, book_name: str, api_key: str, 
                        styles: List[Dict], removal_prompts: List[Dict]) -> Dict:
        """Processa documento completo"""
        start_time = time.time()
        
        try:
            # 1. Lê o documento
            print("Lendo documento...")
            reader = DocumentReader(file_path)
            paragraphs = reader.read_paragraphs()
            doc_info = reader.get_document_info()
            
            # 2. Processa com IA
            print("Processando com IA...")
            ai_processor = AIProcessor(api_key)
            ai_results = ai_processor.process_document(paragraphs, styles, removal_prompts)
            marked_content = ai_results['marked_content']
            
            # 3. Aplica estilos
            print("Aplicando estilos...")
            style_applier = StyleApplier(file_path)
            style_applier.register_styles(styles)
            styled_doc = style_applier.apply_styles(marked_content)

            # Debug: verifica se há marcações
            marked_count = sum(1 for p in marked_content if p.get('markers'))
            print(f"  Parágrafos com marcações: {marked_count} de {len(marked_content)}")
            if marked_count == 0:
                print("  AVISO: Nenhum parágrafo foi marcado pela IA!")
            
            # 4. Remove conteúdo marcado
            print("Removendo conteúdo marcado...")
            clean_doc = style_applier.remove_marked_content(
                styled_doc, marked_content, removal_prompts
            )
            
            # 5. Divide em simulados
            print("Dividindo simulados...")
            splitter = DocumentSplitter()
            simulados = splitter.split_simulados(clean_doc)

            # 6. Cria documentos separados
            print("Criando documentos separados...")
            documents = {}

            # Documento completo
            documents['completo'] = splitter.create_complete_document(clean_doc)

            # Documentos individuais por simulado
            split_docs = splitter.create_split_documents(simulados, clean_doc)
            documents.update(split_docs)
            
            
            # 7. Salva arquivos
            print("Salvando arquivos...")
            file_manager = FileManager(book_name)
            output_dir = file_manager.create_output_structure()
            saved_files = file_manager.save_documents(documents)
            
            # 8. Cria arquivo ZIP
            zip_path = file_manager.create_zip_archive()
            
            # 9. Limpa arquivos temporários
            file_manager.cleanup_temp_files()
            
            # Calcula tempo de processamento
            processing_time = time.time() - start_time
            
            return {
                'success': True,
                'processing_time': f"{int(processing_time // 60)}m {int(processing_time % 60)}s",
                'stats': {
                    'total_pages': doc_info['total_paragraphs'],
                    'questions_processed': ai_results['stats']['marked'],
                    'api_calls': ai_results['stats']['api_calls'],
                    'simulados_found': len(simulados)
                },
                'files': saved_files,
                'output_directory': output_dir,
                'zip_file': zip_path
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'processing_time': f"{int(time.time() - start_time)}s"
            }