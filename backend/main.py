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
        """Processa documento completo com fluxo otimizado"""
        start_time = time.time()
        
        try:
            print("\n" + "="*60)
            print("INICIANDO PROCESSAMENTO DO DOCUMENTO")
            print("="*60)
            
            # 1. Lê o documento
            print("\n[1/7] Lendo documento...")
            reader = DocumentReader(file_path)
            paragraphs = reader.read_paragraphs()
            doc_info = reader.get_document_info()
            
            print(f"✓ Documento lido com sucesso:")
            print(f"  - Total de elementos: {len(paragraphs)}")
            print(f"  - Parágrafos: {doc_info['total_paragraphs']}")
            print(f"  - Imagens: {doc_info['total_images']}")
            print(f"  - Tabelas: {doc_info['total_tables']}")
            
            # 2. Processa com IA (com contexto melhorado)
            print("\n[2/7] Processando com IA...")
            ai_processor = AIProcessor(api_key)
            ai_results = ai_processor.process_document(paragraphs, styles, removal_prompts)
            marked_content = ai_results['marked_content']
            
            print(f"✓ Processamento com IA concluído:")
            print(f"  - Elementos marcados: {ai_results['stats']['marked']}")
            print(f"  - Elementos sem marcação: {ai_results['stats'].get('unmarked', 0)}")  # CORRIGIDO: usa .get() com valor padrão
            print(f"  - Taxa de marcação: {(ai_results['stats']['marked']/ai_results['stats']['total_paragraphs']*100):.1f}%")
            
            # Validação crítica
            if ai_results['stats']['marked'] == 0:
                raise Exception("ERRO CRÍTICO: Nenhum elemento foi marcado pela IA!")
            
            # 3. Aplica estilos (com garantia de aplicação)
            print("\n[3/7] Aplicando estilos...")
            style_applier = StyleApplier(file_path)
            style_applier.register_styles(styles)
            styled_doc = style_applier.apply_styles(marked_content)
            
            # 4. Remove conteúdo marcado (com rastreamento completo)
            print("\n[4/7] Removendo conteúdo marcado...")
            clean_doc = style_applier.remove_marked_content(
                styled_doc, marked_content, removal_prompts
            )
            
            # 5. Divide em simulados (PULAR - não queremos mais dividir)
            print("\n[5/7] Pulando divisão em simulados...")
            simulados = []  # Lista vazia - não divide mais
            
            print("✓ Divisão desabilitada - documento único será gerado")
            
            # 6. Cria documentos finais
            print("\n[6/7] Criando documento final...")
            documents = {}
            
            # Apenas o documento completo estilizado
            documents['completo'] = clean_doc
            print("  ✓ Documento único criado")
            
            # NÃO cria mais documentos separados
            
            # 7. Salva arquivos
            print("\n[7/7] Salvando arquivos...")
            file_manager = FileManager(book_name)
            output_dir = file_manager.create_output_structure()
            saved_files = file_manager.save_documents(documents)
            
            print(f"✓ Arquivos salvos em: {output_dir}")
            print(f"  - Total de arquivos: {len(saved_files)}")
            
            # Cria arquivo ZIP
            zip_path = file_manager.create_zip_archive()
            print(f"✓ Arquivo ZIP criado: {os.path.basename(zip_path)}")
            
            # Limpa arquivos temporários
            file_manager.cleanup_temp_files()
            
            # Calcula tempo de processamento
            processing_time = time.time() - start_time
            
            print("\n" + "="*60)
            print("PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
            print("="*60)
            print(f"Tempo total: {int(processing_time // 60)}m {int(processing_time % 60)}s")
            
            # Prepara resposta detalhada
            return {
                'success': True,
                'processing_time': f"{int(processing_time // 60)}m {int(processing_time % 60)}s",
                'stats': {
                    'total_pages': doc_info['total_paragraphs'],
                    'questions_processed': ai_results['stats']['marked'],
                    'api_calls': ai_results['stats']['api_calls'],
                    'removal_count': len(removal_prompts),
                    'styles_applied': len(styles)
                },
                'files': saved_files,
                'output_directory': output_dir,
                'zip_file': os.path.basename(zip_path),
                'details': {
                    'document_info': doc_info,
                    'ai_stats': ai_results['stats']
                }
            }
            
        except Exception as e:
            processing_time = time.time() - start_time
            error_msg = str(e)
            
            print("\n" + "="*60)
            print("ERRO NO PROCESSAMENTO!")
            print("="*60)
            print(f"Erro: {error_msg}")
            print(f"Tempo decorrido: {int(processing_time)}s")
            
            # Log detalhado do erro
            import traceback
            traceback.print_exc()
            
            return {
                'success': False,
                'error': error_msg,
                'processing_time': f"{int(processing_time)}s",
                'details': {
                    'stage': self._identify_error_stage(error_msg),
                    'suggestion': self._get_error_suggestion(error_msg)
                }
            }
    
    def _identify_error_stage(self, error_msg: str) -> str:
        """Identifica em que estágio ocorreu o erro"""
        if 'lendo documento' in error_msg.lower():
            return 'reading'
        elif 'api' in error_msg.lower() or 'openai' in error_msg.lower():
            return 'ai_processing'
        elif 'estilo' in error_msg.lower():
            return 'styling'
        elif 'remoção' in error_msg.lower():
            return 'removal'
        elif 'simulado' in error_msg.lower():
            return 'splitting'
        elif 'arquivo' in error_msg.lower() or 'salvar' in error_msg.lower():
            return 'saving'
        else:
            return 'unknown'
    
    def _get_error_suggestion(self, error_msg: str) -> str:
        """Fornece sugestão baseada no erro"""
        error_lower = error_msg.lower()
        
        if 'api key' in error_lower or 'unauthorized' in error_lower:
            return 'Verifique se a API Key está correta e tem créditos disponíveis.'
        elif 'rate limit' in error_lower:
            return 'Limite de requisições atingido. Aguarde alguns minutos e tente novamente.'
        elif 'timeout' in error_lower:
            return 'Tempo limite excedido. Tente com um documento menor ou verifique sua conexão.'
        elif 'marcação' in error_lower or 'nenhum elemento' in error_lower:
            return 'Verifique se os prompts de identificação estão corretos para o tipo de documento.'
        elif 'arquivo' in error_lower:
            return 'Verifique se o arquivo .docx é válido e não está corrompido.'
        elif 'memória' in error_lower or 'memory' in error_lower:
            return 'Documento muito grande. Tente processar em partes menores.'
        else:
            return 'Verifique os logs detalhados e tente novamente.'


# Classe auxiliar para monitorar progresso (opcional)
class ProgressMonitor:
    """Monitor de progresso para feedback em tempo real"""
    
    def __init__(self, callback=None):
        self.callback = callback
        self.current_step = ''
        self.current_progress = 0
        self.start_time = time.time()
    
    def update(self, step: str, progress: int, details: str = ''):
        """Atualiza o progresso"""
        self.current_step = step
        self.current_progress = progress
        
        if self.callback:
            self.callback({
                'step': step,
                'progress': progress,
                'details': details,
                'elapsed_time': time.time() - self.start_time
            })
        
        # Log no console
        print(f"[{progress:3d}%] {step}: {details}")
    
    def complete(self):
        """Marca como completo"""
        self.update('Processamento concluído!', 100)