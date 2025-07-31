from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from typing import List, Dict

class StyleApplier:
    def __init__(self, document_path: str):
        self.document_path = document_path
        self.styles_map = {}
        print(f"StyleApplier inicializado com documento: {document_path}")
        
    def register_styles(self, styles: List[Dict]):
        """Registra os estilos a serem aplicados"""
        print(f"Registrando {len(styles)} estilos...")
        for style in styles:
            self.styles_map[style['marker']] = style
            print(f"  - Estilo registrado: {style['name']} -> {style['wordStyle']} (marker: {style['marker']})")
    
    def apply_styles(self, marked_content: List[Dict]) -> Document:
        """Aplica estilos baseados nas marcações preservando imagens e tabelas"""
        print(f"\nCriando novo documento com estilos...")
        
        # Carrega o documento original
        original_doc = Document(self.document_path)
        
        # Cria um novo documento
        new_doc = Document()
        
        # Primeiro, cria TODOS os estilos definidos pelo usuário
        print("\nCriando estilos personalizados no documento:")
        for marker, style_config in self.styles_map.items():
            self._ensure_style_exists(new_doc, style_config)
        
        # Estatísticas
        stats = {
            'total': len(marked_content),
            'styled': 0,
            'images': 0,
            'tables': 0,
            'by_style': {}
        }
        
        print(f"\nProcessando {len(marked_content)} elementos...")
        
        # Processa cada elemento
        for i, element_data in enumerate(marked_content):
            element_type = element_data.get('type', 'paragraph')
            
            try:
                if element_type == 'image':
                    # Copia o elemento original que contém a imagem
                    if 'original_element' in element_data:
                        new_doc.element.body.append(element_data['original_element'])
                        stats['images'] += 1
                        print(f"  Imagem preservada no índice {i}")
                        
                        # Se tem estilo de imagem marcado
                        if element_data.get('markers'):
                            for marker in element_data['markers']:
                                if marker in self.styles_map:
                                    stats['styled'] += 1
                                    style_name = self.styles_map[marker]['name']
                                    stats['by_style'][style_name] = stats['by_style'].get(style_name, 0) + 1
                                    break
                                    
                elif element_type == 'table':
                    # Copia a tabela original
                    if 'original_element' in element_data:
                        new_doc.element.body.append(element_data['original_element'])
                        stats['tables'] += 1
                        print(f"  Tabela preservada no índice {i}")
                        
                else:  # paragraph
                    text = element_data.get('text', '').strip()
                    
                    if not text:
                        continue
                    
                    # Cria novo parágrafo
                    paragraph = new_doc.add_paragraph()
                    
                    # Aplica estilo se marcado
                    applied_style = False
                    if element_data.get('markers') and len(element_data['markers']) > 0:
                        for marker in element_data['markers']:
                            if marker in self.styles_map:
                                style_info = self.styles_map[marker]
                                style_name = style_info['wordStyle']
                                
                                try:
                                    paragraph.style = new_doc.styles[style_name]
                                    stats['styled'] += 1
                                    stats['by_style'][style_info['name']] = stats['by_style'].get(style_info['name'], 0) + 1
                                    applied_style = True
                                    
                                    if i < 20:
                                        print(f"  P{i}: Aplicado estilo '{style_name}' - {text[:50]}...")
                                    
                                    break
                                    
                                except Exception as e:
                                    print(f"  ERRO ao aplicar estilo '{style_name}' no parágrafo {i}: {e}")
                    
                    if not applied_style and i < 20:
                        print(f"  P{i}: SEM ESTILO - {text[:50]}...")
                    
                    # Adiciona o texto com formatação original se disponível
                    if 'runs' in element_data and element_data['runs']:
                        for run_data in element_data['runs']:
                            run = paragraph.add_run(run_data['text'])
                            if run_data.get('bold'):
                                run.bold = True
                            if run_data.get('italic'):
                                run.italic = True
                            if run_data.get('underline'):
                                run.underline = True
                            if run_data.get('font_size'):
                                run.font.size = Pt(run_data['font_size'])
                    else:
                        paragraph.add_run(text)
                        
            except Exception as e:
                print(f"  ERRO ao processar elemento {i}: {e}")
                continue
        
        # Mostra estatísticas
        print(f"\n=== ESTATÍSTICAS DE APLICAÇÃO ===")
        print(f"Total de elementos: {stats['total']}")
        print(f"Elementos com estilo: {stats['styled']}")
        print(f"Imagens preservadas: {stats['images']}")
        print(f"Tabelas preservadas: {stats['tables']}")
        print(f"Elementos sem estilo: {stats['total'] - stats['styled'] - stats['images'] - stats['tables']}")
        print(f"\nPor tipo de estilo:")
        for style_name, count in stats['by_style'].items():
            print(f"  - {style_name}: {count}")
        
        return new_doc
    
    def _ensure_style_exists(self, document: Document, style_config: Dict):
        """Garante que um estilo existe no documento com todas as configurações"""
        style_name = style_config['wordStyle']
        
        try:
            # Tenta criar o estilo
            style = document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            print(f"  ✓ Criado estilo: '{style_name}'")
            
            # Configurações essenciais para aparecer na galeria
            style.hidden = False
            style.quick_style = True
            style.priority = 1  # Alta prioridade
            
            # Baseado no estilo Normal
            style.base_style = document.styles['Normal']
            
            # Aplica cor se definida
            if 'color' in style_config:
                try:
                    color_hex = style_config['color'].lstrip('#')
                    r = int(color_hex[0:2], 16)
                    g = int(color_hex[2:4], 16)
                    b = int(color_hex[4:6], 16)
                    style.font.color.rgb = RGBColor(r, g, b)
                    print(f"    - Cor aplicada: {style_config['color']}")
                except Exception as e:
                    print(f"    - ERRO ao aplicar cor: {e}")
            
            # Nome amigável para a galeria (se diferente)
            if style_config.get('name'):
                try:
                    style.name = style_name  # Nome interno
                    # O nome de exibição é definido pelo próprio nome do estilo
                except:
                    pass
                    
        except ValueError as e:
            if "already in use" in str(e):
                # Estilo já existe, vamos atualizá-lo
                style = document.styles[style_name]
                print(f"  ! Atualizando estilo existente: '{style_name}'")
                
                style.hidden = False
                style.quick_style = True
                style.priority = 1
                
                if 'color' in style_config:
                    try:
                        color_hex = style_config['color'].lstrip('#')
                        r = int(color_hex[0:2], 16)
                        g = int(color_hex[2:4], 16)
                        b = int(color_hex[4:6], 16)
                        style.font.color.rgb = RGBColor(r, g, b)
                    except:
                        pass
            else:
                print(f"  ✗ ERRO ao criar estilo '{style_name}': {e}")
        except Exception as e:
            print(f"  ✗ ERRO inesperado ao criar estilo '{style_name}': {e}")
    
    def remove_marked_content(self, document: Document, marked_content: List[Dict], removal_markers: List[Dict]) -> Document:
        """Remove conteúdo marcado para remoção"""
        print(f"\nRemovendo conteúdo marcado...")
        
        # Identifica intervalos de remoção
        removal_ranges = self._identify_removal_ranges(marked_content, removal_markers)
        print(f"  Encontrados {len(removal_ranges)} intervalos para remoção")
        
        # Cria novo documento
        new_doc = Document()
        
        # Copia estilos personalizados
        for marker, style_config in self.styles_map.items():
            self._ensure_style_exists(new_doc, style_config)
        
        # Copia parágrafos não marcados para remoção
        removed_count = 0
        for i, para in enumerate(document.paragraphs):
            # Verifica se está em intervalo de remoção
            skip = False
            for start, end in removal_ranges:
                if start <= i <= end:
                    skip = True
                    removed_count += 1
                    break
            
            if not skip and para.text.strip():  # Não pula e não é vazio
                # Copia parágrafo
                new_para = new_doc.add_paragraph()
                
                # Tenta manter o estilo
                try:
                    new_para.style = para.style
                except:
                    pass
                
                # Copia o texto com formatação
                for run in para.runs:
                    new_run = new_para.add_run(run.text)
                    if run.bold is not None:
                        new_run.bold = run.bold
                    if run.italic is not None:
                        new_run.italic = run.italic
                    if run.underline is not None:
                        new_run.underline = run.underline
                    if run.font.size:
                        new_run.font.size = run.font.size
                    if run.font.color and run.font.color.rgb:
                        new_run.font.color.rgb = run.font.color.rgb
        
        print(f"  Removidos {removed_count} parágrafos")
        return new_doc
    
    def _identify_removal_ranges(self, marked_content: List[Dict], removal_markers: List[Dict]) -> List[tuple]:
        """Identifica intervalos de parágrafos a serem removidos"""
        ranges = []
        
        for removal in removal_markers:
            start_marker = removal['startMarker']
            end_marker = removal['endMarker']
            
            start_idx = None
            
            for i, para in enumerate(marked_content):
                markers = para.get('markers', [])
                
                if start_marker in markers and start_idx is None:
                    start_idx = i
                    print(f"    Início de remoção '{removal['name']}' no parágrafo {i}")
                
                if end_marker in markers and start_idx is not None:
                    ranges.append((start_idx, i))
                    print(f"    Fim de remoção '{removal['name']}' no parágrafo {i}")
                    start_idx = None
        
        return ranges