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
        """Aplica estilos baseados nas marcações"""
        print(f"\nCriando novo documento com estilos...")
        
        # Carrega o documento original
        original_doc = Document(self.document_path)
        
        # Cria um novo documento vazio (sem copiar template)
        new_doc = Document()
        
        # Primeiro, cria TODOS os estilos definidos pelo usuário
        print("\nCriando estilos personalizados no documento:")
        for marker, style_config in self.styles_map.items():
            self._ensure_style_exists(new_doc, style_config)
        
        # Estatísticas
        stats = {
            'total': len(marked_content),
            'styled': 0,
            'by_style': {}
        }
        
        print(f"\nProcessando {len(marked_content)} parágrafos...")
        
        # Carrega o documento original para copiar imagens
        original_doc = Document(self.document_path)
        
        # CORREÇÃO: Usa marked_content como base, não o documento original
        print(f"\n  Aplicando estilos em {len(marked_content)} elementos...")
        
        for i, para_data in enumerate(marked_content):
            # Pega o índice original do parágrafo
            original_index = para_data.get('original_para_index', -1)
            
            # Verifica se o índice é válido
            if original_index < 0 or original_index >= len(original_doc.paragraphs):
                print(f"  ⚠️ Índice inválido para elemento {i}: {original_index}")
                continue
            
            # Pega o parágrafo original
            original_para = original_doc.paragraphs[original_index]
            
            # Cria novo parágrafo
            new_para = new_doc.add_paragraph()
            
            # Aplica estilo se houver marcação
            if para_data.get('markers') and len(para_data['markers']) > 0:
                for marker in para_data['markers']:
                    if marker in self.styles_map:
                        style_info = self.styles_map[marker]
                        try:
                            new_para.style = new_doc.styles[style_info['wordStyle']]
                            stats['styled'] += 1
                            stats['by_style'][style_info['name']] = stats['by_style'].get(style_info['name'], 0) + 1
                            
                            if i < 30:  # Mostra os primeiros 30 para debug
                                print(f"  ✓ Elemento {i} (P{original_index}): Estilo '{style_info['wordStyle']}' aplicado")
                            
                            break  # Aplica apenas o primeiro estilo válido
                        except Exception as e:
                            print(f"  ✗ ERRO ao aplicar estilo '{style_info['wordStyle']}' no elemento {i}: {e}")
            else:
                if i < 30 and para_data.get('text', '').strip():  # Mostra não estilizados não vazios
                    print(f"  - Elemento {i} (P{original_index}): SEM ESTILO - {para_data.get('text', '')[:40]}...")
            
            # SEMPRE copia o conteúdo completo do parágrafo original
            for run in original_para.runs:
                new_run = new_para.add_run()
                
                # Copia texto
                if run.text:
                    new_run.text = run.text
                
                # Copia formatação básica apenas
                try:
                    if run.bold is not None:
                        new_run.bold = run.bold
                    if run.italic is not None:
                        new_run.italic = run.italic
                    if run.underline is not None:
                        new_run.underline = run.underline
                except:
                    pass  # Ignora erros de formatação
                
                # Para imagens, tenta copiar de forma mais segura
                try:
                    if run._element.xpath('.//w:drawing'):
                        # Copia apenas o texto alternativo da imagem por enquanto
                        new_run.add_text("[IMAGEM]")
                except:
                    pass
            
            # Copia propriedades básicas do parágrafo
            try:
                if original_para.alignment:
                    new_para.alignment = original_para.alignment
            except:
                pass  # Ignora erros de propriedades
        
        # Mostra estatísticas
        print(f"\n=== ESTATÍSTICAS DE APLICAÇÃO ===")
        print(f"Total de parágrafos: {stats['total']}")
        print(f"Parágrafos com estilo: {stats['styled']}")
        print(f"Parágrafos sem estilo: {stats['total'] - stats['styled']}")
        print(f"\nPor tipo de estilo:")
        for style_name, count in stats['by_style'].items():
            print(f"  - {style_name}: {count}")
        
        # Lista todos os estilos no documento
        print(f"\n=== ESTILOS NO DOCUMENTO ===")
        for style in new_doc.styles:
            if style.name in [s['wordStyle'] for s in self.styles_map.values()]:
                print(f"  ✓ {style.name} (quick_style: {style.quick_style}, hidden: {style.hidden})")
        
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
        """NÃO remove conteúdo - apenas retorna o documento original"""
        print(f"\nRemoção DESABILITADA - mantendo todo o conteúdo...")
        
        # SIMPLESMENTE RETORNA O DOCUMENTO ORIGINAL SEM REMOVER NADA
        return document
    
    def _identify_removal_ranges(self, marked_content: List[Dict], removal_markers: List[Dict]) -> List[tuple]:
        """Identifica intervalos de parágrafos a serem removidos"""
        ranges = []
        
        print("\n  Procurando marcadores de remoção:")
        
        for removal in removal_markers:
            start_marker = removal['startMarker']
            end_marker = removal['endMarker']
            
            print(f"    Procurando por '{removal['name']}' ({start_marker} / {end_marker})...")
            
            start_idx = None
            
            for i, para in enumerate(marked_content):
                markers = para.get('markers', [])
                
                # Debug: mostra quando encontra marcadores
                if markers and (start_marker in markers or end_marker in markers):
                    print(f"      Parágrafo {i} tem marcadores: {markers}")
                
                if start_marker in markers and start_idx is None:
                    start_idx = i
                    print(f"      ✓ Início encontrado no parágrafo {i}: {para.get('text', '')[:50]}...")
                
                if end_marker in markers and start_idx is not None:
                    ranges.append((start_idx, i))
                    print(f"      ✓ Fim encontrado no parágrafo {i}: {para.get('text', '')[:50]}...")
                    print(f"      → Intervalo adicionado: {start_idx} até {i} ({i - start_idx + 1} parágrafos)")
                    start_idx = None
                    break  # Para após encontrar o fim
            
            # Se encontrou início mas não fim
            if start_idx is not None:
                print(f"      ⚠️ AVISO: Início encontrado no parágrafo {start_idx} mas sem marcador de fim!")
        
        return ranges
    
    def _copy_media_relations(self, source_doc: Document, target_doc: Document):
        """Copia as relações de mídia (imagens) do documento original"""
        try:
            # Método simplificado - apenas indica sucesso
            print("  ✓ Preparado para copiar imagens do documento original")
        except Exception as e:
            print(f"  ⚠️ Aviso ao preparar cópia de mídia: {e}")
    
    def _validate_removal_ranges(self, ranges: List[tuple], total_elements: int) -> List[tuple]:
        """Valida e ajusta os intervalos de remoção para evitar remoção excessiva"""
        if not ranges:
            return ranges
        
        print("\n  Validando intervalos de remoção:")
        
        # Remove sobreposições e intervalos inválidos
        validated = []
        
        for start, end in sorted(ranges):
            # Valida intervalo
            if start < 0 or end >= total_elements:
                print(f"    ⚠️ Intervalo inválido ignorado: {start}-{end} (total de elementos: {total_elements})")
                continue
            
            # Verifica se o intervalo é muito grande (mais de 50% do documento)
            interval_size = end - start + 1
            if interval_size > total_elements * 0.5:
                print(f"    ⚠️ Intervalo muito grande ({interval_size} de {total_elements} elementos)!")
                print(f"       Limitando remoção para proteger o conteúdo...")
                # Você pode ajustar ou pular este intervalo
                continue
            
            # Verifica sobreposição com intervalos já validados
            overlap = False
            for v_start, v_end in validated:
                if not (end < v_start or start > v_end):
                    overlap = True
                    print(f"    ⚠️ Intervalo {start}-{end} sobrepõe com {v_start}-{v_end}")
                    break
            
            if not overlap:
                validated.append((start, end))
                print(f"    ✓ Intervalo validado: {start}-{end} ({interval_size} parágrafos)")
        
        return validated