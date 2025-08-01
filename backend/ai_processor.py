import json
import time
import requests
from typing import List, Dict
from backend.config import Config

class AIProcessor:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.model = "gpt-4.1"  # Mantendo GPT-4.1 com sua capacidade total
        self.api_url = "https://api.openai.com/v1/chat/completions"
        print(f"AIProcessor inicializado com modelo: {self.model}")
        
    def process_document(self, paragraphs: List[Dict], styles: List[Dict], removal_prompts: List[Dict]) -> Dict:
        """Processa documento com IA para identificar estilos"""
        # Salva estilos e marcadores de remoção para validação posterior
        self.styles = styles
        self.removal_markers = []
        for removal in removal_prompts:
            self.removal_markers.append(removal['startMarker'])
            self.removal_markers.append(removal['endMarker'])
        
        marked_content = []
        processing_stats = {
            'total_paragraphs': len(paragraphs),
            'processed': 0,
            'marked': 0,
            'unmarked': 0,
            'api_calls': 0,
            'failed_batches': 0  # Contador de batches que falharam
        }
        
        print(f"Iniciando processamento de {len(paragraphs)} parágrafos...")
        
        # Processa em lotes maiores - GPT-4.1 aguenta muito mais
        batch_size = 150  # Voltando para 150 como solicitado
        
        # Segunda passada para parágrafos não marcados
        unmarked_paragraphs = []
        
        for i in range(0, len(paragraphs), batch_size):
            batch = paragraphs[i:i + batch_size]
            print(f"Processando batch {i//batch_size + 1} de {(len(paragraphs) + batch_size - 1)//batch_size}")
            
            # Tenta processar o batch com retry
            batch_results = None
            retry_count = 0
            max_retries = 2
            
            while batch_results is None and retry_count <= max_retries:
                if retry_count > 0:
                    print(f"  Tentativa {retry_count + 1} de {max_retries + 1}...")
                    time.sleep(1)  # Espera antes de retry
                
                batch_results = self._process_batch(batch, styles, removal_prompts)
                
                # Se falhou e ainda tem retries, reduz o batch
                if batch_results is None and retry_count < max_retries:
                    # Reduz o batch pela metade
                    if len(batch) > 10:
                        print(f"  Reduzindo tamanho do batch de {len(batch)} para {len(batch)//2}")
                        batch = batch[:len(batch)//2]
                    retry_count += 1
                else:
                    retry_count += 1
            
            # Se todas as tentativas falharam, usa batch sem marcações
            if batch_results is None:
                print(f"  AVISO: Batch {i//batch_size + 1} falhou após {max_retries + 1} tentativas. Continuando sem marcações.")
                batch_results = batch  # Retorna o batch original sem marcações
                processing_stats['failed_batches'] += 1
            
            marked_content.extend(batch_results)
            processing_stats['processed'] += len(batch_results)
            processing_stats['api_calls'] += retry_count
            
            # Pequena pausa para respeitar rate limits
            time.sleep(0.5)
        
        # Calcula estatísticas finais
        processing_stats['marked'] = sum(1 for p in marked_content if p.get('markers') and len(p['markers']) > 0)
        processing_stats['unmarked'] = processing_stats['total_paragraphs'] - processing_stats['marked']
        
        # Coleta parágrafos não marcados para segunda tentativa
        unmarked_paragraphs = [p for p in marked_content if not p.get('markers') or len(p['markers']) == 0]
        
        if unmarked_paragraphs and len(unmarked_paragraphs) > 10:
            print(f"\n⚠️ {len(unmarked_paragraphs)} parágrafos sem marcação. Fazendo segunda passada...")
            
            # Segunda tentativa focada nos não marcados
            for i in range(0, len(unmarked_paragraphs), 20):
                batch = unmarked_paragraphs[i:i + 20]
                print(f"  Reprocessando batch {i//20 + 1} de {(len(unmarked_paragraphs) + 19)//20}")
                
                # Passa o conteúdo completo para análise contextual
                batch_results = self._process_batch_focused(batch, styles, removal_prompts, marked_content)
                
                # Atualiza os resultados originais
                for updated_para in batch_results:
                    if updated_para.get('markers'):
                        # Encontra e atualiza no marked_content original
                        for j, original in enumerate(marked_content):
                            if original['index'] == updated_para['index']:
                                marked_content[j] = updated_para
                                break
                
                time.sleep(0.5)
            
            # Recalcula estatísticas
            processing_stats['marked'] = sum(1 for p in marked_content if p.get('markers') and len(p['markers']) > 0)
            processing_stats['unmarked'] = processing_stats['total_paragraphs'] - processing_stats['marked']
        
        print(f"\nProcessamento concluído:")
        print(f"  - {processing_stats['marked']} parágrafos marcados")
        print(f"  - {processing_stats['unmarked']} parágrafos sem marcação")
        if processing_stats['failed_batches'] > 0:
            print(f"  - {processing_stats['failed_batches']} batches falharam")
        
        # Log de exemplo dos não marcados para debug
        if unmarked_paragraphs and len(unmarked_paragraphs) <= 10:
            print("\nExemplos de parágrafos NÃO marcados:")
            for p in unmarked_paragraphs[:5]:
                print(f"  - P{p['index']}: {p['text'][:60]}...")
        
        return {
            'marked_content': marked_content,
            'stats': processing_stats
        }
    
    def _process_batch(self, batch: List[Dict], styles: List[Dict], removal_prompts: List[Dict]) -> List[Dict]:
        """Processa um lote de parágrafos usando a API"""
        system_prompt = self._build_system_prompt(styles, removal_prompts)
        user_prompt = self._build_user_prompt(batch)
        
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            "temperature": 0.3,
            "max_tokens": 8000  # Aumentado para aproveitar a capacidade do GPT-4.1
        }
        
        try:
            # Faz a requisição para a API
            response = requests.post(
                self.api_url, 
                headers=headers, 
                json=data,
                timeout=45  # Aumentado timeout
            )
            
            # Verifica se houve erro HTTP
            if response.status_code != 200:
                print(f"  Erro na API: Status {response.status_code}")
                print(f"  Resposta: {response.text[:200]}...")
                return None  # Retorna None para indicar falha
            
            # Extrai o conteúdo da resposta
            result_data = response.json()
            content = result_data['choices'][0]['message']['content']
            
            # Tenta extrair e corrigir JSON da resposta
            try:
                # Remove possível texto antes/depois do JSON
                json_start = content.find('{')
                json_end = content.rfind('}') + 1
                if json_start >= 0 and json_end > json_start:
                    content = content[json_start:json_end]
                
                # Tenta corrigir JSON truncado
                content = self._fix_truncated_json(content)
                
                # Parseia a resposta
                result = json.loads(content)
                return self._merge_results(batch, result)
                
            except json.JSONDecodeError as e:
                print(f"  Erro ao parsear JSON: {e}")
                print(f"  Conteúdo recebido (primeiros 500 chars): {content[:500]}")
                print(f"  Conteúdo recebido (últimos 200 chars): {content[-200:]}")
                print(f"  Tamanho total da resposta: {len(content)} caracteres")
                print(f"  Tentando correção automática...")
                
                # Verifica se o marcador está incorreto
                if '[[G' in content and ']]' not in content[content.find('[[G'):]:
                    print("  Detectado marcador incompleto - possível erro na IA")
                
                # Tenta uma correção mais agressiva
                fixed_content = self._aggressive_json_fix(content)
                if fixed_content:
                    try:
                        result = json.loads(fixed_content)
                        print("  ✓ JSON corrigido com sucesso!")
                        return self._merge_results(batch, result)
                    except:
                        pass
                
                print(f"  Falha na correção do JSON")
                return None
                
        except requests.exceptions.Timeout:
            print("  Timeout na requisição à API")
            return None
        except requests.exceptions.RequestException as e:
            print(f"  Erro na requisição HTTP: {e}")
            return None
        except Exception as e:
            print(f"  Erro inesperado: {type(e).__name__}: {str(e)}")
            return None
    
    def _fix_truncated_json(self, content: str) -> str:
        """Tenta corrigir JSON truncado"""
        # Remove espaços extras
        content = content.strip()
        
        # Se não termina com }, adiciona
        if not content.endswith('}'):
            # Conta brackets para tentar fechar corretamente
            open_brackets = content.count('[') - content.count(']')
            open_braces = content.count('{') - content.count('}')
            
            # Fecha arrays abertos
            content += ']' * open_brackets
            # Fecha objetos abertos
            content += '}' * open_braces
        
        return content
    
    def _aggressive_json_fix(self, content: str) -> str:
        """Tentativa mais agressiva de corrigir JSON"""
        try:
            # Encontra o array de paragraphs
            start = content.find('"paragraphs"') 
            if start == -1:
                return None
            
            # Extrai apenas a parte relevante
            start = content.find('[', start)
            if start == -1:
                return None
            
            # Tenta encontrar onde os dados realmente terminam
            paragraphs = []
            current_pos = start + 1
            
            # Parse manual básico
            while current_pos < len(content):
                # Procura por {"index":
                index_start = content.find('{"index":', current_pos)
                if index_start == -1:
                    break
                
                # Procura pelo fim deste objeto
                bracket_count = 1
                pos = index_start + 1
                obj_end = -1
                
                while pos < len(content) and bracket_count > 0:
                    if content[pos] == '{':
                        bracket_count += 1
                    elif content[pos] == '}':
                        bracket_count -= 1
                        if bracket_count == 0:
                            obj_end = pos
                    pos += 1
                
                if obj_end != -1:
                    try:
                        obj_str = content[index_start:obj_end + 1]
                        obj = json.loads(obj_str)
                        paragraphs.append(obj)
                        current_pos = obj_end + 1
                    except:
                        current_pos = index_start + 1
                else:
                    break
            
            if paragraphs:
                return json.dumps({"paragraphs": paragraphs})
            
        except Exception as e:
            print(f"    Erro na correção agressiva: {e}")
        
        return None
    
    def _process_batch_focused(self, batch: List[Dict], styles: List[Dict], removal_prompts: List[Dict], full_content: List[Dict]) -> List[Dict]:
        """Processa batch com foco especial em parágrafos não marcados, usando contexto"""
        system_prompt = """Você DEVE marcar TODOS os parágrafos abaixo. Analise cuidadosamente cada um.

IMPORTANTE: Use o CONTEXTO dos parágrafos anteriores e posteriores para decidir.
Por exemplo:
- Se antes tem uma questão e depois tem outra questão, o meio provavelmente são alternativas
- Se está entre duas alternativas, provavelmente é outra alternativa
- Se após um título vem texto normal, provavelmente é conteúdo/texto principal

REGRA PRINCIPAL: TODO parágrafo DEVE receber uma marcação se corresponder a algum estilo.

ESTILOS DISPONÍVEIS:
"""
        
        for style in styles:
            system_prompt += f"\n- {style['name']}: {style['prompt']}"
            system_prompt += f"\n  Marcador: {style['marker']}\n"
        
        system_prompt += "\nRETORNE APENAS JSON. Use o contexto para marcar corretamente."
        
        user_prompt = "ATENÇÃO: Estes parágrafos não foram marcados. Vou mostrar com CONTEXTO:\n\n"
        
        for para in batch:
            index = para['index']
            text = para['text'].strip()[:200]
            
            # Busca contexto (parágrafo anterior e próximo)
            prev_context = "INÍCIO DO DOCUMENTO"
            next_context = "FIM DO DOCUMENTO"
            
            # Encontra parágrafos vizinhos
            for p in full_content:
                if p['index'] == index - 1:
                    prev_text = p['text'][:100] if p['text'] else ""
                    prev_markers = p.get('markers', [])
                    prev_context = f"{prev_text} [Marcado como: {prev_markers[0] if prev_markers else 'SEM MARCAÇÃO'}]"
                elif p['index'] == index + 1:
                    next_text = p['text'][:100] if p['text'] else ""
                    next_markers = p.get('markers', [])
                    next_context = f"{next_text} [Marcado como: {next_markers[0] if next_markers else 'SEM MARCAÇÃO'}]"
            
            user_prompt += f"\n--- CONTEXTO DO PARÁGRAFO {index} ---\n"
            user_prompt += f"ANTERIOR: {prev_context}\n"
            user_prompt += f">>> ATUAL [NÃO MARCADO]: {text}\n"
            user_prompt += f"PRÓXIMO: {next_context}\n"
        
        user_prompt += "\nCom base no CONTEXTO, marque cada parágrafo apropriadamente!"
        
        # Usa a mesma lógica de processamento
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            "temperature": 0.1,  # Mais determinístico
            "max_tokens": 3000
        }
        
        try:
            response = requests.post(
                self.api_url, 
                headers=headers, 
                json=data,
                timeout=45
            )
            
            if response.status_code == 200:
                result_data = response.json()
                content = result_data['choices'][0]['message']['content']
                
                # Extrai JSON
                json_start = content.find('{')
                json_end = content.rfind('}') + 1
                if json_start >= 0 and json_end > json_start:
                    content = content[json_start:json_end]
                
                result = json.loads(content)
                return self._merge_results(batch, result)
        except Exception as e:
            print(f"    Erro na segunda passada: {e}")
        
        return batch
    
    def _analyze_residue_patterns(self, unmarked_paragraphs: List[Dict]) -> Dict:
        """Analisa padrões dos parágrafos não marcados para identificar se são resíduos"""
        analysis = {
            'empty': 0,
            'very_short': 0,
            'formatting': 0,
            'real_content': 0,
            'examples': []
        }
        
        for para in unmarked_paragraphs:
            text = para.get('text', '').strip()
            
            # Parágrafo vazio ou só espaços
            if not text or text.isspace():
                analysis['empty'] += 1
            # Muito curto (menos de 10 caracteres)
            elif len(text) < 10:
                analysis['very_short'] += 1
            # Possível resíduo de formatação (só pontuação, números isolados, etc)
            elif all(c in '.-_=~*#@$%&()[]{}|\\/:;,<>!? \t\n0123456789' for c in text):
                analysis['formatting'] += 1
            # Parece conteúdo real
            else:
                analysis['real_content'] += 1
                analysis['examples'].append({
                    'index': para['index'],
                    'text': text
                })
        
        return analysis
    
    def _build_system_prompt(self, styles: List[Dict], removal_prompts: List[Dict]) -> str:
        """Constrói o prompt do sistema com as definições de estilos"""
        prompt = """Você é um especialista em análise de documentos educacionais. 
Sua tarefa é identificar e marcar estilos em TODOS os parágrafos do documento.

REGRA FUNDAMENTAL: Você DEVE analisar e marcar TODOS os parágrafos listados.
Não pule NENHUM parágrafo. Se um parágrafo corresponde a algum estilo, MARQUE-O.

REGRAS IMPORTANTES:
1. TODOS os parágrafos devem ser analisados e marcados quando apropriado
2. Se um parágrafo contém questão E alternativas juntas, marque como questão/enunciado
3. Alternativas SEMPRE começam com letras (A, B, C, D, E) seguidas de ) ou .
4. Um parágrafo só pode ter UM tipo de marcação
5. Se identificou um padrão, CONTINUE aplicando em TODOS os casos similares
6. NÃO PULE parágrafos - se tem dúvida, marque com o estilo mais provável

EXEMPLOS OBRIGATÓRIOS:
- Qualquer linha começando com A) B) C) D) E) → SEMPRE marque como alternativa
- Qualquer linha com número seguido de . ou ) → SEMPRE marque como questão
- Linhas com "Resposta:" ou "Gabarito:" → SEMPRE marque como gabarito

ESTILOS A IDENTIFICAR:
"""
        
        for style in styles:
            prompt += f"\n- {style['name']}: {style['prompt']}"
            prompt += f"\n  Marcador a usar: {style['marker']}\n"
        
        if removal_prompts:
            prompt += "\nCONTEÚDO PARA MARCAR REMOÇÃO:\n"
            prompt += "IMPORTANTE: Só marque para remoção conteúdo que NÃO tem nenhum estilo aplicado!\n"
            prompt += "Se um parágrafo já tem um marcador de estilo, NÃO adicione marcadores de remoção.\n\n"
            
            for removal in removal_prompts:
                prompt += f"\n- {removal['name']}: {removal['prompt']}"
                prompt += f"\n  Marcadores: {removal['startMarker']} (início) e {removal['endMarker']} (fim)\n"
        
        prompt += """
FORMATO DE RESPOSTA OBRIGATÓRIO:
Retorne APENAS um objeto JSON válido e COMPLETO, sem truncar.
O JSON deve ter EXATAMENTE este formato:
{"paragraphs": [{"index": 0, "markers": ["[[MARCADOR]]"]}, {"index": 1, "markers": []}, ...]}

IMPORTANTE: 
- SEMPRE use marcadores com colchetes DUPLOS: [[MARCADOR]]
- NÃO use colchetes simples: [MARCADOR]
- Garanta que o JSON esteja COMPLETO e válido
- Cada parágrafo pode ter no MÁXIMO um marcador
- Se não tiver certeza, deixe sem marcação: "markers": []
- Use marcadores EXATAMENTE como definidos acima (copie e cole)
"""
        
        return prompt
    
    def _build_user_prompt(self, batch: List[Dict]) -> str:
        """Constrói o prompt do usuário com o lote de parágrafos"""
        prompt = "Analise os seguintes parágrafos e retorne as marcações em formato JSON:\n\n"
        
        for i, para in enumerate(batch):
            # Não limita o texto - GPT-4.1 pode processar textos completos
            text = para['text'].strip()
            prompt += f"Parágrafo {para['index']}:\n{text}\n\n"
        
        prompt += "\nRetorne o JSON COMPLETO para TODOS os parágrafos listados."
        
        return prompt
    
    def _merge_results(self, batch: List[Dict], ai_results: Dict) -> List[Dict]:
        """Mescla os resultados da IA com o batch original"""
        # Lista de marcadores válidos (estilos + remoções)
        valid_style_markers = [style['marker'] for style in self.styles] if hasattr(self, 'styles') else []
        valid_removal_markers = self.removal_markers if hasattr(self, 'removal_markers') else []
        all_valid_markers = valid_style_markers + valid_removal_markers
        
        # Cria um mapa de resultados
        results_map = {}
        for p in ai_results.get('paragraphs', []):
            if isinstance(p, dict) and 'index' in p and 'markers' in p:
                # Filtra e corrige marcadores
                markers = p.get('markers', [])
                if all_valid_markers:
                    validated_markers = []
                    for marker in markers:
                        # Corrige marcadores sem colchetes duplos
                        corrected_marker = marker
                        if not marker.startswith('[[') and not marker.endswith(']]'):
                            # Tenta adicionar colchetes
                            corrected_marker = f'[[{marker}]]'
                        
                        # Verifica se o marcador corrigido é válido
                        if corrected_marker in all_valid_markers:
                            validated_markers.append(corrected_marker)
                        elif marker in all_valid_markers:
                            validated_markers.append(marker)
                        else:
                            # Tenta sem colchetes também
                            inner_marker = marker.strip('[]')
                            test_marker = f'[[{inner_marker}]]'
                            if test_marker in all_valid_markers:
                                validated_markers.append(test_marker)
                            else:
                                print(f"    AVISO: Marcador inválido ignorado: {marker}")
                    
                    markers = validated_markers
                
                results_map[p['index']] = markers
        
        # Aplica os marcadores aos parágrafos
        for para in batch:
            para['markers'] = results_map.get(para['index'], [])
            
        return batch