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
        # Salva estilos para validação posterior
        self.styles = styles
        
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
        
        # Processa em lotes para otimizar
        batch_size = 150  # Aumentado de 100 para 150 - melhor contexto
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
        
        print(f"\nProcessamento concluído:")
        print(f"  - {processing_stats['marked']} parágrafos marcados")
        print(f"  - {processing_stats['unmarked']} parágrafos sem marcação")
        print(f"  - {processing_stats['failed_batches']} batches falharam")
        
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
    
    def _build_system_prompt(self, styles: List[Dict], removal_prompts: List[Dict]) -> str:
        """Constrói o prompt do sistema com as definições de estilos"""
        prompt = """Você é um especialista em análise de documentos educacionais. 
Sua tarefa é identificar e marcar estilos em textos de simulados educacionais.

REGRAS IMPORTANTES:
1. Cada linha/parágrafo deve ser analisado INDEPENDENTEMENTE
2. Se um parágrafo contém TANTO uma questão QUANTO alternativas, você deve identificar APENAS como questão/enunciado
3. Alternativas SEMPRE começam em uma nova linha/parágrafo
4. Um parágrafo só pode ter UM tipo de marcação (a mais apropriada)
5. Analise o INÍCIO de cada parágrafo para determinar seu tipo
6. SEMPRE retorne um JSON válido e completo
7. Para imagens, use o marcador apropriado se houver um estilo definido para imagens

EXEMPLOS DE PADRÕES COMUNS (apenas para referência):
- Questões frequentemente começam com números: "1.", "2)", "01.", "Questão 1"
- Alternativas começam com letras: "a)", "A)", "(a)", "A."
- Gabaritos podem conter: "Resposta:", "Gabarito:", "Alternativa correta:", letras isoladas com explicação
- Títulos de simulado: "SIMULADO", "Simulado 1", "SIMULADO 1"
- Elementos especiais: "[IMAGEM]", tabelas, gráficos

IMPORTANTE: Use APENAS os estilos definidos abaixo. Não invente marcadores.

ESTILOS A IDENTIFICAR:
"""
        
        for style in styles:
            prompt += f"\n- {style['name']}: {style['prompt']}"
            prompt += f"\n  Marcador a usar: {style['marker']}\n"
        
        if removal_prompts:
            prompt += "\nCONTEÚDO PARA MARCAR REMOÇÃO:\n"
            
            for removal in removal_prompts:
                prompt += f"\n- {removal['name']}: {removal['prompt']}"
                prompt += f"\n  Marcadores: {removal['startMarker']} (início) e {removal['endMarker']} (fim)\n"
        
        prompt += """
FORMATO DE RESPOSTA OBRIGATÓRIO:
Retorne APENAS um objeto JSON válido e COMPLETO, sem truncar.
O JSON deve ter EXATAMENTE este formato:
{"paragraphs": [{"index": 0, "markers": ["marcador1"]}, {"index": 1, "markers": []}, ...]}

IMPORTANTE: 
- Garanta que o JSON esteja COMPLETO e válido
- Cada parágrafo pode ter no MÁXIMO um marcador
- Se não tiver certeza, deixe sem marcação: "markers": []
- SEMPRE complete a resposta com JSON válido, mesmo que seja longo
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
        # Lista de marcadores válidos
        valid_markers = [style['marker'] for style in self.styles] if hasattr(self, 'styles') else []
        
        # Cria um mapa de resultados
        results_map = {}
        for p in ai_results.get('paragraphs', []):
            if isinstance(p, dict) and 'index' in p and 'markers' in p:
                # Filtra apenas marcadores válidos
                markers = p.get('markers', [])
                if valid_markers:
                    # Valida e corrige marcadores
                    validated_markers = []
                    for marker in markers:
                        if marker in valid_markers:
                            validated_markers.append(marker)
                        else:
                            print(f"    AVISO: Marcador inválido ignorado: {marker}")
                    markers = validated_markers
                
                results_map[p['index']] = markers
        
        # Aplica os marcadores aos parágrafos
        for para in batch:
            para['markers'] = results_map.get(para['index'], [])
            
        return batch