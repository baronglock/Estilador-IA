import json
import time
import requests
from typing import List, Dict
from backend.config import Config

class AIProcessor:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.model = "gpt-4.1"  # ou "gpt-3.5-turbo" se preferir
        self.api_url = "https://api.openai.com/v1/chat/completions"
        print(f"AIProcessor inicializado com modelo: {self.model}")
        
    def process_document(self, paragraphs: List[Dict], styles: List[Dict], removal_prompts: List[Dict]) -> Dict:
        """Processa documento com IA para identificar estilos"""
        marked_content = []
        processing_stats = {
            'total_paragraphs': len(paragraphs),
            'processed': 0,
            'marked': 0,
            'api_calls': 0
        }
        
        print(f"Iniciando processamento de {len(paragraphs)} parágrafos...")
        
        # Processa em lotes para otimizar
        batch_size = 50
        for i in range(0, len(paragraphs), batch_size):
            batch = paragraphs[i:i + batch_size]
            print(f"Processando batch {i//batch_size + 1} de {len(paragraphs)//batch_size + 1}")
            
            batch_results = self._process_batch(batch, styles, removal_prompts)
            marked_content.extend(batch_results)
            processing_stats['processed'] += len(batch)
            processing_stats['api_calls'] += 1
            
            # Pequena pausa para respeitar rate limits
            time.sleep(0.3)
        
        processing_stats['marked'] = sum(1 for p in marked_content if p.get('markers'))
        print(f"Processamento concluído. {processing_stats['marked']} parágrafos marcados.")
        
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
            "max_tokens": 4000
        }
        
        try:
            # Faz a requisição para a API
            response = requests.post(
                self.api_url, 
                headers=headers, 
                json=data,
                timeout=30  # timeout de 30 segundos
            )
            
            # Verifica se houve erro HTTP
            if response.status_code != 200:
                print(f"Erro na API: Status {response.status_code}")
                print(f"Resposta: {response.text}")
                return batch
            
            # Extrai o conteúdo da resposta
            result_data = response.json()
            content = result_data['choices'][0]['message']['content']
            
            # Tenta extrair JSON da resposta
            # Remove possível texto antes/depois do JSON
            json_start = content.find('{')
            json_end = content.rfind('}') + 1
            if json_start >= 0 and json_end > json_start:
                content = content[json_start:json_end]
            
            # Parseia a resposta
            result = json.loads(content)
            return self._merge_results(batch, result)
            
        except requests.exceptions.Timeout:
            print("Timeout na requisição à API")
            return batch
        except requests.exceptions.RequestException as e:
            print(f"Erro na requisição HTTP: {e}")
            return batch
        except json.JSONDecodeError as e:
            print(f"Erro ao parsear JSON da IA: {e}")
            print(f"Conteúdo recebido: {content[:200]}...")
            return batch
        except KeyError as e:
            print(f"Erro na estrutura da resposta: {e}")
            print(f"Resposta completa: {result_data}")
            return batch
        except Exception as e:
            print(f"Erro inesperado: {type(e).__name__}: {str(e)}")
            return batch
    
    def _build_system_prompt(self, styles: List[Dict], removal_prompts: List[Dict]) -> str:
        """Constrói o prompt do sistema com as definições de estilos"""
        prompt = """Você é um especialista em análise de documentos educacionais. 
Sua tarefa é identificar e marcar estilos em textos de simulados educacionais.

REGRAS IMPORTANTES:
1. Cada linha/parágrafo deve ser analisado INDEPENDENTEMENTE
2. Se um parágrafo contém TANTO uma questão QUANTO alternativas, você deve identificar APENAS como questão/enunciado
3. Alternativas SEMPRE começam em uma nova linha/parágrafo
4. Um parágrafo só pode ter UM tipo de marcação (a mais apropriada)
5. Analise o INÍCIO do cada parágrafo para determinar seu tipo

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
EXEMPLOS DE IDENTIFICAÇÃO:
- "3. Qual é a capital do Brasil?" → [[ENUNCIADO]]
- "3. Qual é a capital? A) São Paulo" → [[ENUNCIADO]] (NÃO marcar como alternativa)
- "A) São Paulo" → [[ALTERNATIVA]]
- "Resposta: letra A" → [[GABARITO]]

FORMATO DE RESPOSTA OBRIGATÓRIO:
Retorne APENAS um objeto JSON válido, sem nenhum texto adicional antes ou depois.
O JSON deve ter exatamente este formato:
{"paragraphs": [{"index": 0, "markers": ["marcador1"]}, {"index": 1, "markers": []}, ...]}

IMPORTANTE: 
- Cada parágrafo pode ter no MÁXIMO um marcador
- Se não tiver certeza, deixe sem marcação
- NÃO inclua explicações, apenas o JSON
"""
        
        return prompt
    
    def _build_user_prompt(self, batch: List[Dict]) -> str:
        """Constrói o prompt do usuário com o lote de parágrafos"""
        prompt = "Analise os seguintes parágrafos e retorne as marcações em formato JSON:\n\n"
        
        for para in batch:
            # Limita o tamanho do texto para não exceder tokens
            text = para['text'].strip()
            if len(text) > 200:
                text = text[:200] + "..."
            
            prompt += f"Parágrafo {para['index']}:\n{text}\n\n"
        
        prompt += "\nLembre-se: retorne APENAS o JSON no formato especificado, sem texto adicional."
        
        return prompt
    
    def _merge_results(self, batch: List[Dict], ai_results: Dict) -> List[Dict]:
        """Mescla os resultados da IA com o batch original"""
        # Cria um mapa de resultados
        results_map = {}
        for p in ai_results.get('paragraphs', []):
            if isinstance(p, dict) and 'index' in p and 'markers' in p:
                results_map[p['index']] = p.get('markers', [])
        
        # Aplica os marcadores aos parágrafos
        for para in batch:
            para['markers'] = results_map.get(para['index'], [])
            
        return batch