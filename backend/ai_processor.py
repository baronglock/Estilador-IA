import time
import requests
from typing import List, Dict
from backend.config import Config
import re

class AIProcessor:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.model = "gpt-4.1"
        self.api_url = "https://api.openai.com/v1/chat/completions"
        print(f"AIProcessor inicializado com modelo: {self.model}")
        
    def process_document(self, paragraphs, styles, removal_prompts):
        """Processa com contexto contínuo"""
        chunk_size = 50
        context_size = 10
        all_results = []
        
        i = 0
        while i < len(paragraphs):
            # Pega chunk com contexto anterior
            start = max(0, i - context_size)
            end = min(i + chunk_size, len(paragraphs))
            
            chunk = paragraphs[start:end]
            print(f"Processando parágrafos {start+1} a {end} (com contexto)")
            
            # Processa
            marked_chunk = self._process_chunk(
                chunk, 
                styles, 
                removal_prompts,
                is_continuation=(i > 0)
            )
            
            # Adiciona apenas os novos (ignora os de contexto)
            if i > 0:
                new_results = marked_chunk[context_size:]
            else:
                new_results = marked_chunk[:chunk_size]
            
            all_results.extend(new_results)
            
            # Avança
            i += chunk_size
            time.sleep(0.2)
        
        return {
            'marked_content': all_results,
            'stats': {...}
        }
    
    def _smart_chunk_split(self, paragraphs, target_size=50):
        """Quebra em pontos naturais, não no meio de questões"""
        chunks = []
        current_chunk = []
        
        for para in paragraphs:
            current_chunk.append(para)
            
            # Se atingiu tamanho E encontrou ponto natural
            if len(current_chunk) >= target_size:
                # Verifica se é bom ponto para quebrar
                text = para.get('text', '').strip()
                
                # Quebra após gabarito completo ou antes de nova questão
                if (
                    any(word in text.lower() for word in ['resposta:', 'gabarito:']) or
                    (len(current_chunk) > target_size + 10 and 
                    re.match(r'^\d+[\.\)]\s', text))
                ):
                    chunks.append(current_chunk)
                    current_chunk = []
        
        if current_chunk:
            chunks.append(current_chunk)
        
        return chunks
    
    def _process_chunk(self, chunk, styles, removal_prompts, previous_ending=None):
    
        system_message = self._build_prompt(styles, removal_prompts)
        
        # Adiciona contexto do chunk anterior
        if previous_ending:
            user_message = f"""
    CONTEXTO: O chunk anterior terminou com:
    {previous_ending}

    Agora continue processando:
    {text_to_process}
    """
        for para in chunk:
            if para.get('type') == 'image':
                text_to_process += f"[{para['index']}]IMAGEM_AQUI[/{para['index']}]\n\n"
            else:
                text_to_process += f"[{para['index']}]{para['text']}[/{para['index']}]\n\n"
        
        # Monta o prompt
        system_message = self._build_prompt(styles, removal_prompts)
        user_message = f"Leia e marque o texto abaixo:\n\n{text_to_process}"
        
        # Chama a API
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_message},
                {"role": "user", "content": user_message}
            ],
            "temperature": 0.1,
            "max_tokens": 8000
        }
        
        try:
            response = requests.post(self.api_url, headers=headers, json=data, timeout=60)
            
            if response.status_code == 200:
                result = response.json()
                marked_text = result['choices'][0]['message']['content']
                
                # Extrai as marcações
                return self._extract_marks(chunk, marked_text)
            else:
                print(f"Erro API: {response.status_code}")
                return chunk
                
        except Exception as e:
            print(f"Erro: {e}")
            return chunk
    
    def _build_prompt(self, styles: List[Dict], removal_prompts: List[Dict]) -> str:
        """Constrói prompt simples e direto"""
        prompt = """Você vai ler um documento e adicionar marcações de estilo.

REGRA PRINCIPAL: Adicione a marcação NO INÍCIO do parágrafo, mantendo TODO o texto original.

ESTILOS DISPONÍVEIS:
"""
        
        # Lista todos os estilos com seus prompts
        for i, style in enumerate(styles):
            prompt += f"\n{i+1}. {style['name']}"
            prompt += f"\n   Marcador: {style['marker']}"
            if style.get('elementType') == 'text':
                prompt += f"\n   Identificar: {style['prompt']}"
            elif style.get('elementType') == 'image':
                prompt += f"\n   Identificar: Quando encontrar IMAGEM_AQUI"
            prompt += "\n"
        
        # Adiciona transições se existirem
        if styles and 'transitions' in styles[0]:
            prompt += "\nSEQUÊNCIA:\n"
            for style in styles:
                if 'transitions' in style:
                    for t in style['transitions']:
                        if t['type'] == 'required':
                            prompt += f"- Após {styles[t['from']]['name']} SEMPRE vem {styles[t['to']]['name']}\n"
        
        # Marcações de remoção
        if removal_prompts:
            prompt += "\nREMOÇÕES:\n"
            for rem in removal_prompts:
                prompt += f"- {rem['name']}: {rem['prompt']}\n"
                prompt += f"  Marcar início: {rem['startMarker']}, fim: {rem['endMarker']}\n"
        
        prompt += """
EXEMPLO:
[1]Qual é a capital do Brasil?[/1]
FICA:
[1][[ENUNCIADO]]Qual é a capital do Brasil?[/1]

IMPORTANTE: Leia, entenda o contexto e marque corretamente usando os prompts acima."""
        
        return prompt
    
    def _extract_marks(self, original_chunk: List[Dict], marked_text: str) -> List[Dict]:
        """Extrai as marcações do texto retornado"""
        results = []
        
        for para in original_chunk:
            para_copy = para.copy()
            para_copy['markers'] = []
            
            # Busca o parágrafo marcado
            import re
            pattern = rf'\[{para["index"]}\](.*?)\[/{para["index"]}\]'
            match = re.search(pattern, marked_text, re.DOTALL)
            
            if match:
                content = match.group(1)
                # Extrai marcador se houver
                marker_match = re.match(r'(\[\[[A-Z_]+\]\])', content)
                if marker_match:
                    para_copy['markers'] = [marker_match.group(1)]
            
            results.append(para_copy)
        
        return results