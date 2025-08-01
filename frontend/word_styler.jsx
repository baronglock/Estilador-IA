import React, { useState, useCallback } from 'react';
import { Upload, FileText, Settings, Play, Download, Clock, AlertCircle, CheckCircle, Loader2, Plus, Trash2, Edit2, Save } from 'lucide-react';

const WordAIStyler = () => {
  const [file, setFile] = useState(null);
  const [apiKey, setApiKey] = useState('');
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState({ step: '', percent: 0 });
  const [results, setResults] = useState(null);
  const [error, setError] = useState(null);
  const [bookName, setBookName] = useState('');
  
  // Estado para gerenciar estilos e prompts
  const [styles, setStyles] = useState([
    {
      id: 'titulo_simulado',
      name: 'Título do Simulado',
      wordStyle: 'Heading 1',
      marker: '[[TITULO_SIMULADO]]',
      prompt: 'Identifique títulos que contenham "Simulado" seguido de número (ex: Simulado 1, Simulado 2)',
      color: '#1e40af'
    },
    {
      id: 'subtitulo_estudos',
      name: 'Subtítulo Estudos',
      wordStyle: 'Heading 2',
      marker: '[[SUBTITULO_ESTUDOS]]',
      prompt: 'Identifique subtítulos que contenham "Estudos" seguido de intervalo (ex: Estudos 1 a 10)',
      color: '#7c3aed'
    },
    {
      id: 'enunciado',
      name: 'Enunciado',
      wordStyle: 'Question',
      marker: '[[ENUNCIADO]]',
      prompt: 'Identifique enunciados de questões que começam com número seguido de ponto ou parêntese',
      color: '#dc2626'
    },
    {
      id: 'alternativa',
      name: 'Alternativa',
      wordStyle: 'Alternative',
      marker: '[[ALTERNATIVA]]',
      prompt: 'Identifique alternativas que começam com letras (a), b), c), d), e) ou A), B), C), D), E)',
      color: '#ea580c'
    },
    {
      id: 'gabarito',
      name: 'Gabarito',
      wordStyle: 'Answer',
      marker: '[[GABARITO]]',
      prompt: 'Identifique gabaritos que contenham palavras como "Resposta:", "Gabarito:" ou "Alternativa correta:"',
      color: '#16a34a'
    }
  ]);

  const [removalPrompts, setRemovalPrompts] = useState([
    {
      id: 'intro',
      name: 'Introdução/Capa',
      startMarker: '[[REMOVE_INTRO_START]]',
      endMarker: '[[REMOVE_INTRO_END]]',
      prompt: 'Marque o início e fim de conteúdo introdutório, capas, prefácios antes do primeiro simulado'
    },
    {
      id: 'cartao',
      name: 'Cartão Resposta',
      startMarker: '[[REMOVE_CARTAO_START]]',
      endMarker: '[[REMOVE_CARTAO_END]]',
      prompt: 'Marque o início e fim de páginas de cartão resposta ou folhas de resposta'
    }
  ]);

  const [editingStyle, setEditingStyle] = useState(null);
  const [editingRemoval, setEditingRemoval] = useState(null);

  // Funções para gerenciar estilos
  const addStyle = () => {
    const newStyle = {
      id: `style_${Date.now()}`,
      name: 'Novo Estilo',
      wordStyle: 'Normal',
      marker: `[[NOVO_${Date.now()}]]`,
      prompt: 'Defina o prompt para identificar este estilo',
      color: '#6b7280'
    };
    setStyles([...styles, newStyle]);
    setEditingStyle(newStyle.id);
  };

  const updateStyle = (id, field, value) => {
    setStyles(styles.map(style => 
      style.id === id ? { ...style, [field]: value } : style
    ));
  };

  const deleteStyle = (id) => {
    setStyles(styles.filter(style => style.id !== id));
  };

  // Funções para gerenciar prompts de remoção
  const addRemovalPrompt = () => {
    const newRemoval = {
      id: `removal_${Date.now()}`,
      name: 'Nova Remoção',
      startMarker: `[[REMOVE_${Date.now()}_START]]`,
      endMarker: `[[REMOVE_${Date.now()}_END]]`,
      prompt: 'Defina o prompt para identificar o que remover'
    };
    setRemovalPrompts([...removalPrompts, newRemoval]);
    setEditingRemoval(newRemoval.id);
  };

  const updateRemovalPrompt = (id, field, value) => {
    setRemovalPrompts(removalPrompts.map(removal => 
      removal.id === id ? { ...removal, [field]: value } : removal
    ));
  };

  const deleteRemovalPrompt = (id) => {
    setRemovalPrompts(removalPrompts.filter(removal => removal.id !== id));
  };

  // Função para processar o arquivo
  const processFile = async () => {
    if (!file || !apiKey || !bookName) {
      setError('Por favor, preencha todos os campos obrigatórios');
      return;
    }

    setProcessing(true);
    setError(null);
    setProgress({ step: 'Iniciando processamento...', percent: 0 });

    try {
      // Simular processamento (substituir pela implementação real)
      const steps = [
        { step: 'Lendo documento...', percent: 10 },
        { step: 'Analisando estrutura com IA...', percent: 30 },
        { step: 'Aplicando marcações...', percent: 50 },
        { step: 'Aplicando estilos...', percent: 70 },
        { step: 'Gerando documentos finais...', percent: 90 },
        { step: 'Processamento concluído!', percent: 100 }
      ];

      for (const step of steps) {
        setProgress(step);
        await new Promise(resolve => setTimeout(resolve, 1000));
      }

      // Simular resultados
      setResults({
        processTime: '2m 34s',
        totalPages: 120,
        questionsProcessed: 100,
        files: [
          { name: `${bookName}_completo.docx`, type: 'complete', size: '2.4MB' },
          { name: `${bookName}_simulado1_questoes.docx`, type: 'questions', size: '580KB' },
          { name: `${bookName}_simulado1_gabarito.docx`, type: 'answers', size: '340KB' },
          { name: `${bookName}_simulado2_questoes.docx`, type: 'questions', size: '590KB' },
          { name: `${bookName}_simulado2_gabarito.docx`, type: 'answers', size: '350KB' },
          { name: `${bookName}_simulado3_questoes.docx`, type: 'questions', size: '600KB' },
          { name: `${bookName}_simulado3_gabarito.docx`, type: 'answers', size: '360KB' },
          { name: `${bookName}_simulado4_questoes.docx`, type: 'questions', size: '610KB' },
          { name: `${bookName}_simulado4_gabarito.docx`, type: 'answers', size: '370KB' }
        ]
      });

    } catch (err) {
      setError('Erro ao processar arquivo: ' + err.message);
    } finally {
      setProcessing(false);
    }
  };

  const handleFileUpload = (e) => {
    const uploadedFile = e.target.files[0];
    if (uploadedFile && uploadedFile.name.endsWith('.docx')) {
      setFile(uploadedFile);
      setError(null);
    } else {
      setError('Por favor, selecione um arquivo .docx válido');
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        <h1 className="text-3xl font-bold text-gray-900 mb-8">
          Processador de Estilos Word com IA
        </h1>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Coluna de Configuração */}
          <div className="lg:col-span-2 space-y-6">
            {/* Upload e Configurações Básicas */}
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-xl font-semibold mb-4 flex items-center gap-2">
                <Settings className="w-5 h-5" />
                Configurações
              </h2>

              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Arquivo Word (.docx)
                  </label>
                  <div className="relative">
                    <input
                      type="file"
                      accept=".docx"
                      onChange={handleFileUpload}
                      className="hidden"
                      id="file-upload"
                    />
                    <label
                      htmlFor="file-upload"
                      className="flex items-center justify-center gap-2 p-4 border-2 border-dashed border-gray-300 rounded-lg cursor-pointer hover:border-gray-400 transition-colors"
                    >
                      <Upload className="w-5 h-5 text-gray-400" />
                      <span className="text-gray-600">
                        {file ? file.name : 'Clique para selecionar arquivo'}
                      </span>
                    </label>
                  </div>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Nome do Livro
                  </label>
                  <input
                    type="text"
                    value={bookName}
                    onChange={(e) => setBookName(e.target.value)}
                    placeholder="Ex: Simulado_Matematica_2024"
                    className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    API Key (GPT-4.1)
                  </label>
                  <input
                    type="password"
                    value={apiKey}
                    onChange={(e) => setApiKey(e.target.value)}
                    placeholder="sk-..."
                    className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  />
                </div>
              </div>
            </div>

            {/* Gerenciador de Estilos */}
            <div className="bg-white rounded-lg shadow p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xl font-semibold flex items-center gap-2">
                  <FileText className="w-5 h-5" />
                  Estilos e Prompts
                </h2>
                <button
                  onClick={addStyle}
                  className="flex items-center gap-1 px-3 py-1 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                >
                  <Plus className="w-4 h-4" />
                  Adicionar
                </button>
              </div>

              <div className="space-y-3">
                {styles.map((style) => (
                  <div
                    key={style.id}
                    className="border border-gray-200 rounded-lg p-4 space-y-3"
                    style={{ borderLeftWidth: '4px', borderLeftColor: style.color }}
                  >
                    {editingStyle === style.id ? (
                      <div className="space-y-3">
                        <input
                          type="text"
                          value={style.name}
                          onChange={(e) => updateStyle(style.id, 'name', e.target.value)}
                          className="w-full px-2 py-1 border border-gray-300 rounded text-sm"
                          placeholder="Nome do estilo"
                        />
                        <input
                          type="text"
                          value={style.wordStyle}
                          onChange={(e) => updateStyle(style.id, 'wordStyle', e.target.value)}
                          className="w-full px-2 py-1 border border-gray-300 rounded text-sm"
                          placeholder="Estilo no Word"
                        />
                        <input
                          type="text"
                          value={style.marker}
                          onChange={(e) => updateStyle(style.id, 'marker', e.target.value)}
                          className="w-full px-2 py-1 border border-gray-300 rounded text-sm"
                          placeholder="Marcador"
                        />
                        <textarea
                          value={style.prompt}
                          onChange={(e) => updateStyle(style.id, 'prompt', e.target.value)}
                          className="w-full px-2 py-1 border border-gray-300 rounded text-sm h-20"
                          placeholder="Prompt para identificação"
                        />
                        <div className="flex gap-2">
                          <button
                            onClick={() => setEditingStyle(null)}
                            className="px-3 py-1 bg-green-600 text-white rounded text-sm hover:bg-green-700"
                          >
                            <Save className="w-4 h-4" />
                          </button>
                          <input
                            type="color"
                            value={style.color}
                            onChange={(e) => updateStyle(style.id, 'color', e.target.value)}
                            className="w-8 h-8 rounded cursor-pointer"
                          />
                        </div>
                      </div>
                    ) : (
                      <>
                        <div className="flex items-start justify-between">
                          <div className="flex-1">
                            <h3 className="font-semibold text-gray-900">{style.name}</h3>
                            <p className="text-sm text-gray-500">Word: {style.wordStyle}</p>
                            <p className="text-xs text-gray-400 mt-1">Marcador: {style.marker}</p>
                          </div>
                          <div className="flex gap-1">
                            <button
                              onClick={() => setEditingStyle(style.id)}
                              className="p-1 text-gray-500 hover:text-blue-600"
                            >
                              <Edit2 className="w-4 h-4" />
                            </button>
                            <button
                              onClick={() => deleteStyle(style.id)}
                              className="p-1 text-gray-500 hover:text-red-600"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </div>
                        <p className="text-sm text-gray-600">{style.prompt}</p>
                      </>
                    )}
                  </div>
                ))}
              </div>
            </div>

            {/* Gerenciador de Remoções */}
            <div className="bg-white rounded-lg shadow p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xl font-semibold flex items-center gap-2">
                  <Trash2 className="w-5 h-5" />
                  Conteúdo para Remover
                </h2>
                <button
                  onClick={addRemovalPrompt}
                  className="flex items-center gap-1 px-3 py-1 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors text-sm"
                >
                  <Plus className="w-4 h-4" />
                  Adicionar
                </button>
              </div>

              <div className="space-y-3">
                {removalPrompts.map((removal) => (
                  <div key={removal.id} className="border border-gray-200 rounded-lg p-4 space-y-3">
                    {editingRemoval === removal.id ? (
                      <div className="space-y-3">
                        <input
                          type="text"
                          value={removal.name}
                          onChange={(e) => updateRemovalPrompt(removal.id, 'name', e.target.value)}
                          className="w-full px-2 py-1 border border-gray-300 rounded text-sm"
                          placeholder="Nome da remoção"
                        />
                        <textarea
                          value={removal.prompt}
                          onChange={(e) => updateRemovalPrompt(removal.id, 'prompt', e.target.value)}
                          className="w-full px-2 py-1 border border-gray-300 rounded text-sm h-20"
                          placeholder="Prompt para identificação"
                        />
                        <button
                          onClick={() => setEditingRemoval(null)}
                          className="px-3 py-1 bg-green-600 text-white rounded text-sm hover:bg-green-700"
                        >
                          <Save className="w-4 h-4" />
                        </button>
                      </div>
                    ) : (
                      <>
                        <div className="flex items-start justify-between">
                          <div className="flex-1">
                            <h3 className="font-semibold text-gray-900">{removal.name}</h3>
                            <p className="text-xs text-gray-400 mt-1">
                              {removal.startMarker} ... {removal.endMarker}
                            </p>
                          </div>
                          <div className="flex gap-1">
                            <button
                              onClick={() => setEditingRemoval(removal.id)}
                              className="p-1 text-gray-500 hover:text-blue-600"
                            >
                              <Edit2 className="w-4 h-4" />
                            </button>
                            <button
                              onClick={() => deleteRemovalPrompt(removal.id)}
                              className="p-1 text-gray-500 hover:text-red-600"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </div>
                        <p className="text-sm text-gray-600">{removal.prompt}</p>
                      </>
                    )}
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Coluna de Status e Resultados */}
          <div className="space-y-6">
            {/* Botão de Processar */}
            <div className="bg-white rounded-lg shadow p-6">
              <button
                onClick={processFile}
                disabled={processing || !file || !apiKey || !bookName}
                className="w-full flex items-center justify-center gap-2 px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition-colors"
              >
                {processing ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    Processando...
                  </>
                ) : (
                  <>
                    <Play className="w-5 h-5" />
                    Processar Documento
                  </>
                )}
              </button>
            </div>

            {/* Status de Processamento */}
            {processing && (
              <div className="bg-white rounded-lg shadow p-6">
                <h3 className="text-lg font-semibold mb-4 flex items-center gap-2">
                  <Clock className="w-5 h-5" />
                  Status
                </h3>
                <div className="space-y-3">
                  <p className="text-sm text-gray-600">{progress.step}</p>
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div
                      className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                      style={{ width: `${progress.percent}%` }}
                    />
                  </div>
                  <p className="text-xs text-gray-500">{progress.percent}% concluído</p>
                </div>
              </div>
            )}

            {/* Mensagens de Erro */}
            {error && (
              <div className="bg-red-50 border border-red-200 rounded-lg p-4">
                <div className="flex items-start gap-2">
                  <AlertCircle className="w-5 h-5 text-red-600 flex-shrink-0 mt-0.5" />
                  <p className="text-sm text-red-700">{error}</p>
                </div>
              </div>
            )}

            {/* Resultados */}
            {results && (
              <div className="bg-white rounded-lg shadow p-6">
                <h3 className="text-lg font-semibold mb-4 flex items-center gap-2">
                  <CheckCircle className="w-5 h-5 text-green-600" />
                  Processamento Concluído
                </h3>
                
                <div className="space-y-4">
                  <div className="grid grid-cols-2 gap-3 text-sm">
                    <div>
                      <p className="text-gray-500">Tempo total</p>
                      <p className="font-semibold">{results.processTime}</p>
                    </div>
                    <div>
                      <p className="text-gray-500">Páginas</p>
                      <p className="font-semibold">{results.totalPages}</p>
                    </div>
                    <div>
                      <p className="text-gray-500">Questões</p>
                      <p className="font-semibold">{results.questionsProcessed}</p>
                    </div>
                    <div>
                      <p className="text-gray-500">Arquivos</p>
                      <p className="font-semibold">{results.files.length}</p>
                    </div>
                  </div>

                  <div className="border-t pt-4">
                    <h4 className="font-medium text-gray-900 mb-3">Arquivos Gerados</h4>
                    <div className="space-y-2">
                      {results.files.map((file, index) => (
                        <div
                          key={index}
                          className="flex items-center justify-between p-2 bg-gray-50 rounded-lg"
                        >
                          <div className="flex items-center gap-2">
                            <FileText className="w-4 h-4 text-gray-500" />
                            <span className="text-sm text-gray-700">{file.name}</span>
                          </div>
                          <div className="flex items-center gap-2">
                            <span className="text-xs text-gray-500">{file.size}</span>
                            <button className="p-1 text-blue-600 hover:text-blue-700">
                              <Download className="w-4 h-4" />
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <button className="w-full mt-4 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors flex items-center justify-center gap-2">
                    <Download className="w-5 h-5" />
                    Baixar Todos os Arquivos
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default WordAIStyler;