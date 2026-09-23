import React, { useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Key, 
  ExternalLink, 
  CheckCircle2, 
  AlertCircle, 
  Sparkles, 
  Copy, 
  X, 
  RefreshCw, 
  Check, 
  ShieldCheck, 
  ArrowRight,
  HelpCircle
} from 'lucide-react';
import { Button } from './Button';
import { testGroqKey } from '../lib/ai/aiService';

interface GeminiApiKeyTutorialModalProps {
  isOpen: boolean;
  onClose: () => void;
  onKeySaved: (key: string) => void;
  currentKey?: string;
  systemDefaultKey?: string;
}

export const GeminiApiKeyTutorialModal: React.FC<GeminiApiKeyTutorialModalProps> = ({
  isOpen,
  onClose,
  onKeySaved,
  currentKey = '',
  systemDefaultKey = ''
}) => {
  const [apiKeyInput, setApiKeyInput] = useState(currentKey);
  const [isKeyVisible, setIsKeyVisible] = useState(false);
  const [isTesting, setIsTesting] = useState(false);
  const [testResult, setTestResult] = useState<{ ok?: boolean; message?: string } | null>(null);
  const [copiedLink, setCopiedLink] = useState(false);

  const STUDIO_URL = 'https://aistudio.google.com/app/apikey';

  const handleCopyLink = () => {
    navigator.clipboard.writeText(STUDIO_URL);
    setCopiedLink(true);
    setTimeout(() => setCopiedLink(false), 2000);
  };

  const handleTestKey = async () => {
    const keyToTest = apiKeyInput.trim();
    if (!keyToTest) {
      setTestResult({ ok: false, message: 'Por favor, cole uma chave da API antes de testar.' });
      return;
    }

    setIsTesting(true);
    setTestResult(null);

    try {
      if (keyToTest.startsWith('gsk_')) {
        const result = await testGroqKey(keyToTest);
        if (result.ok) {
          setTestResult({
            ok: true,
            message: '✓ Conexão bem-sucedida com Groq! Sua chave está ativa e pronta para uso.'
          });
        } else {
          throw new Error(result.message);
        }
      } else {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${keyToTest}`;
        const res = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: "Responda apenas: 'OK'." }] }]
          })
        });

        const data = await res.json();
        if (!res.ok) {
          throw new Error(data.error?.message || `Erro HTTP ${res.status}`);
        }

        setTestResult({
          ok: true,
          message: '✓ Conexão bem-sucedida! Sua chave está ativa e pronta para uso.'
        });
      }
    } catch (err: any) {
      setTestResult({
        ok: false,
        message: 'Falha na validação: ' + (err?.message || 'Verifique se a chave foi copiada por completo.')
      });
    } finally {
      setIsTesting(false);
    }
  };

  const handleSave = () => {
    const key = apiKeyInput.trim();
    if (!key) {
      setTestResult({ ok: false, message: 'Insira uma chave válida antes de salvar.' });
      return;
    }
    localStorage.setItem('custom_gemini_api_key', key);
    onKeySaved(key);
    onClose();
  };

  // Volta a usar a chave do sistema, que fica guardada só no servidor.
  const handleUseSystemDefault = () => {
    setApiKeyInput('');
    localStorage.removeItem('custom_gemini_api_key');
    onKeySaved('');
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
      <motion.div 
        initial={{ scale: 0.95, opacity: 0, y: 10 }}
        animate={{ scale: 1, opacity: 1, y: 0 }}
        exit={{ scale: 0.95, opacity: 0, y: 10 }}
        transition={{ duration: 0.2 }}
        className="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-2xl w-full overflow-hidden flex flex-col max-h-[92vh]"
      >
        {/* Header com gradiente moderno */}
        <div className="px-6 py-5 bg-gradient-to-r from-indigo-600 via-indigo-700 to-sky-600 text-white flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-white/10 rounded-2xl backdrop-blur-md">
              <Sparkles size={22} className="text-amber-300 animate-pulse" />
            </div>
            <div>
              <h3 className="font-extrabold text-lg text-white tracking-tight">
                Como Criar e Configurar sua Chave da IA
              </h3>
              <p className="text-xs text-indigo-100 font-medium">
                Google Gemini • 100% Gratuito e sem cartão de crédito
              </p>
            </div>
          </div>
          <button 
            onClick={onClose}
            className="p-2 text-white/80 hover:text-white hover:bg-white/10 rounded-full transition-colors cursor-pointer"
          >
            <X size={20} />
          </button>
        </div>

        {/* Corpo com o tutorial didático */}
        <div className="p-6 overflow-y-auto space-y-6 flex-1 text-slate-700">
          
          {/* Mensagem de Boas-Vindas para Compartilhamento */}
          <div className="p-4 bg-indigo-50/70 border border-indigo-100 rounded-2xl flex items-start gap-3">
            <ShieldCheck className="text-indigo-600 shrink-0 mt-0.5" size={20} />
            <div className="text-xs text-indigo-950 leading-relaxed">
              <span className="font-bold">Por que configurar sua própria chave?</span>
              <p className="text-indigo-800 mt-0.5">
                Para que cada consultor tenha sua própria cota gratuita e independente de Inteligência Artificial sem sobrecarregar a equipe, o sistema permite cadastrar sua chave pessoal em menos de 1 minuto.
              </p>
            </div>
          </div>

          {/* Passo a Passo Didático */}
          <div className="space-y-3.5">
            <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
              <HelpCircle size={15} className="text-indigo-600" />
              Siga os 4 passos rápidos:
            </h4>

            {/* Passo 1 */}
            <div className="p-3.5 rounded-2xl border border-slate-200/80 bg-slate-50 flex items-start gap-3 hover:border-indigo-200 transition-colors">
              <div className="w-7 h-7 rounded-xl bg-indigo-600 text-white font-black text-xs flex items-center justify-center shrink-0">
                1
              </div>
              <div className="flex-1 text-xs">
                <p className="font-bold text-slate-800">Acesse o Google AI Studio</p>
                <p className="text-slate-500 mt-0.5">
                  Abra o portal oficial da Google onde as chaves gratuitas são emitidas.
                </p>
                <div className="mt-2 flex items-center gap-2 flex-wrap">
                  <a 
                    href={STUDIO_URL} 
                    target="_blank" 
                    rel="noreferrer"
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold text-[11px] shadow-sm transition-all active:scale-95"
                  >
                    <span>Abrir Google AI Studio</span>
                    <ExternalLink size={12} />
                  </a>
                  <button 
                    type="button" 
                    onClick={handleCopyLink}
                    className="inline-flex items-center gap-1 px-2.5 py-1.5 bg-white border border-slate-200 hover:bg-slate-100 text-slate-600 rounded-xl text-[11px] font-semibold transition-colors"
                  >
                    {copiedLink ? <Check size={12} className="text-emerald-600" /> : <Copy size={12} />}
                    <span>{copiedLink ? 'Link Copiado!' : 'Copiar Link'}</span>
                  </button>
                </div>
              </div>
            </div>

            {/* Passo 2 */}
            <div className="p-3.5 rounded-2xl border border-slate-200/80 bg-slate-50 flex items-start gap-3">
              <div className="w-7 h-7 rounded-xl bg-indigo-600 text-white font-black text-xs flex items-center justify-center shrink-0">
                2
              </div>
              <div className="flex-1 text-xs">
                <p className="font-bold text-slate-800">Faça login com sua Conta Google</p>
                <p className="text-slate-500 mt-0.5">
                  Utilize qualquer conta Gmail existente. É 100% gratuito e não exige dados de pagamento.
                </p>
              </div>
            </div>

            {/* Passo 3 */}
            <div className="p-3.5 rounded-2xl border border-slate-200/80 bg-slate-50 flex items-start gap-3">
              <div className="w-7 h-7 rounded-xl bg-indigo-600 text-white font-black text-xs flex items-center justify-center shrink-0">
                3
              </div>
              <div className="flex-1 text-xs">
                <p className="font-bold text-slate-800">Clique em "Create API key" (Criar chave de API)</p>
                <p className="text-slate-500 mt-0.5">
                  Selecione seu projeto (ou crie um novo padrão) e copie a chave gerada. Ela começará com <code className="bg-slate-200 px-1 py-0.5 rounded text-indigo-700 font-mono font-bold">AQ.</code> ou <code className="bg-slate-200 px-1 py-0.5 rounded text-indigo-700 font-mono font-bold">AIzaSy...</code>
                </p>
              </div>
            </div>

            {/* Passo 4 */}
            <div className="p-3.5 rounded-2xl border border-slate-200/80 bg-slate-50 flex items-start gap-3">
              <div className="w-7 h-7 rounded-xl bg-indigo-600 text-white font-black text-xs flex items-center justify-center shrink-0">
                4
              </div>
              <div className="flex-1 text-xs">
                <p className="font-bold text-slate-800">Cole a chave no campo abaixo</p>
                <p className="text-slate-500 mt-0.5">
                  Cole sua chave e clique em <b>Testar e Salvar</b>. Ela ficará salva exclusivamente na sua máquina.
                </p>
              </div>
            </div>
          </div>

          {/* Campo de Inserção da Chave */}
          <div className="pt-3 border-t border-slate-100 space-y-3">
            <label className="block text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center justify-between">
              <span className="flex items-center gap-1.5">
                <Key size={14} className="text-indigo-600" />
                Sua Chave da API do Google Gemini:
              </span>
              <button
                type="button"
                onClick={() => setIsKeyVisible(!isKeyVisible)}
                className="text-[11px] font-semibold text-indigo-600 hover:text-indigo-800 cursor-pointer"
              >
                {isKeyVisible ? 'Ocultar Chave' : 'Visualizar Chave'}
              </button>
            </label>

            <div className="flex flex-col sm:flex-row gap-2">
              <input
                type={isKeyVisible ? 'text' : 'password'}
                placeholder="Cole aqui sua chave (ex: gsk_..., AIzaSy..., ou AQ....)"
                value={apiKeyInput}
                onChange={(e) => {
                  setApiKeyInput(e.target.value);
                  setTestResult(null);
                }}
                className="flex-1 px-4 py-2.5 text-xs rounded-xl border border-slate-300 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 font-mono bg-white text-slate-800 shadow-sm"
              />
              <Button
                type="button"
                onClick={handleTestKey}
                disabled={isTesting || !apiKeyInput.trim()}
                className="bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs px-4 py-2.5 rounded-xl flex items-center gap-1.5 shrink-0 disabled:opacity-50"
              >
                <RefreshCw size={13} className={isTesting ? 'animate-spin' : ''} />
                <span>{isTesting ? 'Testando...' : 'Testar Chave'}</span>
              </Button>
            </div>

            {/* Feedback do Teste */}
            {testResult && (
              <div className={`p-3 rounded-xl text-xs flex items-start gap-2 animate-in fade-in ${
                testResult.ok 
                  ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' 
                  : 'bg-rose-50 text-rose-800 border border-rose-200'
              }`}>
                {testResult.ok ? (
                  <CheckCircle2 size={16} className="text-emerald-600 shrink-0 mt-0.5" />
                ) : (
                  <AlertCircle size={16} className="text-rose-600 shrink-0 mt-0.5" />
                )}
                <span className="leading-relaxed">{testResult.message}</span>
              </div>
            )}
          </div>
        </div>

        {/* Rodapé com ações */}
        <div className="px-6 py-4 bg-slate-50 flex flex-col sm:flex-row items-center justify-between gap-3 shrink-0 border-t border-slate-100">
          <button
            type="button"
            onClick={handleUseSystemDefault}
            className="text-[11px] font-bold text-indigo-600 hover:text-indigo-800 underline cursor-pointer"
          >
            Usar Chave do Sistema
          </button>
          
          <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
            <Button variant="secondary" onClick={onClose} className="text-xs">
              Fechar
            </Button>
            <Button 
              onClick={handleSave} 
              disabled={!apiKeyInput.trim()}
              className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-5 py-2.5 rounded-xl shadow-sm flex items-center gap-1.5"
            >
              <span>Salvar e Ativar IA</span>
              <ArrowRight size={14} />
            </Button>
          </div>
        </div>
      </motion.div>
    </div>
  );
};
