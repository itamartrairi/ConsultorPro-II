---
name: video-tutorial-speech-sync
description: >-
  Guia e padrões de implementação para tutoriais audiovisuais interativos com Web SpeechSynthesis API
  e slides animados. Use quando criar, corrigir ou otimizar componentes de vídeo tutorial, cards de
  apresentação com narração por voz (TTS), prevenindo cortes prematuros de áudio, descompasso de texto
  e bugs de suspensão do Chromium.
---

# Video Tutorial & Speech Synthesis Synchronization

Esta skill documenta o padrão arquitetural e de implementação para players tutoriais e apresentações com narração sintetizada no navegador (`window.speechSynthesis`), eliminando cortes de áudio e descompasso com os textos dos quadros.

## Problemas Recorrentes em Players de Tutorial com TTS

1. **Corte por Temporizador Estático**:
   - Avançar slides com base apenas em tempo fixo (ex: `setInterval` de `duration * 1000`) corta a narração antes do término quando o navegador usa uma voz mais pausada.
2. **Bug de Suspensão do Chromium (>14s)**:
   - Em navegadores baseados em Chromium (Chrome, Edge, Opera, Brave), `window.speechSynthesis` sofre garbage collection silencioso ou pausa involuntária em falas com mais de 14 segundos.
3. **Truncamento de Legendas e Orientações**:
   - Containers de transcrição com altura fixa ou regras de overflow cortam as orientações da cena visualmente.
4. **Vozes Carregadas Assincronamente**:
   - `window.speechSynthesis.getVoices()` retorna vazio na primeira chamada síncrona em vários navegadores se não houver listener de `onvoiceschanged`.

---

## Padrão de Arquitetura de Sincronização

### 1. Inicialização e Carregamento de Vozes
```typescript
useEffect(() => {
  if (typeof window !== 'undefined' && 'speechSynthesis' in window) {
    setHasVoiceSupport(true);
    // Pré-carrega vozes e escuta mudanças assíncronas
    window.speechSynthesis.getVoices();
    const onVoicesChanged = () => {
      window.speechSynthesis.getVoices();
    };
    window.speechSynthesis.onvoiceschanged = onVoicesChanged;

    return () => {
      if (window.speechSynthesis) {
        window.speechSynthesis.onvoiceschanged = null;
      }
    };
  } else {
    setHasVoiceSupport(false);
  }
}, []);
```

### 2. Ciclo de Vida da Narração com Keepalive
```typescript
const keepAliveTimerRef = useRef<NodeJS.Timeout | null>(null);
const speechFinishedRef = useRef(false);

const clearKeepAlive = () => {
  if (keepAliveTimerRef.current) {
    clearInterval(keepAliveTimerRef.current);
    keepAliveTimerRef.current = null;
  }
};

const stopSpeech = () => {
  clearKeepAlive();
  if (typeof window !== 'undefined' && 'speechSynthesis' in window) {
    window.speechSynthesis.cancel();
    setSpeechActive(false);
  }
};

const playVoiceForStep = (stepItem: TutorialStep) => {
  stopSpeech();
  speechFinishedRef.current = false;

  if (isMuted || !hasVoiceSupport || typeof window === 'undefined') {
    speechFinishedRef.current = true;
    return;
  }

  try {
    const utter = new SpeechSynthesisUtterance(stepItem.narration);
    utter.lang = 'pt-BR';
    utter.rate = 1.0; // cadência clara e articulada
    utter.pitch = 1.08;

    const voices = window.speechSynthesis.getVoices();
    const voice = voices.find(v => (v.lang.includes('pt') || v.lang.includes('PT')) && /maria|francisca|leticia|luciana|female|google/i.test(v.name))
      || voices.find(v => v.lang.includes('pt') || v.lang.includes('PT'));

    if (voice) utter.voice = voice;

    utter.onstart = () => {
      setSpeechActive(true);
      speechFinishedRef.current = false;
    };

    utter.onend = () => {
      setSpeechActive(false);
      speechFinishedRef.current = true;
      clearKeepAlive();
    };

    utter.onerror = () => {
      setSpeechActive(false);
      speechFinishedRef.current = true;
      clearKeepAlive();
    };

    window.speechSynthesis.speak(utter);

    // Keepalive anti-corte do Chromium a cada 10s
    clearKeepAlive();
    keepAliveTimerRef.current = setInterval(() => {
      if (typeof window !== 'undefined' && 'speechSynthesis' in window && window.speechSynthesis.speaking) {
        window.speechSynthesis.pause();
        window.speechSynthesis.resume();
      } else {
        clearKeepAlive();
      }
    }, 10000);
  } catch (e) {
    console.warn("TTS não iniciado:", e);
    setSpeechActive(false);
    speechFinishedRef.current = true;
  }
};
```

### 3. Progresso do Slide e Trava de Transição (Hold on Speech)
Nunca avance o slide enquanto a narração estiver ativa:
```typescript
useEffect(() => {
  if (!isPlaying) {
    if (timerRef.current) clearInterval(timerRef.current);
    return;
  }

  const intervalTime = 100;
  const totalDurationMs = Math.max(step.duration * 1000, 20000);
  const stepIncrement = (intervalTime / totalDurationMs) * 100;

  timerRef.current = setInterval(() => {
    setProgress(prev => {
      const isSpeaking = typeof window !== 'undefined' && 'speechSynthesis' in window && window.speechSynthesis.speaking;
      const voiceStillBusy = !isMuted && hasVoiceSupport && (isSpeaking || !speechFinishedRef.current);

      // Trava no final da barra até a fala terminar
      if (prev >= 98 && voiceStillBusy) {
        return 98;
      }

      if (prev >= 100) {
        if (voiceStillBusy) return 99;

        if (currentStepIndex < TUTORIAL_STEPS.length - 1) {
          setCurrentStepIndex(currentStepIndex + 1);
          return 0;
        } else {
          setIsPlaying(false);
          stopSpeech();
          return 100;
        }
      }
      return prev + stepIncrement;
    });
  }, intervalTime);

  return () => {
    if (timerRef.current) clearInterval(timerRef.current);
  };
}, [isPlaying, currentStepIndex, step.duration, isMuted, hasVoiceSupport]);
```

### 4. Apresentação do Texto das Orientações
- Use `select-text break-words leading-relaxed` para garantir legibilidade de todas as orientações.
- Exiba feedback visual quando a locução estiver ativa (`speechActive`), indicando explicitamente se a cena está narrando ou se a transcrição está em exibição estática.
- Garanta que a caixa de transcrição não use alturas fixas restritivas que causem `overflow: hidden`.
