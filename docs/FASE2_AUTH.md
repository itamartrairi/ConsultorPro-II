# Fase 2 — AuthProvider + LoginView

## Objetivo

Tirar autenticação e tela de login do `App.tsx` sem mudar o comportamento.

## Arquivos

```
src/features/auth/
├── hooks/useAuth.ts          # AuthProvider + useAuth()
└── components/LoginView.tsx  # UI de login/registro/reset
```

## Integração no `main.tsx` / `App.tsx`

### 1. Envolver a árvore com AuthProvider

Em `src/main.tsx` (ou no topo de `App`):

```tsx
import { AuthProvider } from './features/auth/hooks/useAuth';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <HashRouter>
      <AuthProvider>
        <App />
      </AuthProvider>
    </HashRouter>
  </StrictMode>
);
```

### 2. No App.tsx

```tsx
import { useAuth } from './features/auth/hooks/useAuth';
import { LoginView } from './features/auth/components/LoginView';

const { user, loading, logout } = useAuth();

if (loading) return <Loader…>;
if (!user) return <LoginView />;
// resto do app logado…
```

Remover do App: estados de auth, onAuthStateChanged manual, handlers de login e JSX da tela de login.

### 3. Critério de saída

- [ ] Login Google / e-mail / registro / reset / logout
- [ ] App logado nas mesmas views
